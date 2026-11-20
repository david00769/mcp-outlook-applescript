import { existsSync, writeFileSync, unlinkSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { randomBytes } from 'node:crypto';
import { executeAppleScriptOrThrow, assertBridgeHealthyForWrite, AppleScriptExecutionError } from './executor.js';
import * as scripts from './scripts.js';
import { parseSendEmailResult } from './parser.js';
import { assertSafeAttachmentPath } from '../utils/paths.js';
import { logAudit } from '../utils/audit.js';
import { AppleScriptError, AttachmentNotFoundError, MailSendError, MailSendIndeterminateError } from '../utils/errors.js';
/**
 * Node-level timeout for the send osascript. MUST be larger than the inner
 * AppleScript `with timeout` (45s in scripts.sendEmail) so the AppleEvent times
 * out cleanly (-1712) instead of Node SIGKILLing osascript mid-send — the
 * latter corrupts the bridge for all subsequent calls.
 */
const SEND_NODE_TIMEOUT_MS = 60_000;
/** Sends email through Outlook via AppleScript. */
/**
 * Validates every attachment and inline-image path (sensitive-file denylist or
 * the OUTLOOK_MCP_ATTACHMENT_DIR allowlist, after resolving symlinks) and that
 * it exists, and returns the params carrying the resolved paths, so the file
 * attached is the file that was checked.
 * @throws ValidationError if a path is disallowed.
 * @throws AttachmentNotFoundError if a path does not exist.
 */
function withCheckedFiles(params) {
    const check = (path) => {
        const safe = assertSafeAttachmentPath(path);
        if (!existsSync(safe)) {
            throw new AttachmentNotFoundError(path);
        }
        return safe;
    };
    return {
        ...params,
        ...(params.attachments != null && {
            attachments: params.attachments.map((a) => ({ ...a, path: check(a.path) })),
        }),
        ...(params.inlineImages != null && {
            inlineImages: params.inlineImages.map((i) => ({ ...i, path: check(i.path) })),
        }),
    };
}
/** Recipient and file counts for an audit line; never addresses or paths. */
function auditCounts(params) {
    return {
        to: params.to.length,
        cc: params.cc?.length ?? 0,
        bcc: params.bcc?.length ?? 0,
        attachments: (params.attachments?.length ?? 0) + (params.inlineImages?.length ?? 0),
    };
}
export class AppleScriptMailSender {
    /**
     * Validates attachment paths, builds the send-email AppleScript, executes
     * it, and parses the result.
     * @param params - Email composition parameters (recipients, body, attachments, etc.).
     * @returns The sent message's ID and timestamp.
     * @throws ValidationError if an attachment or inline image path is disallowed.
     * @throws AttachmentNotFoundError if any attachment or inline image path does not exist.
     * @throws MailSendError if Outlook reports a send failure.
     */
    sendEmail(params) {
        params = withCheckedFiles(params);
        // Pre-send health gate: refuse with guidance if the bridge is degraded,
        // BEFORE composing anything. Never fire a send into a stuck bridge — it
        // hangs and gets SIGKILLed mid-AppleEvent, corrupting the bridge.
        assertBridgeHealthyForWrite('send_email');
        // For HTML bodies, write to a temp file so AppleScript can read it
        // instead of embedding HTML in the AppleScript string literal (which
        // breaks the parser on quotes, colons, and other special chars).
        let bodyFilePath;
        if (params.bodyType === 'html') {
            const id = randomBytes(8).toString('hex');
            bodyFilePath = join(tmpdir(), `mcp-outlook-html-${id}.html`);
            writeFileSync(bodyFilePath, params.body, 'utf8');
        }
        try {
            let scriptParams = {
                to: params.to,
                subject: params.subject,
                body: params.body,
                bodyType: params.bodyType,
                ...(bodyFilePath != null && { bodyFilePath }),
            };
            if (params.cc != null)
                scriptParams = { ...scriptParams, cc: params.cc };
            if (params.bcc != null)
                scriptParams = { ...scriptParams, bcc: params.bcc };
            if (params.replyTo != null)
                scriptParams = { ...scriptParams, replyTo: params.replyTo };
            if (params.attachments != null)
                scriptParams = { ...scriptParams, attachments: params.attachments };
            if (params.inlineImages != null)
                scriptParams = { ...scriptParams, inlineImages: params.inlineImages };
            if (params.accountId != null)
                scriptParams = { ...scriptParams, accountId: params.accountId };
            const script = scripts.sendEmail(scriptParams);
            let output;
            try {
                output = executeAppleScriptOrThrow(script, { timeoutMs: SEND_NODE_TIMEOUT_MS });
            }
            catch (err) {
                // Node killed osascript at 60s — even the 45s inner timeout
                // didn't fire. Outcome indeterminate; do not let callers retry.
                if (err instanceof AppleScriptExecutionError && err.errorType === 'timeout') {
                    throw new MailSendIndeterminateError('osascript was killed at the Node 60s timeout');
                }
                throw err;
            }
            const result = parseSendEmailResult(output);
            if (result == null) {
                throw new AppleScriptError('Failed to parse send email response');
            }
            if (!result.success) {
                const reason = result.error ?? 'Unknown error';
                // Inner `with timeout` fired (-1712): the send was issued but its
                // reply timed out, so it may or may not have queued. Signal
                // INDETERMINATE so the caller verifies rather than blind-retrying.
                if (/-1712|AppleEvent timed out|timed out/i.test(reason)) {
                    throw new MailSendIndeterminateError(`AppleScript send timed out: ${reason}`);
                }
                throw new MailSendError(reason);
            }
            logAudit('send_email', { messageId: result.messageId ?? '', ...auditCounts(params) });
            return {
                messageId: result.messageId ?? '',
                sentAt: result.sentAt ?? '',
            };
        }
        finally {
            // Clean up temp file
            if (bodyFilePath != null) {
                try {
                    unlinkSync(bodyFilePath);
                }
                catch { /* ignore cleanup errors */ }
            }
        }
    }
    /**
     * Validates attachment paths, builds the create-draft AppleScript, executes
     * it, and parses the result. Unlike sendEmail, this saves the message to the
     * account's drafts folder (which syncs to other devices) instead of sending.
     * @param params - Email composition parameters (recipients, body, attachments, etc.).
     * @returns The saved draft's message ID.
     * @throws ValidationError if an attachment or inline image path is disallowed.
     * @throws AttachmentNotFoundError if any attachment or inline image path does not exist.
     * @throws MailSendError if Outlook reports a failure saving the draft.
     * @throws MailSendIndeterminateError if the operation times out mid-flight (verify in Drafts before retrying).
     */
    createDraft(params) {
        params = withCheckedFiles(params);
        // Same pre-write health gate as send: never fire AppleScript composing
        // into a degraded bridge — it hangs, gets SIGKILLed, and corrupts the bridge.
        assertBridgeHealthyForWrite('create_draft');
        let bodyFilePath;
        if (params.bodyType === 'html') {
            const id = randomBytes(8).toString('hex');
            bodyFilePath = join(tmpdir(), `mcp-outlook-html-${id}.html`);
            writeFileSync(bodyFilePath, params.body, 'utf8');
        }
        try {
            let scriptParams = {
                to: params.to,
                subject: params.subject,
                body: params.body,
                bodyType: params.bodyType,
                ...(bodyFilePath != null && { bodyFilePath }),
            };
            if (params.cc != null)
                scriptParams = { ...scriptParams, cc: params.cc };
            if (params.bcc != null)
                scriptParams = { ...scriptParams, bcc: params.bcc };
            if (params.replyTo != null)
                scriptParams = { ...scriptParams, replyTo: params.replyTo };
            if (params.attachments != null)
                scriptParams = { ...scriptParams, attachments: params.attachments };
            if (params.inlineImages != null)
                scriptParams = { ...scriptParams, inlineImages: params.inlineImages };
            if (params.accountId != null)
                scriptParams = { ...scriptParams, accountId: params.accountId };
            const messageId = runDraftScript(scripts.createDraft(scriptParams));
            logAudit('create_draft', { messageId, ...auditCounts(params) });
            return { messageId };
        }
        finally {
            if (bodyFilePath != null) {
                try {
                    unlinkSync(bodyFilePath);
                }
                catch { /* ignore cleanup errors */ }
            }
        }
    }
    /**
     * Saves a threaded reply or reply-all to an existing message as a DRAFT in
     * that message's account Drafts folder. Never sends.
     * @throws MailSendError if Outlook reports a failure saving the draft.
     * @throws MailSendIndeterminateError if it times out mid-flight (verify in Drafts before retrying).
     */
    createReplyDraft(params) {
        assertBridgeHealthyForWrite('reply_to_email');
        const html = params.bodyType === 'html' ? params.body : plainTextToHtml(params.body);
        const bodyFilePath = join(tmpdir(), `mcp-outlook-html-${randomBytes(8).toString('hex')}.html`);
        writeFileSync(bodyFilePath, html, 'utf8');
        try {
            const messageId = runDraftScript(scripts.createReplyDraft({
                messageId: params.messageId,
                replyAll: params.replyAll,
                bodyFilePath,
            }));
            logAudit('reply_to_email', { sourceMessageId: params.messageId, replyAll: params.replyAll, messageId });
            return { messageId };
        }
        finally {
            try {
                unlinkSync(bodyFilePath);
            }
            catch { /* ignore cleanup errors */ }
        }
    }
}
/**
 * Runs a draft-saving script and returns the saved draft's id. A timeout at
 * either layer is INDETERMINATE (the draft may or may not exist), so the caller
 * verifies in Drafts rather than creating a duplicate.
 */
function runDraftScript(script) {
    let output;
    try {
        output = executeAppleScriptOrThrow(script, { timeoutMs: SEND_NODE_TIMEOUT_MS });
    }
    catch (err) {
        if (err instanceof AppleScriptExecutionError && err.errorType === 'timeout') {
            throw new MailSendIndeterminateError('osascript was killed at the Node 60s timeout while saving the draft — check Drafts before retrying');
        }
        throw err;
    }
    const result = parseSendEmailResult(output);
    if (result == null) {
        throw new AppleScriptError('Failed to parse draft response');
    }
    if (!result.success) {
        const reason = result.error ?? 'Unknown error';
        // Inner timeout (-1712): the draft may or may not have saved/moved.
        if (/-1712|AppleEvent timed out|timed out/i.test(reason)) {
            throw new MailSendIndeterminateError(`AppleScript draft save timed out: ${reason} — check Drafts before retrying`);
        }
        throw new MailSendError(reason);
    }
    return result.messageId ?? '';
}
/** Escapes plain text as HTML, one `<br>` per line break, followed by a blank line. */
export function plainTextToHtml(text) {
    const escaped = text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/\r\n|\r|\n/g, '<br>');
    return `<div>${escaped}</div><br>`;
}
/** Creates a new AppleScriptMailSender instance. */
export function createMailSender() {
    return new AppleScriptMailSender();
}
