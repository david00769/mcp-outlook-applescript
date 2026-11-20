/**
 * Filesystem path safety checks for outbound/download operations.
 *
 * The MCP tools let a model choose filesystem paths: `save_path` for
 * downloaded attachments (arbitrary write -> possible persistence) and
 * `attachments[].path` / `inline_images[].path` for outgoing mail and drafts
 * (arbitrary read -> possible exfiltration of sensitive files). These helpers
 * confine those paths. Adopted from MagnasiePro's fork, with symlinks resolved
 * so a link cannot carry a path past either check.
 *
 * Configuration (environment variables):
 *  - OUTLOOK_MCP_DOWNLOAD_DIR: base directory downloads must stay within.
 *    Defaults to ~/Downloads.
 *  - OUTLOOK_MCP_ATTACHMENT_DIR: if set, outgoing attachments must stay
 *    within this directory (strict allowlist mode). If unset, a denylist of
 *    sensitive locations is enforced instead (so ordinary files still work).
 */
import { existsSync, lstatSync, realpathSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, dirname, isAbsolute, join, resolve, sep } from 'node:path';
import { ValidationError } from './errors.js';
/** Expands a leading `~` (or `~/`) to the current user's home directory. */
export function expandHome(p) {
    if (p === '~') {
        return homedir();
    }
    if (p.startsWith('~/')) {
        return homedir() + p.slice(1);
    }
    return p;
}
/** Resolves a path to an absolute, normalized form with `~` expansion. */
function normalizeAbsolute(p) {
    const expanded = expandHome(p);
    if (!isAbsolute(expanded)) {
        throw new ValidationError(`Path must be absolute: ${p}`);
    }
    return resolve(expanded);
}
/** Resolves symlinks when the path exists; otherwise returns it unchanged. */
function realIfExists(p) {
    return existsSync(p) ? realpathSync(p) : p;
}
/**
 * Returns true when `child` is `base` itself or nested inside `base`. Compared
 * case-insensitively on macOS, whose default APFS volume ignores case, so
 * `~/downloads/x` is inside `~/Downloads`.
 */
function isWithin(base, child) {
    const fold = (p) => (process.platform === 'darwin' ? p.toLowerCase() : p);
    const b = fold(base);
    const c = fold(child);
    return c === b || c.startsWith(b + sep);
}
/**
 * Path segments that must never be an attachment source. Matched
 * case-insensitively against the resolved (symlink-free) absolute path, so
 * `/etc` also matches macOS's `/private/etc`.
 */
const SENSITIVE_PATH_SEGMENTS = [
    `${sep}.ssh${sep}`,
    `${sep}.aws${sep}`,
    `${sep}.gnupg${sep}`,
    `${sep}.gcloud${sep}`,
    `${sep}.kube${sep}`,
    `${sep}.docker${sep}`,
    `${sep}.config${sep}`,
    `${sep}.appstoreconnect${sep}`,
    `${sep}secrets${sep}`,
    `${sep}library${sep}keychains${sep}`,
    `${sep}etc${sep}`,
];
/** Basenames that indicate secret material. */
const SENSITIVE_BASENAMES = [
    'id_rsa', 'id_dsa', 'id_ecdsa', 'id_ed25519',
    '.netrc', '.npmrc', '.pgpass', '.htpasswd', '.git-credentials', '.pypirc',
    '.bash_history', '.zsh_history', '.claude.json',
];
/** Basename prefixes that indicate secret material (`.env`, `.env.local`, ...). */
const SENSITIVE_PREFIXES = ['.env'];
/** File extensions that typically hold private keys / credentials. */
const SENSITIVE_EXTENSIONS = ['.pem', '.key', '.p8', '.p12', '.pfx', '.keychain', '.env'];
/**
 * Validates a destination path for a downloaded attachment.
 * Confines writes to OUTLOOK_MCP_DOWNLOAD_DIR (default ~/Downloads), after
 * resolving symlinks in the base and the destination directory, and refuses
 * to write through an existing symlink. Limit: Outlook writes the file after
 * this check, so a symlink swapped in between is not caught; the check stops a
 * model-chosen path, not a local process racing it.
 * @returns The resolved absolute path to write to.
 * @throws ValidationError if the path escapes the allowed base directory.
 */
export function assertSafeDownloadPath(savePath) {
    const base = realIfExists(normalizeAbsolute(expandHome(process.env['OUTLOOK_MCP_DOWNLOAD_DIR'] ?? `${homedir()}${sep}Downloads`)));
    const requested = normalizeAbsolute(savePath);
    const resolved = join(realIfExists(dirname(requested)), basename(requested));
    if (!isWithin(base, resolved)) {
        throw new ValidationError(`Refusing to write outside the allowed download directory (${base}). ` +
            `Set OUTLOOK_MCP_DOWNLOAD_DIR to change it.`);
    }
    if (existsSync(resolved) && lstatSync(resolved).isSymbolicLink()) {
        throw new ValidationError(`Refusing to write through a symbolic link: ${savePath}`);
    }
    return resolved;
}
/**
 * Validates a source path for an outgoing-mail attachment or inline image.
 * The path is resolved through symlinks first. In allowlist mode
 * (OUTLOOK_MCP_ATTACHMENT_DIR set) it must stay inside that directory.
 * Otherwise a denylist of sensitive locations (SSH/GPG/cloud credentials,
 * keychains, private keys, secret stores, /etc, dotfile config) is rejected
 * to blunt file exfiltration by email.
 * @returns The resolved absolute path.
 * @throws ValidationError if the path is disallowed.
 */
export function assertSafeAttachmentPath(attachmentPath) {
    const resolved = realIfExists(normalizeAbsolute(attachmentPath));
    const allowDir = process.env['OUTLOOK_MCP_ATTACHMENT_DIR'];
    if (allowDir != null && allowDir.length > 0) {
        const base = realIfExists(normalizeAbsolute(allowDir));
        if (!isWithin(base, resolved)) {
            throw new ValidationError(`Attachment path must be inside the allowed attachment directory (${base}).`);
        }
        return resolved;
    }
    const lower = resolved.toLowerCase();
    const name = lower.split(sep).pop() ?? '';
    const disallowed = SENSITIVE_PATH_SEGMENTS.some((seg) => lower.includes(seg)) ||
        SENSITIVE_BASENAMES.includes(name) ||
        SENSITIVE_PREFIXES.some((prefix) => name.startsWith(prefix)) ||
        SENSITIVE_EXTENSIONS.some((ext) => name.endsWith(ext));
    if (disallowed) {
        throw new ValidationError(`Refusing to attach a sensitive file: ${attachmentPath}. ` +
            `Set OUTLOOK_MCP_ATTACHMENT_DIR to restrict attachments to a specific directory.`);
    }
    return resolved;
}
