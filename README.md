# mcp-outlook-applescript

A [Model Context Protocol](https://modelcontextprotocol.io/) server for Microsoft Outlook on Mac. Access your mail, calendar, contacts, tasks, and notes through 51 MCP tools powered by AppleScript.

## Features

- **AppleScript backend** — works with classic Outlook for Mac, no network or authentication required
- **51 tools** — full coverage of mail, calendar, contacts, tasks, and notes
- **Approval system** — destructive operations (delete, move, junk) require a two-step prepare/confirm flow
- **Bridge protection** — adaptive inter-call throttle, folder-list TTL cache, and a `BRIDGE_STRESSED` safety brake reduce stress on Outlook's AppleScript scripting bridge under sustained MCP usage. See `CLAUDE.md` for details.
- **Zero native dependencies** — pure TypeScript + AppleScript, no C++ compilation needed

### Available Tools

| Category | Tools |
|----------|-------|
| **Mail** | `list_folders`, `list_emails`, `search_emails`, `get_email`, `get_unread_count`, `send_email`, `create_draft`, `reply_to_email`, `list_attachments`, `download_attachment` |
| **Mail Organization** | `mark_email_read`, `mark_email_unread`, `set_email_flag`, `clear_email_flag`, `set_email_categories`, `prepare/confirm_delete_email`, `prepare/confirm_move_email`, `prepare/confirm_archive_email`, `prepare/confirm_junk_email`, `prepare/confirm_batch_*` |
| **Folders** | `create_folder`, `rename_folder`, `move_folder`, `prepare/confirm_delete_folder`, `prepare/confirm_empty_folder` |
| **Calendar** | `list_calendars`, `list_events`, `get_event`, `search_events`, `create_event`, `update_event`, `delete_event`, `respond_to_event` |
| **Contacts** | `list_contacts`, `search_contacts`, `get_contact` |
| **Tasks** | `list_tasks`, `search_tasks`, `get_task` |
| **Notes** | `list_notes`, `search_notes`, `get_note` |
| **Accounts** | `list_accounts` |

## Install

```bash
git clone https://github.com/hasan-imam/mcp-outlook-applescript.git
cd mcp-outlook-applescript
npm install && npm run build && npm pack && npm install -g mcp-outlook-applescript-*.tgz
cd .. && rm -rf mcp-outlook-applescript
```

This installs `mcp-outlook-applescript` as a global command. The checkout can be deleted after install.

To update, re-run the commands above.

### Claude Code

```bash
claude mcp add outlook -- mcp-outlook-applescript
```

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "mcp-outlook-applescript"
    }
  }
}
```

## Safety

- **Untrusted content.** Every value read from Outlook (subjects, bodies, names, addresses, folder and account names) is passed through `sanitizeField`, which follows every `{` with an invisible zero-width space, so no two braces stay adjacent and a crafted email cannot forge fields or records in tool output. A `{` in returned text is therefore followed by an invisible character, so compare such text loosely.
- **Attachments and downloads.** `download_attachment` writes only inside `OUTLOOK_MCP_DOWNLOAD_DIR` (default `~/Downloads`) and never through a symbolic link. Attachments and inline images for `send_email` and `create_draft` may not be sensitive files (SSH, GPG and cloud credentials, keychains, private keys such as `.pem`/`.p8`, `.env`, `secrets` folders, `/etc`, dotfile config); set `OUTLOOK_MCP_ATTACHMENT_DIR` to allow only one directory instead. Paths are checked after resolving symlinks; Outlook writes a download after the check, so a local process racing it with a symlink is outside this guard.
- **Replies are drafts.** `reply_to_email` (reply or reply-all) saves a threaded draft in the original account's Drafts folder and never sends.
- **Audit log.** Sends, drafts, replies, downloads and mailbox moves/deletes write one JSON line to stderr with ids and counts only (no addresses, subjects or paths).

| Environment variable | Purpose | Default |
|----------------------|---------|---------|
| `OUTLOOK_MCP_DOWNLOAD_DIR` | Directory downloaded attachments must stay within | `~/Downloads` |
| `OUTLOOK_MCP_ATTACHMENT_DIR` | If set, outgoing attachments must stay within this directory | unset (sensitive-file denylist) |

## Requirements

- macOS with Microsoft Outlook (classic) installed and running
- Node.js >= 20
- Automation permission for Outlook (System Settings > Privacy & Security > Automation)

## Development

```bash
npm install
npm run build      # compile TypeScript → dist/
npm run typecheck   # type-check without emitting
npm test            # 121 unit tests
bash scripts/audit.sh  # static quality audit (build, security, package, functional)
```

## License

MIT — see [LICENSE](LICENSE).
