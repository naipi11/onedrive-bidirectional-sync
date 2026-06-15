# OneDrive Bidirectional Sync

Synchronize Obsidian vault files across Windows, Linux, macOS, iOS, and Android
through a private OneDrive app folder. The plugin communicates directly with
Microsoft Graph and does not require the OneDrive desktop client.

## Features

- Uses the least-privileged Microsoft Graph `Files.ReadWrite.AppFolder`
  permission.
- Uploads and downloads file changes between multiple devices.
- Propagates file deletions in both directions.
- Preserves a local conflict copy when the same file changed on both sides.
- Supports automatic synchronization at a configurable interval.
- Excludes `.obsidian` by default to avoid syncing device-specific settings and
  the plugin's authentication data.

## Important safety information

This plugin is an early release. Back up your vault before the first sync.
Avoid editing the same file on multiple devices at the same time.

Authentication tokens are stored in the plugin's local Obsidian data file.
Never share that file. The plugin always excludes its own `data.json` from
synchronization.

## Installation

Install the plugin from the Obsidian community plugin directory when available.
For manual installation, place `manifest.json` and `main.js` in:

```text
<vault>/.obsidian/plugins/onedrive-bidirectional-sync/
```

Then reload Obsidian and enable **OneDrive Bidirectional Sync** under Community
plugins.

## Microsoft Entra application setup

Each user currently needs a Microsoft Entra application registration:

1. Create an application registration in the Microsoft Entra admin center.
2. Select the account type that supports organizational directories and
   personal Microsoft accounts.
3. Enable **Allow public client flows** under Authentication.
4. Add the delegated Microsoft Graph permission `Files.ReadWrite.AppFolder`.
5. Enter the application's **Application (client) ID** in the plugin settings.

All devices must use the same client ID, Microsoft account, and vault ID.

## How synchronization works

Remote files are stored under the application's private OneDrive app folder:

```text
OneDrive/Apps/<application name>/vaults/<vault ID>/
```

The plugin stores a local synchronization snapshot and compares it with local
file metadata and OneDrive item ETags. A one-sided change is propagated to the
other side. If both sides changed, the local version is preserved as a conflict
copy and the remote version replaces the original path.

## Current limitations

- Simple uploads support files up to 250 MB.
- Initial synchronization resolves different same-path files using modification
  time. Upload from the primary device before connecting additional devices.
- iOS and Android may suspend background timers. Sync after opening Obsidian or
  run the manual sync command.
- Microsoft Graph delta queries, resumable uploads, end-to-end encryption, and
  automated integration tests are not yet implemented.

## Development

```bash
pnpm install
pnpm typecheck
pnpm build
```

## References

- [Microsoft Graph OneDrive app folder](https://learn.microsoft.com/graph/onedrive-sharepoint-appfolder)
- [Microsoft identity platform device authorization grant](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-device-code)
- [Microsoft Graph upload or replace file contents](https://learn.microsoft.com/graph/api/driveitem-put-content)

## License

[MIT](LICENSE)
