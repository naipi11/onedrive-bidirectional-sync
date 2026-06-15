# Community plugin submission checklist

Repository target: `naipi11/onedrive-bidirectional-sync`

Before submitting:

- MIT license added.
- Test sign-in, upload, download, deletion, and conflicts with a real OneDrive account.
- Test on desktop and at least one mobile platform, including iPadOS.
- Create a public GitHub repository and push this source tree.
- Create the `0.1.0` tag. The release workflow will attach `main.js`, `manifest.json`, and `versions.json`.
- Fork `obsidianmd/obsidian-releases`.
- Add the following entry to `community-plugins.json` and open a pull request:

```json
{
  "id": "onedrive-bidirectional-sync",
  "name": "OneDrive Bidirectional Sync",
  "author": "naipi11",
  "description": "Synchronize vault files across devices through a private OneDrive app folder.",
  "repo": "naipi11/onedrive-bidirectional-sync"
}
```

Do not submit until the test claims above are true.
