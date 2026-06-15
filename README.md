# OneDrive Bidirectional Sync

[MIT License](LICENSE)

一个不依赖本地 OneDrive 客户端的 Obsidian 双向同步插件 MVP，可运行于 Windows、Linux、macOS、iOS 和 Android。

## 工作方式

- 通过 Microsoft Graph 直接访问 OneDrive。
- 仅申请 `Files.ReadWrite.AppFolder` 权限，远端数据位于 `OneDrive/Apps/<应用名称>/vaults/<库 ID>/`。
- 使用本机同步快照判断文件在本地或远端是否变化。
- 两端同时修改同一文件时，保留一个带“本地冲突”后缀的副本，并将远端版本写入原路径。
- 默认不同步 `.obsidian`，避免插件自身令牌和设备工作区配置被同步。

## 安装与构建

```powershell
npm install
npm run build
```

将 `manifest.json`、`main.js` 复制到库的 `.obsidian/plugins/onedrive-bidirectional-sync/`，然后在 Obsidian 中启用插件。

## Microsoft Entra 应用注册

1. 在 Microsoft Entra 管理中心创建应用注册。
2. “支持的账户类型”选择同时支持组织目录与个人 Microsoft 账户。
3. 在“身份验证”中启用“允许公共客户端流”。
4. 添加 Microsoft Graph 委托权限 `Files.ReadWrite.AppFolder`；`offline_access` 会在登录时请求。
5. 将 Application (client) ID 填入插件设置。

所有设备必须填写相同的客户端 ID、登录同一个 Microsoft 账户，并填写相同的库 ID。

## 当前限制

- 单文件上传使用 Graph 简单上传接口，最大支持 250 MB。
- 初始同步按修改时间决定同名文件方向；建议先在主设备上传，再连接其他设备。
- iOS 和 Android 会限制后台运行，打开 Obsidian 后或手动执行同步更可靠。
- 令牌保存在 Obsidian 插件数据文件中。不要同步或分享插件自身的 `data.json`。
- 这是 MVP，尚未实现 Graph delta 增量扫描、分块上传、端到端加密和自动化测试。

## 官方接口文档

- [Microsoft Graph OneDrive app folder](https://learn.microsoft.com/graph/onedrive-sharepoint-appfolder)
- [Microsoft identity platform device authorization grant](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-device-code)
- [Microsoft Graph upload or replace file contents](https://learn.microsoft.com/graph/api/driveitem-put-content)
