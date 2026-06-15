import { App, Modal, Notice, Plugin, PluginSettingTab, Setting } from "obsidian";
import { MicrosoftAuth } from "./auth";
import { GraphClient } from "./graph";
import { SyncEngine } from "./sync";
import type { SyncSettings, TokenState } from "./types";

const DEFAULT_SETTINGS: SyncSettings = {
  clientId: "",
  tenant: "common",
  vaultId: "",
  intervalMinutes: 10,
  syncOnStartup: true,
  syncConfigDir: false,
  excludedPatterns: ".trash/**\n.trash/",
  token: null,
  entries: {},
  lastSyncAt: 0
};

export default class OneDriveSyncPlugin extends Plugin {
  declare settings: SyncSettings;
  private intervalId: number | null = null;
  private syncInProgress = false;

  async onload(): Promise<void> {
    await this.loadSettings();
    this.addSettingTab(new OneDriveSyncSettingTab(this.app, this));
    this.addCommand({ id: "sync-now", name: "立即同步", callback: () => void this.syncNow() });
    this.addRibbonIcon("refresh-cw", "OneDrive 双向同步", () => void this.syncNow());
    this.configureTimer();

    if (this.settings.syncOnStartup && this.settings.token && this.settings.clientId) {
      this.app.workspace.onLayoutReady(() => window.setTimeout(() => void this.syncNow(), 3000));
    }
  }

  onunload(): void {
    if (this.intervalId !== null) window.clearInterval(this.intervalId);
  }

  async login(): Promise<void> {
    if (!this.settings.clientId.trim()) {
      new Notice("请先填写 Microsoft Entra 应用客户端 ID");
      return;
    }
    try {
      const auth = this.createAuth();
      const code = await auth.beginDeviceCode();
      new DeviceCodeModal(this.app, code.message, code.user_code, code.verification_uri).open();
      await auth.finishDeviceCode(code);
      new Notice("Microsoft 账户登录成功");
    } catch (error) {
      new Notice(errorMessage(error), 10000);
    }
  }

  async logout(): Promise<void> {
    this.settings.token = null;
    this.settings.entries = {};
    await this.saveSettings();
    new Notice("已清除本机 Microsoft 登录信息和同步快照");
  }

  async syncNow(): Promise<void> {
    if (!this.settings.clientId || !this.settings.token) {
      new Notice("请先在设置中登录 Microsoft 账户");
      return;
    }
    if (!this.settings.vaultId.trim()) {
      new Notice("库 ID 不能为空");
      return;
    }
    if (this.syncInProgress) {
      new Notice("OneDrive 同步已在进行中");
      return;
    }
    this.syncInProgress = true;
    new Notice("OneDrive 同步开始");
    try {
      const engine = new SyncEngine(this.app, new GraphClient(this.createAuth()), this.settings, () => this.saveSettings());
      const result = await engine.sync();
      new Notice(
        `同步完成：上传 ${result.uploaded}，下载 ${result.downloaded}，冲突 ${result.conflicts}，删除 ${result.deletedLocal + result.deletedRemote}`,
        8000
      );
    } catch (error) {
      console.error("OneDrive sync failed", error);
      new Notice(`同步失败：${errorMessage(error)}`, 12000);
    } finally {
      this.syncInProgress = false;
    }
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
  }

  configureTimer(): void {
    if (this.intervalId !== null) window.clearInterval(this.intervalId);
    this.intervalId = null;
    if (this.settings.intervalMinutes > 0) {
      this.intervalId = window.setInterval(() => void this.syncNow(), this.settings.intervalMinutes * 60_000);
      this.registerInterval(this.intervalId);
    }
  }

  private createAuth(): MicrosoftAuth {
    return new MicrosoftAuth(this.settings.clientId.trim(), this.settings.tenant.trim() || "common", this.settings.token, async (token: TokenState) => {
      this.settings.token = token;
      await this.saveSettings();
    });
  }

  private async loadSettings(): Promise<void> {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
    if (!this.settings.vaultId) {
      this.settings.vaultId = crypto.randomUUID();
      await this.saveSettings();
    }
  }
}

class DeviceCodeModal extends Modal {
  constructor(app: App, private readonly message: string, private readonly code: string, private readonly uri: string) {
    super(app);
  }

  onOpen(): void {
    this.titleEl.setText("登录 Microsoft 账户");
    this.contentEl.createEl("p", { text: this.message });
    this.contentEl.createEl("p", { text: `代码：${this.code}` });
    new Setting(this.contentEl)
      .addButton((button) => button.setButtonText("复制代码").onClick(() => void navigator.clipboard.writeText(this.code)))
      .addButton((button) => button.setButtonText("打开登录页面").setCta().onClick(() => window.open(this.uri)));
    this.contentEl.createEl("p", { text: "完成授权后，此窗口可关闭；插件会继续等待登录结果。" });
  }
}

class OneDriveSyncSettingTab extends PluginSettingTab {
  constructor(app: App, private readonly plugin: OneDriveSyncPlugin) {
    super(app, plugin);
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl("h2", { text: "OneDrive 双向同步" });
    containerEl.createEl("p", {
      text: "所有设备必须使用相同的客户端 ID 和库 ID。插件只访问 OneDrive 的应用专属目录。"
    });

    new Setting(containerEl)
      .setName("Microsoft Entra 客户端 ID")
      .setDesc("应用注册的 Application (client) ID；需启用公共客户端流并授予 Files.ReadWrite.AppFolder。")
      .addText((text) => text.setPlaceholder("xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx").setValue(this.plugin.settings.clientId).onChange(async (value) => {
        this.plugin.settings.clientId = value.trim();
        await this.plugin.saveSettings();
      }));

    new Setting(containerEl)
      .setName("租户")
      .setDesc("common 支持个人账户和组织账户；也可填写租户 ID。")
      .addText((text) => text.setValue(this.plugin.settings.tenant).onChange(async (value) => {
        this.plugin.settings.tenant = value.trim() || "common";
        await this.plugin.saveSettings();
      }));

    new Setting(containerEl)
      .setName("库 ID")
      .setDesc("同一个库在所有设备上必须一致。首次安装后，将此值复制到其他设备。")
      .addText((text) => text.setValue(this.plugin.settings.vaultId).onChange(async (value) => {
        this.plugin.settings.vaultId = value.trim();
        this.plugin.settings.entries = {};
        await this.plugin.saveSettings();
      }));

    new Setting(containerEl)
      .setName(this.plugin.settings.token ? "已登录 Microsoft 账户" : "登录 Microsoft 账户")
      .addButton((button) => button.setButtonText("登录").setCta().onClick(() => void this.plugin.login()))
      .addButton((button) => button.setButtonText("退出").onClick(() => void this.plugin.logout()));

    new Setting(containerEl)
      .setName("自动同步间隔（分钟）")
      .setDesc("设为 0 可关闭定时同步。移动系统可能暂停后台计时器。")
      .addText((text) => text.setValue(String(this.plugin.settings.intervalMinutes)).onChange(async (value) => {
        this.plugin.settings.intervalMinutes = Math.max(0, Number.parseInt(value, 10) || 0);
        await this.plugin.saveSettings();
        this.plugin.configureTimer();
      }));

    new Setting(containerEl)
      .setName("启动后同步")
      .addToggle((toggle) => toggle.setValue(this.plugin.settings.syncOnStartup).onChange(async (value) => {
        this.plugin.settings.syncOnStartup = value;
        await this.plugin.saveSettings();
      }));

    new Setting(containerEl)
      .setName("同步 .obsidian 配置目录")
      .setDesc("默认关闭。开启后插件、主题与工作区配置也会同步，发生冲突的风险更高。")
      .addToggle((toggle) => toggle.setValue(this.plugin.settings.syncConfigDir).onChange(async (value) => {
        this.plugin.settings.syncConfigDir = value;
        await this.plugin.saveSettings();
      }));

    new Setting(containerEl)
      .setName("排除规则")
      .setDesc("每行一个 glob，例如 attachments/cache/**。")
      .addTextArea((text) => text.setValue(this.plugin.settings.excludedPatterns).onChange(async (value) => {
        this.plugin.settings.excludedPatterns = value;
        await this.plugin.saveSettings();
      }));

    new Setting(containerEl)
      .setName("立即同步")
      .setDesc(this.plugin.settings.lastSyncAt ? `上次完成：${new Date(this.plugin.settings.lastSyncAt).toLocaleString()}` : "尚未同步")
      .addButton((button) => button.setButtonText("同步").setCta().onClick(() => void this.plugin.syncNow()));
  }
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
