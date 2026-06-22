import { Version } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { Dialog } from '@microsoft/sp-dialog';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import {
  createDefaultPreviewSettings,
  DEFAULT_APP_BASE_URL,
  type IPreviewConfigurationStatus,
  type IPreviewSettings,
  normalizeBaseUrl,
  PreviewSettingsService
} from '../../extensions/bpmnOpenCommandSet/previewSettings';
import { PreviewSettingsDialog } from '../../extensions/bpmnOpenCommandSet/PreviewSettingsDialog';
import { APP_VERSION, COMMAND_SET_COMPONENT_ID } from '../../shared/appConstants';

export interface IFilePreviewAdminWebPartProps {
  configSiteUrl?: string;
}

export default class FilePreviewAdminWebPart extends BaseClientSideWebPart<IFilePreviewAdminWebPartProps> {
  private configurationStatus: IPreviewConfigurationStatus | undefined;
  private settings: IPreviewSettings = createDefaultPreviewSettings(DEFAULT_APP_BASE_URL);
  private settingsService: PreviewSettingsService | undefined;

  public render(): void {
    const configSiteUrl = this.getConfigWebAbsoluteUrl();
    const canManage = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    const enabledExtensions = this.settings.extensions
      .filter((extension) => extension.enabled && extension.renderer !== 'coming-soon')
      .map((extension) => extension.extension)
      .join(', ') || 'None';
    const isInitialized = Boolean(
      this.configurationStatus?.configListExists &&
        this.configurationStatus.configFieldExists &&
        this.configurationStatus.configItemExists
    );

    this.domElement.innerHTML = `
      <section class="bpf-admin-page">
        <div class="bpf-admin-page__header">
          <div>
            <p class="bpf-admin-page__eyebrow">Microsoft 365 File Preview Framework</p>
            <h2>Tenant file preview settings</h2>
            <p>Manage the central defaults used by SharePoint document libraries across the tenant.</p>
          </div>
          <button class="bpf-admin-page__primary" data-action="open-settings" type="button" ${canManage ? '' : 'disabled'}>
            Open settings
          </button>
        </div>
        <div class="bpf-admin-page__status ${isInitialized ? '' : 'bpf-admin-page__status--attention'}">
          <div>
            <span>Configuration status</span>
            <strong>${isInitialized ? 'Initialized' : 'Setup required'}</strong>
          </div>
          <ul>
            <li data-status="${this.configurationStatus?.configListExists ? 'ok' : 'missing'}">Config list</li>
            <li data-status="${this.configurationStatus?.configFieldExists ? 'ok' : 'missing'}">Config field</li>
            <li data-status="${this.configurationStatus?.configItemExists ? 'ok' : 'missing'}">Config item</li>
          </ul>
          <button class="bpf-admin-page__secondary" data-action="initialize-config" type="button" ${
            canManage && !isInitialized ? '' : 'disabled'
          }>
            Initialize tenant configuration
          </button>
        </div>
        <div class="bpf-admin-page__grid">
          <div>
            <span>Configuration site</span>
            <strong>${escapeHtml(configSiteUrl)}</strong>
          </div>
          <div>
            <span>Enabled extensions</span>
            <strong>${escapeHtml(enabledExtensions)}</strong>
          </div>
          <div>
            <span>License tier</span>
            <strong>${escapeHtml(this.settings.license.tier)}</strong>
          </div>
          <div>
            <span>Native File Handler</span>
            <strong>${this.settings.fileHandlerEnabled ? 'Enabled' : 'Optional / not enabled'}</strong>
          </div>
        </div>
        <p class="bpf-admin-page__message" data-role="message">${
          canManage
            ? 'Use this page as the tenant administrator entry point. The settings are central and are not stored in individual libraries.'
            : 'You need Manage Web permissions on this site to edit the tenant file preview settings.'
        }</p>
        <footer class="bpf-admin-page__footer">
          <span>Version ${escapeHtml(APP_VERSION)}</span>
          <span>Command set ${escapeHtml(COMMAND_SET_COMPONENT_ID)}</span>
        </footer>
      </section>
    `;

    this.applyStyles();
    this.domElement.querySelector('[data-action="open-settings"]')?.addEventListener('click', () => {
      this.openSettingsDialog();
    });
    this.domElement.querySelector('[data-action="initialize-config"]')?.addEventListener('click', () => {
      this.initializeConfiguration().catch((error: unknown) => {
        this.setMessage(error instanceof Error ? error.message : 'Could not initialize configuration.', true);
      });
    });

    this.loadSettings().catch((error: unknown) => {
      this.setMessage(error instanceof Error ? error.message : 'Could not load settings.', true);
    });
  }

  protected async onInit(): Promise<void> {
    this.settingsService = new PreviewSettingsService(this.context.spHttpClient, this.getConfigWebAbsoluteUrl());
    await this.loadSettings();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.2');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Tenant file preview configuration'
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('configSiteUrl', {
                  description: 'Use the App Catalog or a dedicated admin site for tenant-level configuration. Leave blank to use the tenant root site.',
                  label: 'Central configuration site URL',
                  placeholder: getTenantRootSiteUrl(this.context.pageContext.web.absoluteUrl)
                })
              ],
              groupName: 'Configuration scope'
            }
          ]
        }
      ]
    };
  }

  private async loadSettings(): Promise<void> {
    this.configurationStatus = await this.getSettingsService().getConfigurationStatus();
    this.settings = await this.getSettingsService().getSettings(DEFAULT_APP_BASE_URL);
    this.renderLoadedState();
  }

  private async initializeConfiguration(): Promise<void> {
    if (!this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb)) {
      Dialog.alert('You need Manage Web permissions to initialize file preview settings.').catch(() => undefined);
      return;
    }

    this.setMessage('Initializing tenant configuration...', false);
    this.configurationStatus = await this.getSettingsService().initializeConfiguration(this.settings);
    this.settings = await this.getSettingsService().getSettings(DEFAULT_APP_BASE_URL);
    this.setMessage('Tenant configuration initialized. Existing settings are preserved on upgrades.', false);
    this.render();
  }

  private openSettingsDialog(): void {
    if (!this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb)) {
      Dialog.alert('You need Manage Web permissions to update file preview settings.').catch(() => undefined);
      return;
    }

    const dialog = new PreviewSettingsDialog(this.getSettingsService(), this.settings, this.getAdminScriptContext(), (savedSettings) => {
      this.settings = savedSettings;
      this.render();
    });
    dialog.show().catch((error: unknown) => {
      this.setMessage(error instanceof Error ? error.message : 'Could not open settings.', true);
    });
  }

  private renderLoadedState(): void {
    const enabledExtensionsElement = this.domElement.querySelector('.bpf-admin-page__grid div:nth-child(2) strong');
    const licenseElement = this.domElement.querySelector('.bpf-admin-page__grid div:nth-child(3) strong');
    const fileHandlerElement = this.domElement.querySelector('.bpf-admin-page__grid div:nth-child(4) strong');
    const statusElement = this.domElement.querySelector('.bpf-admin-page__status') as HTMLElement | null;
    const statusValueElement = this.domElement.querySelector('.bpf-admin-page__status strong');
    const initializeButton = this.domElement.querySelector('[data-action="initialize-config"]') as HTMLButtonElement | null;

    const enabledExtensions = this.settings.extensions
      .filter((extension) => extension.enabled && extension.renderer !== 'coming-soon')
      .map((extension) => extension.extension)
      .join(', ') || 'None';
    const isInitialized = Boolean(
      this.configurationStatus?.configListExists &&
        this.configurationStatus.configFieldExists &&
        this.configurationStatus.configItemExists
    );

    if (enabledExtensionsElement) {
      enabledExtensionsElement.textContent = enabledExtensions;
    }
    if (licenseElement) {
      licenseElement.textContent = this.settings.license.tier;
    }
    if (fileHandlerElement) {
      fileHandlerElement.textContent = this.settings.fileHandlerEnabled ? 'Enabled' : 'Optional / not enabled';
    }
    if (statusElement) {
      statusElement.classList.toggle('bpf-admin-page__status--attention', !isInitialized);
    }
    if (statusValueElement) {
      statusValueElement.textContent = isInitialized ? 'Initialized' : 'Setup required';
    }
    this.updateStatusPill(0, Boolean(this.configurationStatus?.configListExists));
    this.updateStatusPill(1, Boolean(this.configurationStatus?.configFieldExists));
    this.updateStatusPill(2, Boolean(this.configurationStatus?.configItemExists));
    if (initializeButton) {
      initializeButton.disabled = isInitialized || !this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    }
  }

  private updateStatusPill(index: number, isOk: boolean): void {
    const pill = this.domElement.querySelectorAll('.bpf-admin-page__status li')[index] as HTMLElement | undefined;
    if (pill) {
      pill.dataset.status = isOk ? 'ok' : 'missing';
    }
  }

  private setMessage(message: string, isError: boolean): void {
    const messageElement = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.textContent = message;
    messageElement.classList.toggle('bpf-admin-page__message--error', isError);
  }

  private getSettingsService(): PreviewSettingsService {
    if (!this.settingsService) {
      this.settingsService = new PreviewSettingsService(this.context.spHttpClient, this.getConfigWebAbsoluteUrl());
    }

    return this.settingsService;
  }

  private getConfigWebAbsoluteUrl(): string {
    return normalizeBaseUrl(this.properties.configSiteUrl || getTenantRootSiteUrl(this.context.pageContext.web.absoluteUrl));
  }

  private getAdminScriptContext(): {
    componentId: string;
    configSiteUrl: string;
    currentSiteUrl: string;
    tenantHostName: string;
    tenantId: string;
  } {
    return {
      componentId: COMMAND_SET_COMPONENT_ID,
      configSiteUrl: this.getConfigWebAbsoluteUrl(),
      currentSiteUrl: this.context.pageContext.web.absoluteUrl,
      tenantHostName: new URL(this.context.pageContext.web.absoluteUrl).hostname,
      tenantId: getTenantId(this.context.pageContext.aadInfo)
    };
  }

  private applyStyles(): void {
    if (this.domElement.querySelector('style[data-bpf-admin-page]')) {
      return;
    }

    const style = document.createElement('style');
    style.dataset.bpfAdminPage = 'true';
    style.textContent = `
      .bpf-admin-page {
        background: #ffffff;
        border: 1px solid #e1dfdd;
        box-sizing: border-box;
        color: #242424;
        display: grid;
        font-family: "Segoe UI", Arial, sans-serif;
        gap: 20px;
        padding: 24px;
      }
      .bpf-admin-page__header {
        align-items: start;
        display: flex;
        gap: 20px;
        justify-content: space-between;
      }
      .bpf-admin-page__eyebrow,
      .bpf-admin-page p {
        color: #605e5c;
        margin: 0;
      }
      .bpf-admin-page__eyebrow {
        font-size: 12px;
        font-weight: 600;
        letter-spacing: 0;
        text-transform: uppercase;
      }
      .bpf-admin-page h2 {
        font-size: 26px;
        font-weight: 600;
        margin: 4px 0 6px;
      }
      .bpf-admin-page__grid {
        display: grid;
        gap: 12px;
        grid-template-columns: repeat(4, minmax(0, 1fr));
      }
      .bpf-admin-page__status {
        align-items: center;
        border: 1px solid #d1d1d1;
        display: grid;
        gap: 16px;
        grid-template-columns: minmax(0, 1fr) minmax(260px, 1fr) auto;
        padding: 14px;
      }
      .bpf-admin-page__status--attention {
        border-color: #f2c811;
        background: #fffdf3;
      }
      .bpf-admin-page__status div {
        display: grid;
        gap: 4px;
      }
      .bpf-admin-page__status span,
      .bpf-admin-page__status li {
        color: #605e5c;
        font-size: 12px;
        font-weight: 600;
      }
      .bpf-admin-page__status strong {
        font-size: 15px;
      }
      .bpf-admin-page__status ul {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        list-style: none;
        margin: 0;
        padding: 0;
      }
      .bpf-admin-page__status li {
        border: 1px solid #d1d1d1;
        padding: 5px 8px;
      }
      .bpf-admin-page__status li[data-status="ok"] {
        border-color: #107c10;
        color: #107c10;
      }
      .bpf-admin-page__status li[data-status="missing"] {
        border-color: #8a8886;
      }
      .bpf-admin-page__grid div {
        border: 1px solid #edebe9;
        display: grid;
        gap: 6px;
        min-height: 74px;
        padding: 14px;
      }
      .bpf-admin-page__grid span {
        color: #605e5c;
        font-size: 12px;
        font-weight: 600;
      }
      .bpf-admin-page__grid strong {
        font-size: 15px;
        font-weight: 600;
        overflow-wrap: anywhere;
      }
      .bpf-admin-page__primary {
        background: #0078d4;
        border: 1px solid #0078d4;
        color: #ffffff;
        cursor: pointer;
        font: inherit;
        min-height: 36px;
        padding: 0 18px;
      }
      .bpf-admin-page__secondary {
        background: #ffffff;
        border: 1px solid #8a8886;
        color: #242424;
        cursor: pointer;
        font: inherit;
        min-height: 36px;
        padding: 0 14px;
      }
      .bpf-admin-page__primary:disabled {
        background: #f3f2f1;
        border-color: #c8c6c4;
        color: #605e5c;
        cursor: not-allowed;
      }
      .bpf-admin-page__secondary:disabled {
        background: #f3f2f1;
        border-color: #c8c6c4;
        color: #605e5c;
        cursor: not-allowed;
      }
      .bpf-admin-page__message--error {
        color: #a4262c;
      }
      .bpf-admin-page__footer {
        align-items: center;
        border-top: 1px solid #edebe9;
        color: #605e5c;
        display: flex;
        flex-wrap: wrap;
        font-size: 12px;
        gap: 12px;
        justify-content: space-between;
        padding-top: 14px;
      }
      @media (max-width: 900px) {
        .bpf-admin-page__header,
        .bpf-admin-page__grid,
        .bpf-admin-page__status {
          grid-template-columns: 1fr;
        }
        .bpf-admin-page__header {
          display: grid;
        }
      }
    `;
    this.domElement.prepend(style);
  }
}

function getTenantRootSiteUrl(webAbsoluteUrl: string): string {
  const parsed = new URL(webAbsoluteUrl);
  return `${parsed.protocol}//${parsed.hostname}`;
}

function getTenantId(aadInfo: unknown): string {
  const candidate = aadInfo as { tenantId?: { toString: () => string } };
  return candidate.tenantId?.toString() || '';
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
