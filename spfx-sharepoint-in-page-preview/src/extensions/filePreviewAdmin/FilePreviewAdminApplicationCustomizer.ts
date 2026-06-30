import { Log } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import {
  createDefaultPreviewSettings,
  DEFAULT_APP_BASE_URL,
  normalizeBaseUrl,
  PreviewSettingsService
} from '../bpmnOpenCommandSet/previewSettings';
import { PreviewSettingsDialog } from '../bpmnOpenCommandSet/PreviewSettingsDialog';
import { APP_VERSION, COMMAND_SET_COMPONENT_ID } from '../../shared/appConstants';

export interface IFilePreviewAdminApplicationCustomizerProperties {
  configSiteUrl?: string;
}

const LOG_SOURCE: string = 'FilePreviewAdminApplicationCustomizer';

export default class FilePreviewAdminApplicationCustomizer extends BaseApplicationCustomizer<IFilePreviewAdminApplicationCustomizerProperties> {
  private launcherElement: HTMLButtonElement | undefined;
  private placeholder: { domElement: HTMLElement; dispose: () => void } | undefined;
  private settingsService: PreviewSettingsService | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized File Preview tenant admin launcher ${APP_VERSION}`);

    if (!this.shouldRenderLauncher()) {
      this.openFromQueryString();
      return Promise.resolve();
    }

    this.renderLauncher();
    this.openFromQueryString();
    return Promise.resolve();
  }

  protected onDispose(): void {
    this.placeholder?.dispose();
    this.placeholder = undefined;
    this.launcherElement = undefined;
    super.onDispose();
  }

  private shouldRenderLauncher(): boolean {
    if (!this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb)) {
      return false;
    }
    return this.isTenantRootWeb() || this.isAppCatalogSite();
  }

  private isTenantRootWeb(): boolean {
    return normalizeBaseUrl(this.context.pageContext.web.absoluteUrl) === getTenantRootSiteUrl(this.context.pageContext.web.absoluteUrl);
  }

  private isAppCatalogSite(): boolean {
    // SharePoint sets isAppCatalogSite on the page context of the tenant App Catalog site.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return Boolean((this.context.pageContext as any).legacyPageContext?.isAppCatalogSite);
  }

  private renderLauncher(): void {
    if (this.launcherElement || this.placeholder) {
      return;
    }

    this.placeholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, {
      onDispose: () => {
        this.placeholder = undefined;
        this.launcherElement = undefined;
      }
    });

    if (!this.placeholder) {
      Log.warn(LOG_SOURCE, 'Bottom placeholder unavailable — admin launcher will not be shown.');
      return;
    }

    const style = document.createElement('style');
    style.textContent = `
      .bpf-tenant-admin-launcher {
        align-items: center;
        background: #242424;
        border: 1px solid #605e5c;
        box-shadow: 0 -1px 0 rgba(0, 0, 0, 0.1);
        color: #ffffff;
        cursor: pointer;
        display: inline-flex;
        font: 600 13px "Segoe UI", Arial, sans-serif;
        gap: 8px;
        min-height: 40px;
        padding: 0 14px;
      }
      .bpf-tenant-admin-launcher:focus,
      .bpf-tenant-admin-launcher:hover {
        background: #000000;
      }
      .bpf-tenant-admin-launcher__mark {
        background: #0078d4;
        color: #ffffff;
        display: inline-flex;
        font-size: 11px;
        font-weight: 700;
        justify-content: center;
        min-width: 36px;
        padding: 2px 4px;
      }
    `;
    this.placeholder.domElement.appendChild(style);

    this.launcherElement = document.createElement('button');
    this.launcherElement.className = 'bpf-tenant-admin-launcher';
    this.launcherElement.type = 'button';
    this.launcherElement.title = 'Open File Preview tenant admin settings';
    this.launcherElement.innerHTML = `<span class="bpf-tenant-admin-launcher__mark">BPMN</span><span>File Preview Admin</span>`;
    this.launcherElement.addEventListener('click', () => {
      this.openAdminSettings().catch((error: unknown) => {
        Dialog.alert(getAdminErrorMessage(error)).catch(() => undefined);
      });
    });

    this.placeholder.domElement.appendChild(this.launcherElement);
  }

  private openFromQueryString(): void {
    const query = new URLSearchParams(window.location.search);
    if (query.get('m365FilePreviewAdmin') !== '1') {
      return;
    }

    if (!this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb)) {
      return;
    }

    this.openAdminSettings().catch((error: unknown) => {
      Dialog.alert(getAdminErrorMessage(error)).catch(() => undefined);
    });
  }

  private async openAdminSettings(): Promise<void> {
    const settingsService = this.getSettingsService();
    const defaults = createDefaultPreviewSettings(DEFAULT_APP_BASE_URL);
    await settingsService.initializeConfiguration(defaults);
    const settings = await settingsService.getSettings(DEFAULT_APP_BASE_URL);
    const dialog = new PreviewSettingsDialog(settingsService, settings, this.getAdminScriptContext(), () => undefined);
    await dialog.show();
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

}

function getTenantRootSiteUrl(webAbsoluteUrl: string): string {
  const parsed = new URL(webAbsoluteUrl);
  return `${parsed.protocol}//${parsed.hostname}`;
}

function getTenantId(aadInfo: unknown): string {
  const candidate = aadInfo as { tenantId?: { toString: () => string } };
  return candidate.tenantId?.toString() || '';
}

function getAdminErrorMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : '';
  const parsedMessage = tryReadSharePointErrorMessage(message);
  const detail = parsedMessage || message || 'SharePoint returned an unexpected error.';

  return (
    'File Preview Admin could not finish tenant configuration repair. ' +
    'Confirm you are opening the tenant root site as a SharePoint admin, then refresh and try again. ' +
    `Details: ${detail}`
  );
}

function tryReadSharePointErrorMessage(value: string): string {
  const trimmed = value.trim();
  if (!trimmed.startsWith('{')) {
    return '';
  }

  try {
    const parsed = JSON.parse(trimmed) as { error?: { message?: string | { value?: string } } };
    const message = parsed.error?.message;
    if (typeof message === 'string') {
      return message;
    }

    return message?.value || '';
  } catch {
    return '';
  }
}
