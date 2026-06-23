import { Log } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import {
  createDefaultPreviewSettings,
  DEFAULT_APP_BASE_URL,
  findExtensionSettings,
  type IFileExtensionSettings,
  type IPreviewSettings,
  normalizeBaseUrl,
  PreviewSettingsService
} from '../bpmnOpenCommandSet/previewSettings';
import { PreviewSettingsDialog } from '../bpmnOpenCommandSet/PreviewSettingsDialog';
import { BpmnViewerDialog } from '../bpmnOpenCommandSet/BpmnViewerDialog';
import { DrawioViewerDialog } from '../bpmnOpenCommandSet/DrawioViewerDialog';
import { APP_VERSION, COMMAND_SET_COMPONENT_ID } from '../../shared/appConstants';

export interface IFilePreviewAdminApplicationCustomizerProperties {
  configSiteUrl?: string;
  showOnAllSites?: boolean;
}

const LOG_SOURCE: string = 'FilePreviewAdminApplicationCustomizer';
const CANDIDATE_EXTENSIONS: string[] = ['.bpmn', '.drawio', '.jt', '.step'];

export default class FilePreviewAdminApplicationCustomizer extends BaseApplicationCustomizer<IFilePreviewAdminApplicationCustomizerProperties> {
  private launcherElement: HTMLButtonElement | undefined;
  private previewLauncherElement: HTMLButtonElement | undefined;
  private previewSettings: IPreviewSettings = createDefaultPreviewSettings(DEFAULT_APP_BASE_URL);
  private selectedFile: { extensionSettings: IFileExtensionSettings; fileName: string; serverRelativeUrl: string } | undefined;
  private settingsService: PreviewSettingsService | undefined;
  private urlWatchHandle: number | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized File Preview tenant admin launcher ${APP_VERSION}`);
    this.loadPreviewSettings().catch((error: unknown) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error('Could not load File Preview settings.'));
    });
    this.startSelectedFileWatcher();

    if (!this.shouldRenderLauncher()) {
      this.openFromQueryString();
      return Promise.resolve();
    }

    this.renderLauncher();
    this.openFromQueryString();
    return Promise.resolve();
  }

  protected onDispose(): void {
    if (this.urlWatchHandle !== undefined) {
      window.clearInterval(this.urlWatchHandle);
      this.urlWatchHandle = undefined;
    }

    this.removePreviewLauncher();
    this.launcherElement?.remove();
    this.launcherElement = undefined;
    super.onDispose();
  }

  private shouldRenderLauncher(): boolean {
    if (!this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb)) {
      return false;
    }

    return Boolean(this.properties.showOnAllSites) || this.isTenantRootWeb();
  }

  private renderLauncher(): void {
    if (this.launcherElement) {
      return;
    }

    const style = document.createElement('style');
    style.textContent = `
      .bpf-tenant-admin-launcher {
        align-items: center;
        background: #242424;
        border: 1px solid #605e5c;
        bottom: 20px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.18);
        color: #ffffff;
        cursor: pointer;
        display: inline-flex;
        font: 600 13px "Segoe UI", Arial, sans-serif;
        gap: 8px;
        min-height: 40px;
        padding: 0 14px;
        position: fixed;
        right: 20px;
        z-index: 1000000;
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
    document.head.appendChild(style);

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

    document.body.appendChild(this.launcherElement);
  }

  private renderPreviewLauncher(): void {
    if (this.previewLauncherElement) {
      return;
    }

    this.ensureLauncherStyles();
    this.previewLauncherElement = document.createElement('button');
    this.previewLauncherElement.className = 'bpf-file-preview-launcher';
    this.previewLauncherElement.type = 'button';
    this.updatePreviewLauncherText();
    this.previewLauncherElement.addEventListener('click', () => {
      this.openSelectedFilePreview().catch((error: unknown) => {
        Dialog.alert(error instanceof Error ? error.message : 'Could not open the selected file.').catch(() => undefined);
      });
    });

    document.body.appendChild(this.previewLauncherElement);
  }

  private removePreviewLauncher(): void {
    this.previewLauncherElement?.remove();
    this.previewLauncherElement = undefined;
  }

  private updatePreviewLauncherText(): void {
    if (!this.previewLauncherElement) {
      return;
    }

    const label = getPreviewLauncherLabel(this.selectedFile);
    this.previewLauncherElement.textContent = label;
    this.previewLauncherElement.title = `${label} with the SharePoint preview viewer`;
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

  private async loadPreviewSettings(): Promise<void> {
    this.previewSettings = await this.getSettingsService().getSettings(DEFAULT_APP_BASE_URL);
    this.updateSelectedFilePreviewLauncher();
  }

  private async openSelectedFilePreview(): Promise<void> {
    await this.loadPreviewSettings().catch(() => undefined);
    this.selectedFile = this.getSelectedFile(false);

    if (!this.selectedFile) {
      this.updateSelectedFilePreviewLauncher();
      throw new Error('Select one enabled file type, then try File Preview app again.');
    }

    const dialog =
      this.selectedFile.extensionSettings.renderer === 'diagrams-net-embed'
        ? new DrawioViewerDialog(
            this.context.spHttpClient,
            this.context.pageContext.web.absoluteUrl,
            this.selectedFile.serverRelativeUrl,
            this.selectedFile.fileName,
            this.selectedFile.extensionSettings
          )
        : new BpmnViewerDialog(
            this.context.spHttpClient,
            this.context.pageContext.web.absoluteUrl,
            this.selectedFile.serverRelativeUrl,
            this.selectedFile.fileName,
            this.selectedFile.extensionSettings
          );

    await dialog.show();
  }

  private startSelectedFileWatcher(): void {
    this.updateSelectedFilePreviewLauncher();
    this.urlWatchHandle = window.setInterval(() => {
      this.updateSelectedFilePreviewLauncher();
    }, 750);
  }

  private updateSelectedFilePreviewLauncher(): void {
    const selectedFile = this.getSelectedFile(true);
    this.selectedFile = selectedFile;

    if (selectedFile) {
      this.renderPreviewLauncher();
      this.updatePreviewLauncherText();
    } else {
      this.removePreviewLauncher();
    }
  }

  private getSelectedFile(allowCandidateFallback: boolean):
    | { extensionSettings: IFileExtensionSettings; fileName: string; serverRelativeUrl: string }
    | undefined {
    return this.getSelectedFileFromUrl(allowCandidateFallback) || this.getSelectedFileFromPage(allowCandidateFallback);
  }

  private getSelectedFileFromUrl(allowCandidateFallback: boolean):
    | { extensionSettings: IFileExtensionSettings; fileName: string; serverRelativeUrl: string }
    | undefined {
    const query = new URLSearchParams(window.location.search);
    const id = query.get('id');
    if (!id || id.indexOf('.') === -1) {
      return undefined;
    }

    const serverRelativeUrl = normalizeServerRelativeUrl(id);
    if (!serverRelativeUrl) {
      return undefined;
    }

    const fileName = decodeURIComponent(serverRelativeUrl.split('/').pop() || '');
    const extensionSettings =
      findExtensionSettings(this.previewSettings, fileName) ||
      (allowCandidateFallback ? getCandidateExtensionSettings(fileName) : undefined);
    if (!extensionSettings) {
      return undefined;
    }

    return {
      extensionSettings,
      fileName,
      serverRelativeUrl
    };
  }

  private getSelectedFileFromPage(allowCandidateFallback: boolean):
    | { extensionSettings: IFileExtensionSettings; fileName: string; serverRelativeUrl: string }
    | undefined {
    const selectedRow = getSelectedDocumentRow();
    if (!selectedRow) {
      return undefined;
    }

    const fileName = getSupportedFileNameFromText(selectedRow.textContent || '');
    if (!fileName) {
      return undefined;
    }

    const serverRelativeUrl = this.getServerRelativeUrlForSelectedRow(selectedRow, fileName);
    if (!serverRelativeUrl) {
      return undefined;
    }

    const extensionSettings =
      findExtensionSettings(this.previewSettings, fileName) ||
      (allowCandidateFallback ? getCandidateExtensionSettings(fileName) : undefined);
    if (!extensionSettings) {
      return undefined;
    }

    return {
      extensionSettings,
      fileName,
      serverRelativeUrl
    };
  }

  private getServerRelativeUrlForSelectedRow(row: Element, fileName: string): string {
    const rowLinkUrl = getServerRelativeUrlFromSelectedRowLink(row);
    if (rowLinkUrl) {
      return rowLinkUrl;
    }

    const query = new URLSearchParams(window.location.search);
    const id = normalizeServerRelativeUrl(query.get('id') || '');
    if (id) {
      if (hasCandidateExtension(id)) {
        return id;
      }

      return `${id.replace(/\/+$/, '')}/${encodeURIComponent(fileName)}`;
    }

    const rootFolder = normalizeServerRelativeUrl(query.get('RootFolder') || '');
    if (rootFolder) {
      return `${rootFolder.replace(/\/+$/, '')}/${encodeURIComponent(fileName)}`;
    }

    return '';
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

  private isTenantRootWeb(): boolean {
    return normalizeBaseUrl(this.context.pageContext.web.absoluteUrl) === getTenantRootSiteUrl(this.context.pageContext.web.absoluteUrl);
  }

  private ensureLauncherStyles(): void {
    if (document.getElementById('bpf-launcher-styles')) {
      return;
    }

    const style = document.createElement('style');
    style.id = 'bpf-launcher-styles';
    style.textContent = `
      .bpf-file-preview-launcher {
        align-items: center;
        background: #0078d4;
        border: 1px solid #005a9e;
        bottom: 20px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.18);
        color: #ffffff;
        cursor: pointer;
        display: inline-flex;
        font: 600 13px "Segoe UI", Arial, sans-serif;
        gap: 8px;
        min-height: 40px;
        padding: 0 16px;
        position: fixed;
        right: 20px;
        z-index: 1000000;
      }
      .bpf-file-preview-launcher:focus,
      .bpf-file-preview-launcher:hover {
        background: #005a9e;
      }
    `;
    document.head.appendChild(style);
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

function normalizeServerRelativeUrl(value: string): string {
  if (!value) {
    return '';
  }

  const decoded = decodeURIComponent(value);
  if (!decoded.startsWith('/')) {
    return '';
  }

  return decoded;
}

function getSelectedDocumentRow(): Element | undefined {
  const rowSelectors = [
    '[data-automationid="DetailsRow"][aria-selected="true"]',
    '[role="row"][aria-selected="true"]',
    '[aria-selected="true"]'
  ];

  for (const selector of rowSelectors) {
    const rows = Array.from(document.querySelectorAll(selector));
    const selectedRow = rows.find((row) => getSupportedFileNameFromText(row.textContent || ''));
    if (selectedRow) {
      return selectedRow;
    }
  }

  const checkboxSelectors = [
    '[role="checkbox"][aria-checked="true"]',
    'input[type="checkbox"]:checked'
  ];

  for (const selector of checkboxSelectors) {
    const checkboxes = Array.from(document.querySelectorAll(selector));
    for (const checkbox of checkboxes) {
      const selectedRow = findContainingDocumentRow(checkbox);
      if (selectedRow && getSupportedFileNameFromText(selectedRow.textContent || '')) {
        return selectedRow;
      }
    }
  }

  return undefined;
}

function findContainingDocumentRow(element: Element): Element | undefined {
  const row = element.closest('[data-automationid="DetailsRow"], [role="row"], [data-list-index]');
  if (row) {
    return row;
  }

  let current: Element | null = element;
  for (let depth = 0; current && depth < 8; depth++) {
    if (getSupportedFileNameFromText(current.textContent || '')) {
      return current;
    }

    current = current.parentElement;
  }

  return undefined;
}

function getSupportedFileNameFromText(value: string): string {
  const normalized = value.replace(/\s+/g, ' ');
  const match = normalized.match(/([^\s\\/:*?"<>|]+\.(?:bpmn|drawio|jt|step))/i);
  return match?.[1] || '';
}

function getServerRelativeUrlFromSelectedRowLink(row: Element): string {
  const links = Array.from(row.querySelectorAll('a[href]')) as HTMLAnchorElement[];
  for (const link of links) {
    const href = link.getAttribute('href') || '';
    const serverRelativeUrl = getServerRelativeUrlFromHref(href);
    if (serverRelativeUrl && hasCandidateExtension(serverRelativeUrl)) {
      return serverRelativeUrl;
    }
  }

  return '';
}

function getServerRelativeUrlFromHref(href: string): string {
  if (!href) {
    return '';
  }

  try {
    const parsed = new URL(href, window.location.origin);
    const id = normalizeServerRelativeUrl(parsed.searchParams.get('id') || '');
    if (id) {
      return id;
    }

    const sourcedoc = normalizeServerRelativeUrl(parsed.searchParams.get('sourcedoc') || '');
    if (sourcedoc) {
      return sourcedoc;
    }

    return parsed.hostname === window.location.hostname ? decodeURIComponent(parsed.pathname) : '';
  } catch {
    return '';
  }
}

function hasCandidateExtension(value: string): boolean {
  const normalizedValue = value.toLowerCase();
  return CANDIDATE_EXTENSIONS.some((extension) => normalizedValue.endsWith(extension));
}

function getCandidateExtensionSettings(fileName: string): IFileExtensionSettings | undefined {
  const normalizedFileName = fileName.toLowerCase();
  const extension = CANDIDATE_EXTENSIONS.find((candidate) => normalizedFileName.endsWith(candidate));
  if (!extension) {
    return undefined;
  }

  return {
    displayName: `${extension.toUpperCase()} file`,
    enabled: true,
    extension,
    mode: extension === '.bpmn' || extension === '.drawio' ? 'modeler' : 'viewer',
    renderer: extension === '.drawio' ? 'diagrams-net-embed' : extension === '.bpmn' ? 'bpmn-js' : 'coming-soon'
  };
}

function getPreviewLauncherLabel(selectedFile: { extensionSettings: IFileExtensionSettings } | undefined): string {
  const extension = selectedFile?.extensionSettings.extension;
  if (extension === '.drawio') {
    return 'Preview DrawIO';
  }
  if (extension === '.bpmn') {
    return 'Preview BPMN';
  }

  return 'Preview file';
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
