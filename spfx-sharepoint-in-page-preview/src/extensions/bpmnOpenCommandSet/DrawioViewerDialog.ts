import { SPHttpClient } from '@microsoft/sp-http';
import { BaseDialog, type IDialogConfiguration } from '@microsoft/sp-dialog';
import type { IFileExtensionSettings } from './previewSettings';
import { SharePointFileService, type ISharePointFileMetadata } from './sharePointFileService';

const DRAWIO_EMBED_ORIGIN: string = 'https://embed.diagrams.net';
const DRAWIO_EMBED_URL: string = `${DRAWIO_EMBED_ORIGIN}/?embed=1&proto=json&spin=1&libraries=1&noSaveBtn=1&noExitBtn=1`;

type DrawioMessage = {
  event?: string;
  message?: string;
  xml?: string;
};

export class DrawioViewerDialog extends BaseDialog {
  private fileService: SharePointFileService;
  private hostElement: HTMLElement | undefined;
  private iframeElement: HTMLIFrameElement | undefined;
  private isDirty: boolean = false;
  private metadata: ISharePointFileMetadata | undefined;
  private readonly onMessageBound = this.onMessage.bind(this);
  private readonly onWindowResizeBound = (): void => this.applyDialogChrome();
  private xml: string = '';

  public constructor(
    spHttpClient: SPHttpClient,
    webAbsoluteUrl: string,
    private readonly serverRelativeUrl: string,
    private readonly fileName: string,
    private readonly extensionSettings: IFileExtensionSettings
  ) {
    super({ isBlocking: false });
    this.fileService = new SharePointFileService(spHttpClient, webAbsoluteUrl);
  }

  public render(): void {
    const root = this.ensureHostElement();
    root.innerHTML = `
      <div class="drawio-dialog">
        <div class="drawio-dialog__header">
          <div class="drawio-dialog__title">
            <span class="drawio-dialog__badge">DRAWIO</span>
            <span class="drawio-dialog__name" title="${escapeHtml(this.fileName)}">${escapeHtml(this.fileName)}</span>
            <span class="drawio-dialog__status" data-role="status">Loading</span>
          </div>
          <div class="drawio-dialog__actions">
            <button class="drawio-dialog__button" data-action="reload" type="button" aria-label="Reload" title="Reload">${renderIcon(
              'refresh'
            )}</button>
            <button class="drawio-dialog__button" data-action="save" disabled type="button" aria-label="Save" title="Save">${renderIcon(
              'save'
            )}</button>
            <button class="drawio-dialog__button" data-action="download" type="button" aria-label="Download" title="Download">${renderIcon(
              'download'
            )}</button>
            <button class="drawio-dialog__button" data-action="fullscreen" type="button" aria-label="Open full screen" title="Open full screen">${renderIcon(
              'external'
            )}</button>
            <button class="drawio-dialog__close" type="button" aria-label="Close preview" title="Close">&times;</button>
          </div>
        </div>
        <div class="drawio-dialog__message" data-role="message" hidden></div>
        <iframe class="drawio-dialog__frame" data-role="frame" title="diagrams.net preview" sandbox="allow-downloads allow-forms allow-popups allow-same-origin allow-scripts" src="${DRAWIO_EMBED_URL}"></iframe>
      </div>
    `;

    this.applyDialogChrome();
    this.scheduleChromeRefresh();
    this.iframeElement = this.rootElement.querySelector('[data-role="frame"]') as HTMLIFrameElement | undefined;
    this.wireEvents();
    window.addEventListener('message', this.onMessageBound);
    this.load().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not open draw.io file.'));
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  protected onAfterClose(): void {
    this.exitFullscreen().catch(() => undefined);
    window.removeEventListener('message', this.onMessageBound);
    window.removeEventListener('resize', this.onWindowResizeBound);
    this.hostElement?.remove();
    this.hostElement = undefined;
    this.domElement.innerHTML = '';
    super.onAfterClose();
  }

  private async load(): Promise<void> {
    this.setBusy(true, 'Loading');
    this.setMessage('');
    this.isDirty = false;
    this.updateSaveButton();

    const [metadata, xml] = await Promise.all([
      this.fileService.getMetadata(this.serverRelativeUrl),
      this.fileService.getContent(this.serverRelativeUrl)
    ]);
    this.metadata = metadata;
    this.xml = xml;
    this.renderMetadata();
    this.setBusy(false, 'Ready');
    this.sendLoadMessage();
  }

  private async save(): Promise<void> {
    this.requestSaveXml();
  }

  private async saveXml(xml: string): Promise<void> {
    this.setBusy(true, 'Saving');
    try {
      this.xml = xml;
      this.metadata = await this.fileService.saveContent(this.serverRelativeUrl, this.xml, this.metadata?.eTag);
      this.isDirty = false;
      this.renderMetadata();
      this.setBusy(false, 'Saved');
      this.updateSaveButton();
    } catch (error) {
      this.setBusy(false, 'Error');
      this.setError(error instanceof Error ? error.message : 'Could not save draw.io file.');
    }
  }

  private download(): void {
    const blob = new Blob([this.xml], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = this.fileName || 'diagram.drawio';
    anchor.click();
    URL.revokeObjectURL(url);
  }

  private onMessage(event: MessageEvent): void {
    if (event.origin !== DRAWIO_EMBED_ORIGIN || !this.iframeElement || event.source !== this.iframeElement.contentWindow) {
      return;
    }

    const message = parseDrawioMessage(event.data);
    if (!message) {
      return;
    }

    switch (message.event) {
      case 'init':
        this.sendLoadMessage();
        break;
      case 'autosave':
      case 'save':
        if (typeof message.xml === 'string') {
          this.xml = message.xml;
          this.isDirty = true;
          this.setStatus('Unsaved');
          this.updateSaveButton();
          if (this.isEditable() && message.event === 'save') {
            this.saveXml(message.xml).catch((error: unknown) =>
              this.setError(error instanceof Error ? error.message : 'Could not save draw.io file.')
            );
          }
        }
        break;
      case 'exit':
        this.close().catch(() => undefined);
        break;
      default:
        if (message.message) {
          this.setMessage(message.message);
        }
        break;
    }
  }

  private sendLoadMessage(): void {
    if (!this.iframeElement?.contentWindow || !this.xml) {
      return;
    }

    this.iframeElement.contentWindow.postMessage(
      JSON.stringify({
        action: 'load',
        autosave: this.isEditable() ? 1 : 0,
        modified: 'unsavedChanges',
        saveAndExit: '0',
        title: this.fileName,
        xml: this.xml
      }),
      DRAWIO_EMBED_ORIGIN
    );
  }

  private requestSaveXml(): void {
    this.iframeElement?.contentWindow?.postMessage(JSON.stringify({ action: 'save' }), DRAWIO_EMBED_ORIGIN);
  }

  private wireEvents(): void {
    this.rootElement.querySelector('[data-action="reload"]')?.addEventListener('click', () => {
      this.load().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not reload file.'));
    });
    this.rootElement.querySelector('[data-action="save"]')?.addEventListener('click', () => {
      this.save().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not save file.'));
    });
    this.rootElement.querySelector('[data-action="download"]')?.addEventListener('click', () => this.download());
    this.rootElement.querySelector('[data-action="fullscreen"]')?.addEventListener('click', () => {
      this.toggleFullscreen().catch((error: unknown) =>
        this.setError(error instanceof Error ? error.message : 'Could not open full screen.')
      );
    });
    document.addEventListener('fullscreenchange', () => this.updateFullscreenButton());
    this.rootElement.querySelector('.drawio-dialog__close')?.addEventListener('click', () => {
      this.close().catch(() => undefined);
    });
    window.addEventListener('resize', this.onWindowResizeBound);
  }

  private applyDialogChrome(): void {
    setImportantStyle(this.rootElement, {
      background: '#ffffff',
      bottom: '0',
      boxSizing: 'border-box',
      height: 'auto',
      inset: '0',
      left: '0',
      margin: '0',
      maxHeight: 'none',
      maxWidth: 'none',
      minHeight: '0',
      minWidth: '0',
      overflow: 'hidden',
      position: 'fixed',
      right: '0',
      top: '0',
      width: 'auto',
      zIndex: '1000002'
    });
    tagPreviewAncestors(this.domElement);

    const modal = this.domElement.closest('.ms-Modal') as HTMLElement | null;
    if (modal) {
      setImportantStyle(modal, {
        bottom: '0',
        display: 'block',
        height: '100%',
        inset: '0',
        maxHeight: '100%',
        maxWidth: '100%',
        minHeight: '100%',
        minWidth: '100%',
        overflow: 'hidden',
        position: 'fixed',
        right: '0',
        top: '0',
        width: '100%'
      });
    }

    const focusTrap = this.domElement.closest('[data-is-focus-trap-zone="true"]') as HTMLElement | null;
    if (focusTrap) {
      setImportantStyle(focusTrap, {
        height: '100%',
        maxHeight: '100%',
        maxWidth: '100%',
        minHeight: '100%',
        minWidth: '100%',
        width: '100%'
      });
    }

    const dialogMain = this.domElement.closest('.ms-Dialog-main') as HTMLElement | null;
    if (dialogMain) {
      dialogMain.classList.add('bpf-preview-dialog-main');
      setImportantStyle(dialogMain, {
        borderRadius: '0',
        bottom: '0',
        boxSizing: 'border-box',
        height: 'auto',
        inset: '0',
        left: '0',
        margin: '0',
        maxHeight: 'none',
        maxWidth: 'none',
        minHeight: '0',
        minWidth: '0',
        overflow: 'hidden',
        position: 'fixed',
        right: '0',
        top: '0',
        transform: 'none',
        width: 'auto'
      });
    }

    const dialogInner = this.domElement.closest('.ms-Dialog-inner') as HTMLElement | null;
    if (dialogInner) {
      setImportantStyle(dialogInner, {
        height: '100%',
        maxHeight: '100%',
        padding: '0'
      });
    }

    const dialogContent = this.domElement.closest('.ms-Dialog-content') as HTMLElement | null;
    if (dialogContent) {
      setImportantStyle(dialogContent, {
        height: '100%',
        maxHeight: '100%'
      });
    }

    const modalScroll = this.domElement.closest('.ms-Modal-scrollableContent') as HTMLElement | null;
    if (modalScroll) {
      setImportantStyle(modalScroll, {
        height: '100%',
        maxHeight: '100%',
        overflow: 'hidden'
      });
    }

    const layerContent = this.domElement.closest('.ms-Layer-content') as HTMLElement | null;
    if (layerContent) {
      setImportantStyle(layerContent, {
        bottom: '0',
        height: '100%',
        inset: '0',
        left: '0',
        maxHeight: '100%',
        maxWidth: '100%',
        overflow: 'hidden',
        position: 'fixed',
        right: '0',
        top: '0',
        width: '100%'
      });
    }

    this.domElement.style.display = 'none';
    applyPreviewOverlayChrome();

    if (this.rootElement.querySelector('style[data-bpf-preview-style="drawio"]')) {
      return;
    }

    const style = document.createElement('style');
    style.dataset.bpfPreviewStyle = 'drawio';
    style.textContent = `
      .ms-Layer [data-bpf-preview-layer="true"],
      .ms-Layer [data-bpf-preview-modal="true"],
      .ms-Layer [data-bpf-preview-focus-trap="true"] {
        bottom: 0 !important;
        height: 100% !important;
        inset: 0 !important;
        left: 0 !important;
        margin: 0 !important;
        max-height: none !important;
        max-width: none !important;
        min-height: 0 !important;
        min-width: 0 !important;
        overflow: hidden !important;
        position: fixed !important;
        right: 0 !important;
        top: 0 !important;
        transform: none !important;
        width: auto !important;
      }
      .ms-Layer .ms-Dialog-main.bpf-preview-dialog-main {
        bottom: 0 !important;
        box-sizing: border-box !important;
        border-radius: 0 !important;
        height: auto !important;
        inset: 0 !important;
        left: 0 !important;
        margin: 0 !important;
        max-height: none !important;
        max-width: none !important;
        min-height: 0 !important;
        min-width: 0 !important;
        overflow: hidden !important;
        position: fixed !important;
        right: 0 !important;
        top: 0 !important;
        transform: none !important;
        width: auto !important;
      }
      .ms-Layer .ms-Layer-content:has(.bpf-preview-dialog-main) {
        bottom: 0 !important;
        height: auto !important;
        inset: 0 !important;
        left: 0 !important;
        max-height: none !important;
        max-width: none !important;
        overflow: hidden !important;
        position: fixed !important;
        right: 0 !important;
        top: 0 !important;
        width: auto !important;
      }
      .ms-Layer .ms-Dialog-main.bpf-preview-dialog-main .ms-Dialog-inner,
      .ms-Layer .ms-Dialog-main.bpf-preview-dialog-main .ms-Dialog-content,
      .ms-Layer .ms-Dialog-main.bpf-preview-dialog-main .ms-Modal-scrollableContent {
        height: 100% !important;
        max-height: 100% !important;
        overflow: hidden !important;
        padding: 0 !important;
      }
      .bpf-preview-dialog-main:fullscreen {
        height: 100dvh !important;
        max-height: 100dvh !important;
        max-width: 100dvw !important;
        width: 100dvw !important;
      }
      .drawio-dialog {
        background: #ffffff;
        color: #f5f5f5;
        display: flex;
        flex-direction: column;
        font-family: "Segoe UI", Arial, sans-serif;
        height: 100%;
        min-height: 0;
        overflow: hidden;
        width: 100%;
      }
      .drawio-dialog__header {
        align-items: center;
        background: #1b1b1b;
        border-bottom: 1px solid rgba(255,255,255,.12);
        display: flex;
        flex: 0 0 52px;
        gap: 16px;
        justify-content: space-between;
        min-height: 52px;
        padding: 6px 12px 6px 16px;
      }
      .drawio-dialog__title,
      .drawio-dialog__actions {
        align-items: center;
        display: flex;
        gap: 8px;
        min-width: 0;
      }
      .drawio-dialog__title {
        flex: 1 1 auto;
      }
      .drawio-dialog__actions {
        flex: 0 0 auto;
        flex-wrap: nowrap;
        justify-content: flex-end;
      }
      .drawio-dialog__badge {
        background: #242424;
        border: 1px solid rgba(255,255,255,.18);
        border-radius: 4px;
        color: #d8ebff;
        flex: 0 0 auto;
        font-size: 13px;
        font-weight: 700;
        padding: 6px 8px;
      }
      .drawio-dialog__name {
        font-size: 15px;
        font-weight: 600;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .drawio-dialog__status {
        color: #a6a6a6;
        font-size: 12px;
      }
      .drawio-dialog__button,
      .drawio-dialog__close {
        align-items: center;
        background: #242424;
        border: 1px solid rgba(255,255,255,.2);
        color: #ffffff;
        cursor: pointer;
        display: inline-flex;
        font: inherit;
        justify-content: center;
      }
      .drawio-dialog__button {
        border-radius: 4px;
        height: 36px;
        min-width: 36px;
        padding: 0;
        width: 36px;
      }
      .drawio-dialog__button svg {
        display: block;
        height: 18px;
        width: 18px;
      }
      .drawio-dialog__button:hover:not(:disabled),
      .drawio-dialog__close:hover {
        background: rgba(255,255,255,.1);
      }
      .drawio-dialog__button:disabled {
        color: #777777;
        cursor: not-allowed;
      }
      .drawio-dialog__close {
        border-radius: 4px;
        font-size: 24px;
        height: 36px;
        line-height: 1;
        padding: 0 0 3px;
        width: 36px;
      }
      .drawio-dialog__message {
        flex: 0 0 auto;
        border-bottom: 1px solid rgba(255,255,255,.12);
        color: #f8d66d;
        font-size: 13px;
        padding: 8px 20px;
      }
      .drawio-dialog__message--error {
        background: #fde7e9;
        color: #7b1f23;
      }
      .drawio-dialog__frame {
        background: #ffffff;
        border: 0;
        flex: 1 1 auto;
        height: auto;
        min-height: 0;
        width: 100%;
      }
      .bpf-preview-dialog-main:fullscreen .drawio-dialog {
        min-height: 100dvh;
      }
      @media (max-width: 720px) {
        .drawio-dialog__header {
          align-items: stretch;
          flex-direction: column;
          flex-basis: auto;
        }
        .drawio-dialog__actions {
          justify-content: flex-start;
        }
      }
    `;
    this.rootElement.prepend(style);
  }

  private scheduleChromeRefresh(): void {
    window.requestAnimationFrame(() => this.applyDialogChrome());
    window.setTimeout(() => this.applyDialogChrome(), 150);
  }

  private async toggleFullscreen(): Promise<void> {
    const host = this.rootElement;
    if (!host) {
      return;
    }

    if (document.fullscreenElement) {
      await this.exitFullscreen();
      return;
    }

    await host.requestFullscreen();
    this.updateFullscreenButton();
  }

  private async exitFullscreen(): Promise<void> {
    if (document.fullscreenElement) {
      await document.exitFullscreen();
      this.updateFullscreenButton();
    }
  }

  private updateFullscreenButton(): void {
    const button = this.rootElement.querySelector('[data-action="fullscreen"]') as HTMLButtonElement | null;
    if (!button) {
      return;
    }

    const isFullscreen = Boolean(document.fullscreenElement);
    button.innerHTML = renderIcon(isFullscreen ? 'restore' : 'external');
    button.setAttribute('aria-label', isFullscreen ? 'Exit full screen' : 'Open full screen');
    button.title = isFullscreen ? 'Exit full screen' : 'Open full screen';
  }

  private isEditable(): boolean {
    return true;
  }

  private setBusy(isBusy: boolean, status: string): void {
    this.setStatus(status);
    this.rootElement.querySelectorAll('.drawio-dialog__button').forEach((button) => {
      const typedButton = button as HTMLButtonElement;
      if (typedButton.dataset.action === 'save') {
        typedButton.disabled = isBusy || !this.isEditable();
      } else {
        typedButton.disabled = isBusy;
      }
    });
  }

  private setStatus(status: string): void {
    const statusElement = this.rootElement.querySelector('[data-role="status"]') as HTMLElement | null;
    if (statusElement) {
      statusElement.textContent = status;
    }
  }

  private setMessage(message: string): void {
    const messageElement = this.rootElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.hidden = message.length === 0;
    messageElement.textContent = message;
    messageElement.classList.remove('drawio-dialog__message--error');
  }

  private setError(message: string): void {
    const messageElement = this.rootElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.hidden = false;
    messageElement.textContent = message;
    messageElement.classList.add('drawio-dialog__message--error');
  }

  private renderMetadata(): void {
    const nameElement = this.rootElement.querySelector('.drawio-dialog__name') as HTMLElement | null;
    if (!nameElement || !this.metadata) {
      return;
    }

    const parts: string[] = [];
    if (this.metadata.timeLastModified) {
      parts.push(`Modified ${new Date(this.metadata.timeLastModified).toLocaleString()}`);
    }
    if (this.metadata.length) {
      parts.push(formatBytes(Number(this.metadata.length)));
    }
    parts.push('External diagrams.net renderer');
    nameElement.title = `${this.fileName}\n${parts.join(' | ')}`;
  }

  private updateSaveButton(): void {
    const saveButton = this.rootElement.querySelector('[data-action="save"]') as HTMLButtonElement | null;
    if (saveButton) {
      saveButton.disabled = !this.isEditable();
    }
  }

  private get rootElement(): HTMLElement {
    return this.hostElement || this.domElement;
  }

  private ensureHostElement(): HTMLElement {
    if (this.hostElement) {
      return this.hostElement;
    }

    const hostElement = document.createElement('div');
    hostElement.className = 'bpf-preview-portal bpf-preview-portal--drawio';
    document.body.appendChild(hostElement);
    this.hostElement = hostElement;
    this.domElement.innerHTML = '';
    this.domElement.style.display = 'none';
    return hostElement;
  }
}

function parseDrawioMessage(value: unknown): DrawioMessage | undefined {
  if (typeof value === 'object' && value !== null) {
    return value as DrawioMessage;
  }

  if (typeof value !== 'string') {
    return undefined;
  }

  try {
    return JSON.parse(value) as DrawioMessage;
  } catch {
    return undefined;
  }
}

function formatBytes(value: number): string {
  if (!Number.isFinite(value)) {
    return '';
  }

  if (value < 1024) {
    return `${value} B`;
  }

  if (value < 1024 * 1024) {
    return `${Math.round(value / 102.4) / 10} KB`;
  }

  return `${Math.round(value / 1024 / 102.4) / 10} MB`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function renderIcon(name: string): string {
  const paths: Record<string, string> = {
    download: '<path d="M12 3v12" /><path d="M7 10l5 5 5-5" /><path d="M5 21h14" />',
    external: '<path d="M14 3h7v7" /><path d="M21 3l-9 9" /><path d="M19 14v5H5V5h5" />',
    refresh: '<path d="M20 6v5h-5" /><path d="M4 18v-5h5" /><path d="M18 9a6 6 0 0 0-10-3L4 10" /><path d="M6 15a6 6 0 0 0 10 3l4-4" />',
    restore: '<path d="M9 3H4v5" /><path d="M4 3l7 7" /><path d="M15 21h5v-5" /><path d="M20 21l-7-7" />',
    save: '<path d="M5 3h12l2 2v16H5z" /><path d="M8 3v6h8" /><path d="M8 21v-7h8v7" />'
  };

  return `<svg aria-hidden="true" focusable="false" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${paths[name] || ''}</svg>`;
}

function tagPreviewAncestors(root: HTMLElement): void {
  (root.closest('.ms-Layer-content') as HTMLElement | null)?.setAttribute('data-bpf-preview-layer', 'true');
  (root.closest('.ms-Modal') as HTMLElement | null)?.setAttribute('data-bpf-preview-modal', 'true');
  (root.closest('[data-is-focus-trap-zone="true"]') as HTMLElement | null)?.setAttribute(
    'data-bpf-preview-focus-trap',
    'true'
  );
}

function setImportantStyle(element: HTMLElement, styles: Record<string, string>): void {
  Object.keys(styles).forEach((propertyName) => {
    element.style.setProperty(toCssPropertyName(propertyName), styles[propertyName], 'important');
  });
}

function toCssPropertyName(propertyName: string): string {
  return propertyName.replace(/[A-Z]/g, (match) => `-${match.toLowerCase()}`);
}

function applyPreviewOverlayChrome(): void {
  const overlays = Array.from(document.querySelectorAll('.ms-Overlay, [data-is-focus-trap-zone] + .ms-Overlay')) as HTMLElement[];
  overlays.forEach((overlay) => {
    overlay.style.backgroundColor = 'rgba(0, 0, 0, 0.72)';
  });
}
