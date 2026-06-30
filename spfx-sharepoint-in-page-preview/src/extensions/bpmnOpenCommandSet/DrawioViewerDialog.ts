import { SPHttpClient } from '@microsoft/sp-http';
import { BaseDialog, type IDialogConfiguration } from '@microsoft/sp-dialog';
import type { IFileExtensionSettings } from './previewSettings';
import { SharePointFileService, type ISharePointFileMetadata } from './sharePointFileService';
import { renderIcon } from '../../shared/icons';

const DRAWIO_EMBED_ORIGIN: string = 'https://embed.diagrams.net';
const DRAWIO_EMBED_URL: string = `${DRAWIO_EMBED_ORIGIN}/?embed=1&proto=json&spin=1&libraries=1&noSaveBtn=1&noExitBtn=1`;

type DrawioMessage = {
  event?: string;
  message?: string;
  xml?: string;
};

export class DrawioViewerDialog extends BaseDialog {
  private fileService: SharePointFileService;
  private iframeElement: HTMLIFrameElement | undefined;
  private isDirty: boolean = false;
  private metadata: ISharePointFileMetadata | undefined;
  private readonly onMessageBound = this.onMessage.bind(this);
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
    this.domElement.style.cssText = 'box-sizing:border-box;display:flex;flex-direction:column;height:100dvh;inset:0;overflow:hidden;position:fixed;width:100vw;z-index:2147483647;';
    this.makeFullViewport();
    window.requestAnimationFrame(() => this.makeFullViewport());
    window.setTimeout(() => this.makeFullViewport(), 300);
    this.domElement.innerHTML = `
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
            <button class="drawio-dialog__button" data-action="fit" type="button" aria-label="Fit to screen" title="Fit to screen">${renderIcon(
              'maximize'
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

    this.ensureStyles();
    this.iframeElement = this.domElement.querySelector('[data-role="frame"]') as HTMLIFrameElement | undefined;
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

  private fitDiagram(): void {
    // Reloading the current XML causes diagrams.net to reset the viewport and fit the diagram.
    window.setTimeout(() => this.sendLoadMessage(), 50);
  }

  private wireEvents(): void {
    this.domElement.querySelector('[data-action="reload"]')?.addEventListener('click', () => {
      this.load().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not reload file.'));
    });
    this.domElement.querySelector('[data-action="save"]')?.addEventListener('click', () => {
      this.save().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not save file.'));
    });
    this.domElement.querySelector('[data-action="download"]')?.addEventListener('click', () => this.download());
    this.domElement.querySelector('[data-action="fit"]')?.addEventListener('click', () => this.fitDiagram());
    this.domElement.querySelector('[data-action="fullscreen"]')?.addEventListener('click', () => {
      this.toggleFullscreen().catch((error: unknown) =>
        this.setError(error instanceof Error ? error.message : 'Could not open full screen.')
      );
    });
    document.addEventListener('fullscreenchange', () => this.updateFullscreenButton());
    this.domElement.querySelector('.drawio-dialog__close')?.addEventListener('click', () => {
      this.close().catch(() => undefined);
    });
  }

  private makeFullViewport(): void {
    // Fluent UI's dialog open animation applies CSS transform to ancestor elements,
    // making them a new containing block for position:fixed children (CSS spec).
    // Fix: clear all containing-block-creating properties on every ancestor, then
    // measure the raw offset and apply an inverse translate to reach viewport origin.
    let parent: HTMLElement | null = this.domElement.parentElement;
    while (parent && parent !== document.body) {
      parent.style.setProperty('animation', 'none', 'important');
      parent.style.setProperty('transition', 'none', 'important');
      parent.style.setProperty('transform', 'none', 'important');
      parent.style.setProperty('will-change', 'auto', 'important');
      parent.style.setProperty('filter', 'none', 'important');
      parent.style.setProperty('perspective', 'none', 'important');
      parent.style.setProperty('contain', 'none', 'important');
      parent.style.setProperty('max-width', 'none', 'important');
      parent.style.setProperty('max-height', 'none', 'important');
      parent.style.setProperty('overflow', 'visible', 'important');
      parent.style.setProperty('border-radius', '0', 'important');
      parent = parent.parentElement;
    }

    const el = this.domElement;
    el.style.setProperty('position', 'fixed', 'important');
    el.style.setProperty('inset', '0', 'important');
    el.style.setProperty('width', '100vw', 'important');
    el.style.setProperty('height', '100dvh', 'important');
    el.style.setProperty('z-index', '2147483647', 'important');

    // Remove any prior compensation transform so getBoundingClientRect reflects
    // the raw containing-block offset (not the already-translated position).
    // removeProperty + getBoundingClientRect are synchronous; no paint occurs between.
    el.style.removeProperty('transform');
    const rect = el.getBoundingClientRect();
    if (rect.left !== 0 || rect.top !== 0) {
      el.style.setProperty('transform', `translate(${-rect.left}px,${-rect.top}px)`, 'important');
    }
  }

  private ensureStyles(): void {
    if (this.domElement.querySelector('style[data-bpf-preview-style="drawio"]')) {
      return;
    }

    const style = document.createElement('style');
    style.dataset.bpfPreviewStyle = 'drawio';
    style.textContent = `
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
      :fullscreen .drawio-dialog {
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
    this.domElement.appendChild(style);
  }

  private async toggleFullscreen(): Promise<void> {
    if (document.fullscreenElement) {
      await this.exitFullscreen();
      return;
    }

    await this.domElement.requestFullscreen();
    this.updateFullscreenButton();
  }

  private async exitFullscreen(): Promise<void> {
    if (document.fullscreenElement) {
      await document.exitFullscreen();
      this.updateFullscreenButton();
    }
  }

  private updateFullscreenButton(): void {
    const button = this.domElement.querySelector('[data-action="fullscreen"]') as HTMLButtonElement | null;
    if (!button) {
      return;
    }

    const isFullscreen = Boolean(document.fullscreenElement);
    button.innerHTML = renderIcon(isFullscreen ? 'restore' : 'external');
    button.setAttribute('aria-label', isFullscreen ? 'Exit full screen' : 'Open full screen');
    button.title = isFullscreen ? 'Exit full screen' : 'Open full screen';

    // Reload the diagram into the expanded viewport so diagrams.net auto-fits on entry.
    if (isFullscreen) {
      window.setTimeout(() => this.sendLoadMessage(), 200);
    }
  }

  private isEditable(): boolean {
    return this.extensionSettings.mode === 'modeler';
  }

  private setBusy(isBusy: boolean, status: string): void {
    this.setStatus(status);
    this.domElement.querySelectorAll('.drawio-dialog__button').forEach((button) => {
      const typedButton = button as HTMLButtonElement;
      if (typedButton.dataset.action === 'save') {
        typedButton.disabled = isBusy || !this.isEditable();
      } else {
        typedButton.disabled = isBusy;
      }
    });
  }

  private setStatus(status: string): void {
    const statusElement = this.domElement.querySelector('[data-role="status"]') as HTMLElement | null;
    if (statusElement) {
      statusElement.textContent = status;
    }
  }

  private setMessage(message: string): void {
    const messageElement = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.hidden = message.length === 0;
    messageElement.textContent = message;
    messageElement.classList.remove('drawio-dialog__message--error');
  }

  private setError(message: string): void {
    const messageElement = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.hidden = false;
    messageElement.textContent = message;
    messageElement.classList.add('drawio-dialog__message--error');
  }

  private renderMetadata(): void {
    const nameElement = this.domElement.querySelector('.drawio-dialog__name') as HTMLElement | null;
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
    const saveButton = this.domElement.querySelector('[data-action="save"]') as HTMLButtonElement | null;
    if (saveButton) {
      saveButton.disabled = !this.isEditable();
    }
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
