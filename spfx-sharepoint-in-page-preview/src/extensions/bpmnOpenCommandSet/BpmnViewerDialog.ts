import { SPHttpClient } from '@microsoft/sp-http';
import { BaseDialog, type IDialogConfiguration } from '@microsoft/sp-dialog';
import BpmnModeler from 'bpmn-js/lib/Modeler';
import BpmnViewer from 'bpmn-js/lib/NavigatedViewer';
import { ensureBpmnAssetStyles } from './bpmnAssetStyles';
import type { IFileExtensionSettings } from './previewSettings';
import { SharePointFileService, type ISharePointFileMetadata } from './sharePointFileService';
import { renderIcon } from '../../shared/icons';

type BpmnInstance = BpmnModeler | BpmnViewer;

export class BpmnViewerDialog extends BaseDialog {
  private bpmnInstance: BpmnInstance | undefined;
  private canvasElement: HTMLElement | undefined;
  private fileService: SharePointFileService;
  private isDirty: boolean = false;
  private metadata: ISharePointFileMetadata | undefined;
  private resizeObserver: ResizeObserver | undefined;
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
    ensureBpmnAssetStyles();

    this.domElement.style.cssText = 'box-sizing:border-box;display:flex;flex-direction:column;height:100dvh;inset:0;overflow:hidden;position:fixed;width:100vw;z-index:2147483647;';
    this.makeFullViewport();
    window.requestAnimationFrame(() => this.makeFullViewport());
    window.setTimeout(() => this.makeFullViewport(), 300);
    this.domElement.innerHTML = `
      <div class="bpmn-dialog">
        <div class="bpmn-dialog__header">
          <div class="bpmn-dialog__title">
            <span class="bpmn-dialog__badge">${escapeHtml(this.extensionSettings.extension.replace('.', '').toUpperCase())}</span>
            <span class="bpmn-dialog__name" title="${escapeHtml(this.fileName)}">${escapeHtml(this.fileName)}</span>
            <span class="bpmn-dialog__status" data-role="status">Loading</span>
          </div>
          <div class="bpmn-dialog__actions">
            <button class="bpmn-dialog__button" data-action="reload" type="button" aria-label="Reload" title="Reload">${renderIcon(
              'refresh'
            )}</button>
            <button class="bpmn-dialog__button" data-action="save" disabled type="button" aria-label="Save" title="Save">${renderIcon(
              'save'
            )}</button>
            <button class="bpmn-dialog__button" data-action="validate" type="button" aria-label="Validate" title="Validate">${renderIcon(
              'check'
            )}</button>
            <button class="bpmn-dialog__button" data-action="download" type="button" aria-label="Download" title="Download">${renderIcon(
              'download'
            )}</button>
            <button class="bpmn-dialog__button" data-action="fullscreen" type="button" aria-label="Open full screen" title="Open full screen">${renderIcon(
              'external'
            )}</button>
            <button class="bpmn-dialog__button" data-action="zoom-out" type="button" aria-label="Zoom out" title="Zoom out">${renderIcon(
              'zoomOut'
            )}</button>
            <button class="bpmn-dialog__button" data-action="zoom-in" type="button" aria-label="Zoom in" title="Zoom in">${renderIcon(
              'zoomIn'
            )}</button>
            <button class="bpmn-dialog__button" data-action="fit" type="button" aria-label="Fit to screen" title="Fit to screen">${renderIcon(
              'maximize'
            )}</button>
            <button class="bpmn-dialog__close" type="button" aria-label="Close preview" title="Close">&times;</button>
          </div>
        </div>
        <div class="bpmn-dialog__message" data-role="message" hidden></div>
        <div class="bpmn-dialog__canvas" data-role="canvas"></div>
      </div>
    `;

    this.ensureStyles();
    this.canvasElement = this.domElement.querySelector('[data-role="canvas"]') as HTMLElement | undefined;
    this.wireEvents();
    this.load().catch((error: unknown) => {
      this.setError(error instanceof Error ? error.message : 'Could not open the selected file.');
    });
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  protected onAfterClose(): void {
    this.exitFullscreen().catch(() => undefined);
    this.resizeObserver?.disconnect();
    this.resizeObserver = undefined;
    this.destroyRenderer();
    super.onAfterClose();
  }

  private async load(): Promise<void> {
    this.setBusy(true, 'Loading');
    this.setMessage('');
    this.destroyRenderer();
    this.isDirty = false;
    this.updateSaveButton();

    const [metadata, xml] = await Promise.all([
      this.fileService.getMetadata(this.serverRelativeUrl),
      this.fileService.getContent(this.serverRelativeUrl)
    ]);
    this.metadata = metadata;
    this.xml = xml;
    this.renderMetadata();
    await this.renderDiagram(xml);
    this.setBusy(false, this.isEditable() ? 'Ready' : 'Preview');
  }

  private async renderDiagram(xml: string): Promise<void> {
    const container = this.canvasElement;
    if (!container) {
      throw new Error('The BPMN canvas was not created.');
    }

    container.innerHTML = '';
    this.bpmnInstance = this.isEditable()
      ? new BpmnModeler({
          container,
          keyboard: { bindTo: window }
        })
      : new BpmnViewer({ container });

    const importResult = await this.bpmnInstance.importXML(xml);
    this.setWarnings(importResult.warnings);

    this.watchCanvasSize();
    this.fitDiagram();
    window.requestAnimationFrame(() => this.fitDiagram());
    window.setTimeout(() => this.fitDiagram(), 150);

    if (this.isEditable()) {
      const eventBus = this.bpmnInstance.get<{ on: (eventName: string, callback: () => void) => void }>('eventBus');
      eventBus.on('commandStack.changed', () => {
        this.isDirty = true;
        this.setStatus('Unsaved');
        this.updateSaveButton();
      });
    }
  }

  private async save(): Promise<void> {
    if (!this.bpmnInstance || !this.isEditable()) {
      return;
    }

    this.setBusy(true, 'Saving');
    this.setMessage('');

    try {
      const result = await this.bpmnInstance.saveXML({ format: true });
      this.xml = result.xml || '';
      this.metadata = await this.fileService.saveContent(this.serverRelativeUrl, this.xml, this.metadata?.eTag);
      this.isDirty = false;
      this.renderMetadata();
      this.updateSaveButton();
      this.setBusy(false, 'Saved');
    } catch (error) {
      this.setBusy(false, 'Error');
      this.setError(error instanceof Error ? error.message : 'Could not save the BPMN file.');
    }
  }

  private async validate(): Promise<void> {
    if (!this.bpmnInstance) {
      return;
    }

    try {
      const result = await this.bpmnInstance.saveXML({ format: true });
      if (!result.xml?.trim()) {
        throw new Error('The diagram XML is empty.');
      }

      this.setMessage('BPMN XML is valid for this renderer.');
      this.setStatus(this.isDirty ? 'Unsaved' : 'Valid');
    } catch (error) {
      this.setError(error instanceof Error ? error.message : 'Could not validate BPMN XML.');
      this.setStatus('Error');
    }
  }

  private async download(): Promise<void> {
    const result = this.bpmnInstance ? await this.bpmnInstance.saveXML({ format: true }) : { xml: this.xml };
    const blob = new Blob([result.xml || this.xml], { type: 'application/xml' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = this.fileName || 'diagram.bpmn';
    anchor.click();
    URL.revokeObjectURL(url);
  }

  private zoom(value: string | number): void {
    const canvas = this.bpmnInstance?.get<{ zoom: (value: string | number, point?: string) => void }>('canvas');
    canvas?.zoom(value, 'auto');
  }

  private fitDiagram(): void {
    const canvas = this.bpmnInstance?.get<{
      resized?: () => void;
      zoom: (value: string | number, point?: string) => void;
    }>('canvas');
    canvas?.resized?.();
    canvas?.zoom('fit-viewport', 'auto');
  }

  private watchCanvasSize(): void {
    this.resizeObserver?.disconnect();
    if (!this.canvasElement || typeof ResizeObserver === 'undefined') {
      return;
    }

    this.resizeObserver = new ResizeObserver(() => this.fitDiagram());
    this.resizeObserver.observe(this.canvasElement);
  }

  private wireEvents(): void {
    this.domElement.querySelector('[data-action="reload"]')?.addEventListener('click', () => {
      this.load().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not reload file.'));
    });
    this.domElement.querySelector('[data-action="save"]')?.addEventListener('click', () => {
      this.save().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not save file.'));
    });
    this.domElement.querySelector('[data-action="validate"]')?.addEventListener('click', () => {
      this.validate().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not validate file.'));
    });
    this.domElement.querySelector('[data-action="download"]')?.addEventListener('click', () => {
      this.download().catch((error: unknown) => this.setError(error instanceof Error ? error.message : 'Could not download file.'));
    });
    this.domElement.querySelector('[data-action="fullscreen"]')?.addEventListener('click', () => {
      this.toggleFullscreen().catch((error: unknown) =>
        this.setError(error instanceof Error ? error.message : 'Could not open full screen.')
      );
    });
    document.addEventListener('fullscreenchange', () => this.updateFullscreenButton());
    this.domElement.querySelector('[data-action="zoom-out"]')?.addEventListener('click', () => this.zoom(0.8));
    this.domElement.querySelector('[data-action="zoom-in"]')?.addEventListener('click', () => this.zoom(1.2));
    this.domElement.querySelector('[data-action="fit"]')?.addEventListener('click', () => this.zoom('fit-viewport'));
    this.domElement.querySelector('.bpmn-dialog__close')?.addEventListener('click', () => {
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
    if (this.domElement.querySelector('style[data-bpf-preview-style="bpmn"]')) {
      return;
    }

    const style = document.createElement('style');
    style.dataset.bpfPreviewStyle = 'bpmn';
    style.textContent = `
      .bpmn-dialog {
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
      .bpmn-dialog__header {
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
      .bpmn-dialog__title,
      .bpmn-dialog__actions {
        align-items: center;
        display: flex;
        gap: 8px;
        min-width: 0;
      }
      .bpmn-dialog__title {
        flex: 1 1 auto;
      }
      .bpmn-dialog__actions {
        flex: 0 0 auto;
        flex-wrap: nowrap;
        justify-content: flex-end;
      }
      .bpmn-dialog__badge {
        background: #242424;
        border: 1px solid rgba(255,255,255,.18);
        border-radius: 4px;
        color: #d8ebff;
        flex: 0 0 auto;
        font-size: 13px;
        font-weight: 700;
        padding: 6px 8px;
      }
      .bpmn-dialog__name {
        font-size: 15px;
        font-weight: 600;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .bpmn-dialog__status {
        color: #a6a6a6;
        font-size: 12px;
      }
      .bpmn-dialog__button,
      .bpmn-dialog__close {
        align-items: center;
        background: #242424;
        border: 1px solid rgba(255,255,255,.2);
        color: #ffffff;
        cursor: pointer;
        display: inline-flex;
        font: inherit;
        justify-content: center;
      }
      .bpmn-dialog__button {
        border-radius: 4px;
        height: 36px;
        min-width: 36px;
        padding: 0;
        width: 36px;
      }
      .bpmn-dialog__button svg {
        display: block;
        height: 18px;
        width: 18px;
      }
      .bpmn-dialog__button:hover:not(:disabled),
      .bpmn-dialog__close:hover {
        background: rgba(255,255,255,.1);
      }
      .bpmn-dialog__button:disabled {
        color: #777777;
        cursor: not-allowed;
      }
      .bpmn-dialog__close {
        border-radius: 4px;
        font-size: 24px;
        height: 36px;
        line-height: 1;
        padding: 0 0 3px;
        width: 36px;
      }
      .bpmn-dialog__message {
        flex: 0 0 auto;
        border-bottom: 1px solid rgba(255,255,255,.12);
        color: #f8d66d;
        font-size: 13px;
        padding: 8px 20px;
      }
      .bpmn-dialog__message--error {
        background: #fde7e9;
        color: #7b1f23;
      }
      .bpmn-dialog__canvas {
        background: #ffffff;
        flex: 1 1 auto;
        height: auto;
        min-height: 0;
        overflow: hidden;
        width: 100%;
      }
      .bpmn-dialog__canvas .djs-container,
      .bpmn-dialog__canvas .djs-container > svg {
        background: #ffffff;
        height: 100% !important;
        width: 100% !important;
      }
      :fullscreen .bpmn-dialog {
        min-height: 100dvh;
      }
      @media (max-width: 720px) {
        .bpmn-dialog__header {
          align-items: stretch;
          flex-direction: column;
          flex-basis: auto;
        }
        .bpmn-dialog__actions {
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
  }

  private destroyRenderer(): void {
    if (this.bpmnInstance) {
      this.bpmnInstance.destroy();
      this.bpmnInstance = undefined;
    }
  }

  private isEditable(): boolean {
    return this.extensionSettings.mode === 'modeler';
  }

  private setBusy(isBusy: boolean, status: string): void {
    this.setStatus(status);
    this.domElement.querySelectorAll('.bpmn-dialog__button').forEach((button) => {
      const typedButton = button as HTMLButtonElement;
      if (typedButton.dataset.action === 'save') {
        typedButton.disabled = isBusy || !this.isDirty || !this.isEditable();
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

  private setWarnings(warnings: Array<Error | { message?: string }> | undefined): void {
    if (!warnings?.length) {
      this.setMessage('');
      return;
    }

    this.setMessage(warnings.slice(0, 3).map((warning) => warning.message || 'BPMN import warning.').join(' '));
  }

  private setMessage(message: string): void {
    const messageElement = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.hidden = message.length === 0;
    messageElement.textContent = message;
    messageElement.classList.remove('bpmn-dialog__message--error');
  }

  private setError(message: string): void {
    const messageElement = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.hidden = false;
    messageElement.textContent = message;
    messageElement.classList.add('bpmn-dialog__message--error');
  }

  private renderMetadata(): void {
    const nameElement = this.domElement.querySelector('.bpmn-dialog__name') as HTMLElement | null;
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
    parts.push('SharePoint-hosted renderer');
    nameElement.title = `${this.fileName}\n${parts.join(' | ')}`;
  }

  private updateSaveButton(): void {
    const saveButton = this.domElement.querySelector('[data-action="save"]') as HTMLButtonElement | null;
    if (saveButton) {
      saveButton.disabled = !this.isDirty || !this.isEditable();
    }
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
