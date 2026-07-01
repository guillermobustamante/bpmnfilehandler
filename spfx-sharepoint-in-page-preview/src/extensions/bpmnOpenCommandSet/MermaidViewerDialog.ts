import { SPHttpClient } from '@microsoft/sp-http';
import type { IFileExtensionSettings } from './previewSettings';
import { SharePointFileService, type ISharePointFileMetadata } from './sharePointFileService';
import { renderIcon } from '../../shared/icons';
import { ensureDialogBaseStyles } from '../../shared/dialogUtils';

// mermaid is loaded dynamically to keep it out of the main SPFx bundle.
// License: MIT — https://github.com/mermaid-js/mermaid
type MermaidModule = typeof import('mermaid');

export class MermaidViewerDialog {
  private el: HTMLDialogElement | undefined;
  private fileService: SharePointFileService;
  private metadata: ISharePointFileMetadata | undefined;
  private mermaidModule: MermaidModule | undefined;
  private svgContent: string = '';
  private rawText: string = '';
  private isDirty: boolean = false;
  private renderTimer: number = 0;
  private readonly onFullscreenChangeBound = (): void => this.updateFullscreenButton();

  // Pan / zoom state
  private zoom: number = 1;
  private panX: number = 0;
  private panY: number = 0;
  private isDragging: boolean = false;
  private dragStartX: number = 0;
  private dragStartY: number = 0;
  private dragStartPanX: number = 0;
  private dragStartPanY: number = 0;

  public constructor(
    private readonly hostEl: HTMLElement,
    spHttpClient: SPHttpClient,
    webAbsoluteUrl: string,
    private readonly serverRelativeUrl: string,
    private readonly fileName: string,
    private readonly extensionSettings: IFileExtensionSettings
  ) {
    this.fileService = new SharePointFileService(spHttpClient, webAbsoluteUrl);
  }

  public open(): void {
    ensureDialogBaseStyles(this.hostEl);
    const dlg = document.createElement('dialog');
    dlg.className = 'bpf-viewer-dialog';
    this.hostEl.appendChild(dlg);
    this.el = dlg;
    this.render();
    dlg.showModal();
    dlg.addEventListener('close', () => { this.afterClose(); }, { once: true });
  }

  private closeDialog(): void {
    this.el?.close();
  }

  private render(): void {
    const badge = this.extensionSettings.extension.replace('.', '').toUpperCase();
    const editable = this.isEditable();

    const saveButton = editable
      ? `<button class="mmd-dialog__button" data-action="save" type="button" aria-label="Save" title="Save" disabled>${renderIcon('save')}</button>`
      : '';

    const bodyHtml = editable
      ? `<div class="mmd-dialog__split">
          <div class="mmd-dialog__editor-pane">
            <div class="mmd-dialog__editor-label">Mermaid source</div>
            <textarea class="mmd-dialog__editor-area" data-role="editor" spellcheck="false" autocomplete="off" autocorrect="off" autocapitalize="off"></textarea>
          </div>
          <div class="mmd-dialog__canvas" data-role="canvas">
            <div class="mmd-dialog__diagram" data-role="diagram"></div>
          </div>
        </div>`
      : `<div class="mmd-dialog__canvas" data-role="canvas">
          <div class="mmd-dialog__diagram" data-role="diagram"></div>
        </div>`;

    this.el!.innerHTML = `
      <div class="mmd-dialog">
        <div class="mmd-dialog__header">
          <div class="mmd-dialog__title">
            <span class="mmd-dialog__badge">${escapeHtml(badge)}</span>
            <span class="mmd-dialog__name" title="${escapeHtml(this.fileName)}">${escapeHtml(this.fileName)}</span>
            <span class="mmd-dialog__status" data-role="status">Loading</span>
          </div>
          <div class="mmd-dialog__actions">
            ${saveButton}
            <button class="mmd-dialog__button" data-action="reload" type="button" aria-label="Reload" title="Reload">${renderIcon('refresh')}</button>
            <button class="mmd-dialog__button" data-action="zoom-out" type="button" aria-label="Zoom out" title="Zoom out">${renderIcon('zoomOut')}</button>
            <button class="mmd-dialog__button" data-action="zoom-in" type="button" aria-label="Zoom in" title="Zoom in">${renderIcon('zoomIn')}</button>
            <button class="mmd-dialog__button" data-action="fit" type="button" aria-label="Fit to screen" title="Fit to screen">${renderIcon('maximize')}</button>
            <button class="mmd-dialog__button" data-action="download" type="button" aria-label="Download SVG" title="Download SVG">${renderIcon('download')}</button>
            <button class="mmd-dialog__button" data-action="fullscreen" type="button" aria-label="Open full screen" title="Open full screen">${renderIcon('external')}</button>
            <button class="mmd-dialog__close" type="button" aria-label="Close preview" title="Close">&times;</button>
          </div>
        </div>
        <div class="mmd-dialog__message" data-role="message" hidden></div>
        ${bodyHtml}
      </div>
    `;

    this.ensureStyles();
    this.wireEvents();
    this.load().catch((error: unknown) => {
      this.setError(error instanceof Error ? error.message : 'Could not open Mermaid file.');
    });
  }

  private afterClose(): void {
    document.removeEventListener('fullscreenchange', this.onFullscreenChangeBound);
    window.clearTimeout(this.renderTimer);
    this.exitFullscreen().catch(() => undefined);
    this.el?.remove();
    this.el = undefined;
  }

  private isEditable(): boolean {
    return this.extensionSettings.mode === 'modeler';
  }

  private async load(): Promise<void> {
    if (this.isDirty) {
      const confirmed = window.confirm('You have unsaved changes. Reload and discard them?');
      if (!confirmed) return;
    }

    this.setBusy(true, 'Loading');
    this.setMessage('');

    const [metadata, text] = await Promise.all([
      this.fileService.getMetadata(this.serverRelativeUrl),
      this.fileService.getContent(this.serverRelativeUrl)
    ]);
    this.metadata = metadata;
    this.rawText = text;
    this.isDirty = false;
    this.renderMetadata();

    const editorEl = this.el!.querySelector('[data-role="editor"]') as HTMLTextAreaElement | null;
    if (editorEl) {
      editorEl.value = text;
    }

    this.setBusy(true, 'Rendering');
    await this.renderDiagram(text);
    this.setBusy(false, this.isEditable() ? 'Ready' : 'Preview');
    this.updateSaveButton();
  }

  private async renderDiagram(text: string): Promise<void> {
    const diagramEl = this.el!.querySelector('[data-role="diagram"]') as HTMLElement | null;
    if (!diagramEl) {
      return;
    }

    // Reject files that are far too large or obviously not Mermaid diagrams.
    const trimmed = text.trim();
    if (trimmed.length > 200_000) {
      throw new Error(
        `This file is ${Math.round(trimmed.length / 1024)} KB, which exceeds the Mermaid rendering limit. ` +
        'Mermaid diagrams should be under 200 KB of text.'
      );
    }
    // Detect STEP/ISO-10303 files which are never Mermaid diagrams.
    if (/^ISO-10303-21/i.test(trimmed)) {
      throw new Error(
        'This file is a STEP/STP CAD file, not a Mermaid diagram. ' +
        'To preview STEP/STP files, enable the STEP CAD model extension in File Preview Admin settings and set its renderer to "occt-step".'
      );
    }

    if (!this.mermaidModule) {
      this.setBusy(true, 'Loading library');
      this.mermaidModule = (await import(/* webpackChunkName: 'mermaid' */ 'mermaid')) as MermaidModule;
      this.mermaidModule.default.initialize({
        startOnLoad: false,
        theme: 'default',
        securityLevel: 'strict',
        suppressErrorRendering: false,
        maxTextSize: 100_000
      });
    }

    const id = `mmd-${Date.now()}`;
    let svg: string;
    try {
      // Race the render against a 15-second timeout to prevent hanging on malformed large inputs.
      const renderPromise = this.mermaidModule.default.render(id, trimmed);
      const timeoutPromise = new Promise<never>((_resolve, reject) => {
        window.setTimeout(
          () => reject(new Error('Mermaid render timed out. The file may not be a valid Mermaid diagram.')),
          15_000
        );
      });
      const renderResult = await Promise.race([renderPromise, timeoutPromise]);
      svg = renderResult.svg;
    } catch (renderErr: unknown) {
      console.error('[MermaidViewerDialog] render failed:', renderErr);
      const msg = renderErr instanceof Error ? renderErr.message : String(renderErr);
      throw new Error(
        `Mermaid could not render this diagram: ${msg.slice(0, 300)}. ` +
        'Verify the file contains valid Mermaid syntax.'
      );
    }
    this.svgContent = svg;
    diagramEl.innerHTML = svg;

    // Apply initial transform so the SVG fills the canvas
    this.resetZoom();
  }

  private async save(): Promise<void> {
    const editorEl = this.el!.querySelector('[data-role="editor"]') as HTMLTextAreaElement | null;
    if (!editorEl || !this.isEditable()) {
      return;
    }

    const text = editorEl.value;
    this.setBusy(true, 'Saving');

    try {
      this.metadata = await this.fileService.saveContent(this.serverRelativeUrl, text, this.metadata?.eTag);
      this.rawText = text;
      this.isDirty = false;
      this.renderMetadata();
      this.setBusy(false, 'Saved');
      this.updateSaveButton();
      window.setTimeout(() => {
        if (!this.isDirty) {
          this.setStatus('Ready');
        }
      }, 2000);
    } catch (error) {
      this.setBusy(false, 'Error');
      this.setError(error instanceof Error ? error.message : 'Could not save the file.');
    }
  }

  private scheduleRerender(): void {
    window.clearTimeout(this.renderTimer);
    this.renderTimer = window.setTimeout(() => {
      const editorEl = this.el!.querySelector('[data-role="editor"]') as HTMLTextAreaElement | null;
      if (!editorEl) return;
      const text = editorEl.value;
      this.setStatus('Rendering…');
      this.renderDiagram(text)
        .then(() => {
          if (!this.isDirty) {
            this.setStatus('Ready');
          } else {
            this.setStatus('Unsaved');
          }
          this.setMessage('');
        })
        .catch((err: unknown) => {
          this.setError(err instanceof Error ? err.message : 'Render error');
        });
    }, 800);
  }

  private updateSaveButton(): void {
    const saveBtn = this.el!.querySelector('[data-action="save"]') as HTMLButtonElement | null;
    if (saveBtn) {
      saveBtn.disabled = !this.isDirty;
    }
  }

  private resetZoom(): void {
    this.zoom = 1;
    this.panX = 0;
    this.panY = 0;
    this.applyTransform();
  }

  private applyTransform(): void {
    const diagramEl = this.el!.querySelector('[data-role="diagram"]') as HTMLElement | null;
    if (diagramEl) {
      diagramEl.style.transform = `translate(${this.panX}px, ${this.panY}px) scale(${this.zoom})`;
    }
  }

  private changeZoom(factor: number, originX?: number, originY?: number): void {
    const canvasEl = this.el!.querySelector('[data-role="canvas"]') as HTMLElement | null;
    const cx = originX ?? (canvasEl ? canvasEl.clientWidth / 2 : 0);
    const cy = originY ?? (canvasEl ? canvasEl.clientHeight / 2 : 0);

    const prevZoom = this.zoom;
    this.zoom = Math.max(0.1, Math.min(10, this.zoom * factor));
    const scale = this.zoom / prevZoom;

    this.panX = cx + (this.panX - cx) * scale;
    this.panY = cy + (this.panY - cy) * scale;
    this.applyTransform();
  }

  private wireEvents(): void {
    this.el!.querySelector('[data-action="reload"]')?.addEventListener('click', () => {
      window.clearTimeout(this.renderTimer);
      this.load().catch((e: unknown) => this.setError(e instanceof Error ? e.message : 'Could not reload.'));
    });
    this.el!.querySelector('[data-action="zoom-out"]')?.addEventListener('click', () => this.changeZoom(0.8));
    this.el!.querySelector('[data-action="zoom-in"]')?.addEventListener('click', () => this.changeZoom(1.25));
    this.el!.querySelector('[data-action="fit"]')?.addEventListener('click', () => this.resetZoom());
    this.el!.querySelector('[data-action="download"]')?.addEventListener('click', () => this.download());
    this.el!.querySelector('[data-action="fullscreen"]')?.addEventListener('click', () => {
      this.toggleFullscreen().catch((e: unknown) => this.setError(e instanceof Error ? e.message : 'Could not open full screen.'));
    });
    document.addEventListener('fullscreenchange', this.onFullscreenChangeBound);
    this.el!.querySelector('.mmd-dialog__close')?.addEventListener('click', () => {
      this.closeDialog();
    });

    if (this.isEditable()) {
      this.el!.querySelector('[data-action="save"]')?.addEventListener('click', () => {
        this.save().catch((e: unknown) => this.setError(e instanceof Error ? e.message : 'Could not save.'));
      });

      const editorEl = this.el!.querySelector('[data-role="editor"]') as HTMLTextAreaElement | null;
      if (editorEl) {
        editorEl.addEventListener('input', () => {
          this.isDirty = true;
          this.updateSaveButton();
          this.scheduleRerender();
        });
        // Tab key inserts two spaces instead of moving focus
        editorEl.addEventListener('keydown', (e: KeyboardEvent) => {
          if (e.key === 'Tab') {
            e.preventDefault();
            const start = editorEl.selectionStart;
            const end = editorEl.selectionEnd;
            editorEl.value = editorEl.value.slice(0, start) + '  ' + editorEl.value.slice(end);
            editorEl.selectionStart = editorEl.selectionEnd = start + 2;
          }
          if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            this.save().catch(() => undefined);
          }
        });
      }
    }

    // Pan/zoom on the canvas
    const canvasEl = this.el!.querySelector('[data-role="canvas"]') as HTMLElement | null;
    if (canvasEl) {
      canvasEl.addEventListener('wheel', (e: WheelEvent) => {
        e.preventDefault();
        const rect = canvasEl.getBoundingClientRect();
        const ox = e.clientX - rect.left;
        const oy = e.clientY - rect.top;
        this.changeZoom(e.deltaY < 0 ? 1.15 : 0.87, ox, oy);
      }, { passive: false });

      canvasEl.addEventListener('mousedown', (e: MouseEvent) => {
        if (e.button !== 0) {
          return;
        }
        this.isDragging = true;
        this.dragStartX = e.clientX;
        this.dragStartY = e.clientY;
        this.dragStartPanX = this.panX;
        this.dragStartPanY = this.panY;
        canvasEl.style.cursor = 'grabbing';
      });
      canvasEl.addEventListener('mousemove', (e: MouseEvent) => {
        if (!this.isDragging) {
          return;
        }
        this.panX = this.dragStartPanX + (e.clientX - this.dragStartX);
        this.panY = this.dragStartPanY + (e.clientY - this.dragStartY);
        this.applyTransform();
      });
      const stopDrag = (): void => {
        this.isDragging = false;
        canvasEl.style.cursor = 'grab';
      };
      canvasEl.addEventListener('mouseup', stopDrag);
      canvasEl.addEventListener('mouseleave', stopDrag);
    }
  }

  private download(): void {
    if (!this.svgContent) {
      return;
    }

    const blob = new Blob([this.svgContent], { type: 'image/svg+xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = this.fileName.replace(/\.(mmd|mermaid)$/i, '') + '.svg';
    a.click();
    URL.revokeObjectURL(url);
  }

  private async toggleFullscreen(): Promise<void> {
    if (document.fullscreenElement) {
      await this.exitFullscreen();
    } else {
      const target = this.el!.firstElementChild as HTMLElement | null;
      await (target ?? document.documentElement).requestFullscreen();
      this.updateFullscreenButton();
    }
  }

  private async exitFullscreen(): Promise<void> {
    if (document.fullscreenElement) {
      await document.exitFullscreen();
      this.updateFullscreenButton();
    }
  }

  private updateFullscreenButton(): void {
    const btn = this.el!.querySelector('[data-action="fullscreen"]') as HTMLButtonElement | null;
    if (!btn) {
      return;
    }

    const isFs = Boolean(document.fullscreenElement);
    btn.innerHTML = renderIcon(isFs ? 'restore' : 'external');
    btn.setAttribute('aria-label', isFs ? 'Exit full screen' : 'Open full screen');
    btn.title = isFs ? 'Exit full screen' : 'Open full screen';
  }

  private ensureStyles(): void {
    if (this.el!.querySelector('style[data-bpf-preview-style="mmd"]')) {
      return;
    }

    const style = document.createElement('style');
    style.dataset.bpfPreviewStyle = 'mmd';
    style.textContent = `
      .mmd-dialog {
        background: #f5f5f5;
        color: #f5f5f5;
        display: flex;
        flex-direction: column;
        font-family: "Segoe UI", Arial, sans-serif;
        height: 100%;
        min-height: 0;
        overflow: hidden;
        width: 100%;
      }
      .mmd-dialog__header {
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
      .mmd-dialog__title,
      .mmd-dialog__actions {
        align-items: center;
        display: flex;
        gap: 8px;
        min-width: 0;
      }
      .mmd-dialog__title { flex: 1 1 auto; }
      .mmd-dialog__actions { flex: 0 0 auto; flex-wrap: nowrap; justify-content: flex-end; }
      .mmd-dialog__badge {
        background: #242424;
        border: 1px solid rgba(255,255,255,.18);
        border-radius: 4px;
        color: #d8ebff;
        flex: 0 0 auto;
        font-size: 13px;
        font-weight: 700;
        padding: 6px 8px;
      }
      .mmd-dialog__name {
        font-size: 15px;
        font-weight: 600;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .mmd-dialog__status { color: #a6a6a6; font-size: 12px; }
      .mmd-dialog__button,
      .mmd-dialog__close {
        align-items: center;
        background: #242424;
        border: 1px solid rgba(255,255,255,.2);
        color: #ffffff;
        cursor: pointer;
        display: inline-flex;
        font: inherit;
        justify-content: center;
      }
      .mmd-dialog__button {
        border-radius: 4px;
        height: 36px;
        min-width: 36px;
        padding: 0;
        width: 36px;
      }
      .mmd-dialog__button svg { display: block; height: 18px; width: 18px; }
      .mmd-dialog__button:hover:not(:disabled),
      .mmd-dialog__close:hover { background: rgba(255,255,255,.1); }
      .mmd-dialog__button:disabled { color: #777; cursor: not-allowed; }
      .mmd-dialog__close {
        border-radius: 4px;
        font-size: 24px;
        height: 36px;
        line-height: 1;
        padding: 0 0 3px;
        width: 36px;
      }
      .mmd-dialog__message {
        flex: 0 0 auto;
        border-bottom: 1px solid rgba(0,0,0,.1);
        color: #7b1f23;
        font-size: 13px;
        max-height: 8em;
        overflow-y: auto;
        padding: 8px 20px;
      }
      .mmd-dialog__message--error { background: #fde7e9; }
      .mmd-dialog__canvas {
        background: #ffffff;
        cursor: grab;
        flex: 1 1 auto;
        min-height: 0;
        overflow: hidden;
        position: relative;
        width: 100%;
      }
      .mmd-dialog__diagram {
        display: inline-block;
        padding: 24px;
        transform-origin: 0 0;
        user-select: none;
      }
      .mmd-dialog__diagram svg {
        display: block;
        max-width: none;
      }
      /* Modeler split layout */
      .mmd-dialog__split {
        display: flex;
        flex: 1 1 auto;
        min-height: 0;
        overflow: hidden;
      }
      .mmd-dialog__editor-pane {
        background: #1e1e1e;
        border-right: 1px solid rgba(255,255,255,.08);
        display: flex;
        flex: 0 0 40%;
        flex-direction: column;
        min-height: 0;
        min-width: 200px;
      }
      .mmd-dialog__split .mmd-dialog__canvas {
        flex: 1 1 auto;
        min-width: 0;
      }
      .mmd-dialog__editor-label {
        background: #252526;
        border-bottom: 1px solid rgba(255,255,255,.06);
        color: #858585;
        flex: 0 0 auto;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: .04em;
        padding: 6px 12px;
        text-transform: uppercase;
      }
      .mmd-dialog__editor-area {
        background: #1e1e1e;
        border: none;
        box-sizing: border-box;
        caret-color: #d4d4d4;
        color: #d4d4d4;
        flex: 1 1 auto;
        font-family: "Cascadia Code", "Fira Code", Consolas, "Courier New", monospace;
        font-size: 13px;
        line-height: 1.6;
        min-height: 0;
        outline: none;
        overflow: auto;
        padding: 12px 16px;
        resize: none;
        tab-size: 2;
        width: 100%;
      }
      .mmd-dialog__editor-area::selection { background: #264f78; }
      .mmd-dialog__editor-area:focus { outline: none; }
      .mmd-dialog:fullscreen { min-height: 100dvh; }
      @media (max-width: 720px) {
        .mmd-dialog__header { align-items: stretch; flex-direction: column; flex-basis: auto; }
        .mmd-dialog__actions { justify-content: flex-start; }
        .mmd-dialog__split { flex-direction: column; }
        .mmd-dialog__editor-pane { flex: 0 0 40%; border-right: none; border-bottom: 1px solid rgba(255,255,255,.08); }
      }
    `;
    this.el!.appendChild(style);
  }

  private setBusy(isBusy: boolean, status: string): void {
    this.setStatus(status);
    this.el!.querySelectorAll('.mmd-dialog__button').forEach((btn) => {
      const typedBtn = btn as HTMLButtonElement;
      if (typedBtn.dataset.action === 'save') {
        // Save button follows isDirty state, not busy state
        typedBtn.disabled = isBusy || !this.isDirty;
      } else {
        typedBtn.disabled = isBusy;
      }
    });
  }

  private setStatus(status: string): void {
    const el = this.el!.querySelector('[data-role="status"]') as HTMLElement | null;
    if (el) {
      el.textContent = status;
    }
  }

  private setMessage(message: string): void {
    const el = this.el!.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!el) {
      return;
    }

    el.hidden = message.length === 0;
    el.textContent = message;
    el.classList.remove('mmd-dialog__message--error');
  }

  private setError(message: string): void {
    const el = this.el!.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!el) {
      return;
    }

    el.hidden = false;
    el.textContent = message;
    el.classList.add('mmd-dialog__message--error');
    this.setBusy(false, 'Error');
  }

  private renderMetadata(): void {
    const nameEl = this.el!.querySelector('.mmd-dialog__name') as HTMLElement | null;
    if (!nameEl || !this.metadata) {
      return;
    }

    const parts: string[] = [];
    if (this.metadata.timeLastModified) {
      parts.push(`Modified ${new Date(this.metadata.timeLastModified).toLocaleString()}`);
    }
    if (this.metadata.length) {
      parts.push(formatBytes(Number(this.metadata.length)));
    }
    const modeLabel = this.isEditable() ? 'Mermaid.js modeler (MIT)' : 'Mermaid.js renderer (MIT)';
    parts.push(modeLabel);
    nameEl.title = `${this.fileName}\n${parts.join(' | ')}`;
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
