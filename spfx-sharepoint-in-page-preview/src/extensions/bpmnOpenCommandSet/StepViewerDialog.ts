import { SPHttpClient } from '@microsoft/sp-http';
import type { IFileExtensionSettings } from './previewSettings';
import { SharePointFileService, type ISharePointFileMetadata } from './sharePointFileService';
import { renderIcon } from '../../shared/icons';
import { ensureDialogBaseStyles } from '../../shared/dialogUtils';

// occt-import-js and three are loaded via dynamic import so they are split into
// separate webpack chunks and only fetched when the dialog actually opens.
// web-ifc license: MIT  https://github.com/IFCjs/web-ifc
// occt-import-js license: LGPL-2.1  https://github.com/kovacsv/occt-import-js
// occt-import-js WASM must remain a separate file per LGPL-2.1 relinkability requirement.
// three license: MIT  https://github.com/mrdoob/three.js

/* eslint-disable @typescript-eslint/no-explicit-any */
type OcctModule = { default: (opts: { locateFile: (p: string) => string }) => Promise<any> };
/* eslint-enable @typescript-eslint/no-explicit-any */
type ThreeModule = typeof import('three');
type OrbitControlsModule = typeof import('three/examples/jsm/controls/OrbitControls');

declare const __webpack_public_path__: string;
// Production: WASM is co-located with JS chunks in the sppkg (SharePoint CDN).
// Dev (localhost): fall back to jsDelivr so the dev server works without extra setup.
const OCCT_WASM_CDN: string = (typeof __webpack_public_path__ !== 'undefined' && window.location.hostname !== 'localhost')
  ? __webpack_public_path__
  : 'https://cdn.jsdelivr.net/npm/occt-import-js@0.0.23/dist/';

// ─── Tree icons (16×16 viewBox, stroke-based) ────────────────────────────────
const S = (d: string): string =>
  `<svg aria-hidden="true" focusable="false" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`;

const SVG_ASSEMBLY = S('<path d="M8 2L14 5v6l-6 3-6-3V5z"/><path d="M8 2v9"/><path d="M2 5l6 3 6-3"/>');
const SVG_PART     = S('<rect x="3" y="6" width="10" height="7" rx="1"/><path d="M3 6l5-2.5L13 6"/>');
const SVG_EYE      = S('<path d="M1 8s3-4.5 7-4.5S15 8 15 8s-3 4.5-7 4.5S1 8 1 8z"/><circle cx="8" cy="8" r="2"/>');
const SVG_EYE_OFF  = S('<path d="M1.5 1.5l13 13"/><path d="M4.3 4.3C2.9 5.4 1.8 6.7 1 8c1.5 2.5 4 4.5 7 4.5 1.1 0 2.1-.2 3-.6"/><path d="M9.8 5.1C12.5 5.8 14.4 7.6 15 8c-.5.9-1.4 1.9-2.5 2.7"/>');
const SVG_CHR      = S('<path d="M6 4l4 4-4 4"/>');
const SVG_CHD      = S('<path d="M4 6l4 4 4-4"/>');
const SVG_CHL      = S('<path d="M10 4L6 8l4 4"/>');

// ─── OCCT data shape ──────────────────────────────────────────────────────────
interface OcctNode {
  name: string;
  meshes: number[];
  children: OcctNode[];
}

// ─── Internal assembly tree ───────────────────────────────────────────────────
interface StepTreeNode {
  id: string;
  name: string;
  meshIndices: number[];
  allMeshIndices: number[];
  children: StepTreeNode[];
  isAssembly: boolean;
  expanded: boolean;
  visible: boolean;
}

export class StepViewerDialog {
  private el: HTMLDialogElement | undefined;
  private fileService: SharePointFileService;
  private metadata: ISharePointFileMetadata | undefined;
  private cancelled: boolean = false;
  private animFrameId: number = 0;
  private renderer: import('three').WebGLRenderer | undefined;
  private scene: import('three').Scene | undefined;
  private camera: import('three').PerspectiveCamera | undefined;
  private orbitControls: import('three/examples/jsm/controls/OrbitControls').OrbitControls | undefined;
  private resizeObserver: ResizeObserver | undefined;
  private threeModule: ThreeModule | undefined;

  private rootNodes: StepTreeNode[] = [];
  private nodeMap: Map<string, StepTreeNode> = new Map();
  private meshObjects: (import('three').Mesh | undefined)[] = [];
  private meshColors: [number, number, number][] = [];
  private selectedNodeId: string | undefined;
  private highlightedMeshes: Map<import('three').Mesh, import('three').Color> = new Map();
  private searchQuery: string = '';
  private raycaster: import('three').Raycaster | undefined;
  private mouseDownPos: { x: number; y: number } | undefined;
  private readonly onFullscreenChangeBound = (): void => this.updateFullscreenButton();

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

    this.el!.innerHTML = `
      <div class="step-dialog">
        <div class="step-dialog__header">
          <div class="step-dialog__title">
            <span class="step-dialog__badge">${escapeHtml(badge)}</span>
            <span class="step-dialog__name" title="${escapeHtml(this.fileName)}">${escapeHtml(this.fileName)}</span>
            <span class="step-dialog__status" data-role="status">Loading</span>
          </div>
          <div class="step-dialog__actions">
            <button class="step-dialog__button" data-action="reload" type="button" aria-label="Reload" title="Reload">${renderIcon('refresh')}</button>
            <button class="step-dialog__button" data-action="fit" type="button" aria-label="Fit to screen" title="Fit to screen">${renderIcon('maximize')}</button>
            <button class="step-dialog__button" data-action="fullscreen" type="button" aria-label="Open full screen" title="Open full screen">${renderIcon('external')}</button>
            <button class="step-dialog__close" type="button" aria-label="Close preview" title="Close">&times;</button>
          </div>
        </div>
        <div class="step-dialog__message" data-role="message" hidden></div>
        <div class="step-dialog__body">
          <aside class="step-tree" data-role="tree-panel" aria-label="Model tree" hidden>
            <div class="step-tree__header">
              <span class="step-tree__title">Model Tree</span>
              <div class="step-tree__hdr-actions">
                <button class="step-tree__hdr-btn" data-action="show-all" title="Show all">${SVG_EYE}</button>
                <button class="step-tree__hdr-btn" data-action="hide-all" title="Hide all">${SVG_EYE_OFF}</button>
                <button class="step-tree__hdr-btn" data-action="collapse-panel" title="Collapse panel" aria-label="Collapse panel">${SVG_CHL}</button>
              </div>
            </div>
            <div class="step-tree__search" hidden>
              <input class="step-tree__search-input" type="text" placeholder="Filter parts…"
                data-role="tree-search" autocomplete="off" spellcheck="false" />
            </div>
            <div class="step-tree__scroll" data-role="tree-nodes" role="tree"></div>
            <div class="step-tree__footer" data-role="tree-footer"></div>
          </aside>
          <button class="step-tree__expand-btn" data-action="expand-panel" hidden
            aria-label="Expand model tree" title="Show model tree">${SVG_CHR}</button>
          <div class="step-dialog__canvas" data-role="canvas"></div>
        </div>
      </div>
    `;

    this.ensureStyles();
    this.wireEvents();
    this.load().catch((error: unknown) => {
      if (!this.cancelled) {
        this.setError(error instanceof Error ? error.message : 'Could not open STEP file.');
      }
    });
  }

  private afterClose(): void {
    document.removeEventListener('fullscreenchange', this.onFullscreenChangeBound);
    this.cancelled = true;
    this.teardownScene();
    this.exitFullscreen().catch(() => undefined);
    this.el?.remove();
    this.el = undefined;
  }

  // ── Core loading pipeline ─────────────────────────────────────────────────

  private async load(): Promise<void> {
    this.setBusy(true, 'Downloading file…');
    this.setMessage('');

    const [metadata, arrayBuffer] = await Promise.all([
      this.fileService.getMetadata(this.serverRelativeUrl),
      this.fileService.getContentAsArrayBuffer(this.serverRelativeUrl)
    ]);
    if (this.cancelled) return;

    this.metadata = metadata;
    this.renderMetadata();

    if (arrayBuffer.byteLength > 200 * 1024 * 1024) {
      throw new Error(
        `This file is ${formatBytes(arrayBuffer.byteLength)}, which exceeds the 200 MB browser limit. ` +
        'Use a dedicated CAD application to open very large assemblies.'
      );
    }

    this.setBusy(true, 'Loading libraries…');
    const [occtMod, three, orbitMod] = await Promise.all([
      import(/* webpackChunkName: 'occt-import-js' */ 'occt-import-js') as Promise<OcctModule>,
      import(/* webpackChunkName: 'three' */ 'three') as Promise<ThreeModule>,
      import(/* webpackChunkName: 'three-orbit-controls' */ 'three/examples/jsm/controls/OrbitControls') as Promise<OrbitControlsModule>
    ]);
    if (this.cancelled) return;
    this.threeModule = three;

    this.setBusy(true, 'Initialising STEP engine…');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let occt: any;
    try {
      occt = await occtMod.default({ locateFile: (p: string) => OCCT_WASM_CDN + p });
    } catch (wasmErr: unknown) {
      console.error('[StepViewerDialog] WASM init failed:', wasmErr);
      throw new Error(
        'The STEP engine (occt-import-js WASM) failed to initialise. ' +
        'WebAssembly may be blocked by your browser security policy or corporate firewall.'
      );
    }
    if (this.cancelled) return;

    this.setBusy(true, 'Parsing file…');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let result: any;
    try {
      result = occt.ReadStepFile(new Uint8Array(arrayBuffer), null);
    } catch (parseErr: unknown) {
      console.error('[StepViewerDialog] ReadStepFile failed:', parseErr);
      throw new Error('The file could not be parsed. Ensure it is a valid AP203, AP214, or AP242 STEP file.');
    }

    if (!result?.success || !result.meshes?.length) {
      throw new Error(
        'No renderable geometry found in this STEP file. ' +
        'Supported schemas: AP203, AP214, AP242 basic solids. ' +
        'Verify the file contains 3D solid bodies in a CAD application.'
      );
    }
    if (this.cancelled) return;

    this.setBusy(true, 'Building geometry…');
    const { scene, camera, renderer } = this.initScene(three);
    this.scene = scene;
    this.camera = camera;
    this.renderer = renderer;

    this.buildTree(result);
    this.createMeshObjects(result, three, scene);

    this.fitCamera(three);
    this.startLoop(three, orbitMod);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const meshCount = (result.meshes as any[]).length;
    const bodyStr = `${meshCount.toLocaleString()} bod${meshCount === 1 ? 'y' : 'ies'}`;
    this.setBusy(false, bodyStr);

    if (this.rootNodes.length > 0) {
      const treePanel = this.el!.querySelector('[data-role="tree-panel"]') as HTMLElement | null;
      if (treePanel) {
        treePanel.hidden = false;
        const searchWrap = treePanel.querySelector('.step-tree__search') as HTMLElement | null;
        if (searchWrap && meshCount > 5) searchWrap.hidden = false;
      }
      this.renderTreePanel();
      const footerEl = this.el!.querySelector('[data-role="tree-footer"]') as HTMLElement | null;
      if (footerEl) {
        footerEl.textContent = `${meshCount.toLocaleString()} bod${meshCount === 1 ? 'y' : 'ies'} · click to select · dbl-click to zoom`;
      }
      this.wireTreeEvents();
      this.wireRaycaster(renderer.domElement, three);
    }
  }

  // ── Assembly tree building ────────────────────────────────────────────────

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private buildTree(result: any): void {
    this.rootNodes = [];
    this.nodeMap = new Map();
    const counter = { n: 0 };
    const raw: OcctNode = result.root;
    const topLevel: OcctNode[] = (raw && raw.children) ? raw.children : [];

    for (let i = 0; i < topLevel.length; i++) {
      this.rootNodes.push(this.buildTreeNode(topLevel[i], counter));
    }

    // Fallback: no hierarchy — flat list of body names
    if (this.rootNodes.length === 0) {
      const meshes = result.meshes as Array<{ name: string }>;
      for (let i = 0; i < meshes.length; i++) {
        const node: StepTreeNode = {
          id: 'n' + counter.n++,
          name: meshes[i].name ? capitalizeName(meshes[i].name) : `Body ${i + 1}`,
          meshIndices: [i],
          allMeshIndices: [i],
          children: [],
          isAssembly: false,
          expanded: false,
          visible: true
        };
        this.rootNodes.push(node);
        this.nodeMap.set(node.id, node);
      }
    }
  }

  private buildTreeNode(raw: OcctNode, counter: { n: number }): StepTreeNode {
    const id = 'n' + counter.n++;
    const children: StepTreeNode[] = [];
    const rawChildren = raw.children || [];
    for (let i = 0; i < rawChildren.length; i++) {
      children.push(this.buildTreeNode(rawChildren[i], counter));
    }

    const meshIndices: number[] = raw.meshes ? raw.meshes.slice() : [];
    const allMeshIndices: number[] = meshIndices.slice();
    for (let i = 0; i < children.length; i++) {
      const ci = children[i].allMeshIndices;
      for (let j = 0; j < ci.length; j++) {
        allMeshIndices.push(ci[j]);
      }
    }

    const node: StepTreeNode = {
      id,
      name: raw.name ? capitalizeName(raw.name) : 'Body',
      meshIndices,
      allMeshIndices,
      children,
      isAssembly: children.length > 0,
      expanded: true,
      visible: true
    };
    this.nodeMap.set(id, node);
    return node;
  }

  // ── Mesh creation ─────────────────────────────────────────────────────────

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private createMeshObjects(result: any, three: ThreeModule, scene: import('three').Scene): void {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const meshData = result.meshes as any[];
    this.meshObjects = new Array(meshData.length);
    this.meshColors = new Array(meshData.length);

    for (let i = 0; i < meshData.length; i++) {
      const md = meshData[i];
      const geo = new three.BufferGeometry();
      geo.setAttribute('position', new three.BufferAttribute(new Float32Array(md.attributes.position.array), 3));
      if (md.attributes.normal) {
        geo.setAttribute('normal', new three.BufferAttribute(new Float32Array(md.attributes.normal.array), 3));
      }
      if (md.index) {
        geo.setIndex(new three.BufferAttribute(new Uint32Array(md.index.array), 1));
      }
      if (!md.attributes.normal) {
        geo.computeVertexNormals();
      }

      const color: [number, number, number] = md.color ?? [0.6, 0.7, 0.8];
      this.meshColors[i] = color;

      const mat = new three.MeshLambertMaterial({
        color: new three.Color(color[0], color[1], color[2]),
        side: three.DoubleSide
      });
      const mesh = new three.Mesh(geo, mat);
      mesh.userData.meshIndex = i;
      this.meshObjects[i] = mesh;
      scene.add(mesh);
    }
  }

  // ── Tree panel rendering ──────────────────────────────────────────────────

  private renderTreePanel(): void {
    const container = this.el!.querySelector('[data-role="tree-nodes"]') as HTMLElement | null;
    if (!container) return;
    const html: string[] = [];
    for (let i = 0; i < this.rootNodes.length; i++) {
      this.renderTreeNode(this.rootNodes[i], 0, html);
    }
    container.innerHTML = html.join('');
  }

  private renderTreeNode(node: StepTreeNode, depth: number, html: string[]): void {
    if (this.searchQuery && !this.nodeMatchesSearch(node, this.searchQuery)) return;

    const isSelected = node.id === this.selectedNodeId;
    const isExpanded = node.expanded || this.searchQuery.length > 0;
    const hasChildren = node.children.length > 0;
    const depthPx = 8 + depth * 16;

    let swatchColor = '';
    if (node.allMeshIndices.length > 0) {
      const c = this.meshColors[node.allMeshIndices[0]];
      if (c) swatchColor = `rgb(${Math.round(c[0] * 255)},${Math.round(c[1] * 255)},${Math.round(c[2] * 255)})`;
    }

    const bodyCount = node.allMeshIndices.length;
    const countEl = bodyCount > 1 ? `<span class="stree-node__count">${bodyCount}</span>` : '';
    const swatchEl = swatchColor ? `<span class="stree-node__swatch" style="background:${escapeHtml(swatchColor)}"></span>` : '';
    const toggleEl = hasChildren
      ? `<button class="stree-node__toggle" data-action="expand" data-node-id="${node.id}" tabindex="-1" aria-label="${isExpanded ? 'Collapse' : 'Expand'}">${isExpanded ? SVG_CHD : SVG_CHR}</button>`
      : `<span class="stree-node__toggle stree-node__toggle--leaf"></span>`;
    const eyeClass = !node.visible ? ' stree-node__eye--off' : '';
    const eyeLabel = node.visible ? 'Hide' : 'Show';

    html.push(
      `<div class="stree-node${isSelected ? ' stree-node--selected' : ''}${!node.visible ? ' stree-node--faded' : ''}"` +
        ` data-node-id="${node.id}" data-role="tree-node" role="treeitem" aria-selected="${isSelected}" aria-expanded="${isExpanded}">` +
        `<div class="stree-node__row" style="padding-left:${depthPx}px" data-action="select" data-node-id="${node.id}">` +
          toggleEl +
          `<span class="stree-node__icon">${node.isAssembly ? SVG_ASSEMBLY : SVG_PART}</span>` +
          swatchEl +
          `<span class="stree-node__label" title="${escapeHtml(node.name)}">${escapeHtml(node.name)}</span>` +
          countEl +
          `<button class="stree-node__eye${eyeClass}" data-action="vis" data-node-id="${node.id}"` +
            ` aria-label="${eyeLabel}" title="${eyeLabel}" tabindex="-1">${node.visible ? SVG_EYE : SVG_EYE_OFF}</button>` +
        `</div>` +
      `</div>`
    );

    if (hasChildren && isExpanded) {
      for (let i = 0; i < node.children.length; i++) {
        this.renderTreeNode(node.children[i], depth + 1, html);
      }
    }
  }

  private nodeMatchesSearch(node: StepTreeNode, query: string): boolean {
    const q = query.toLowerCase();
    if (node.name.toLowerCase().indexOf(q) !== -1) return true;
    for (let i = 0; i < node.children.length; i++) {
      if (this.nodeMatchesSearch(node.children[i], q)) return true;
    }
    return false;
  }

  // ── Tree events ───────────────────────────────────────────────────────────

  private wireTreeEvents(): void {
    const panel = this.el!.querySelector('[data-role="tree-panel"]') as HTMLElement | null;
    if (!panel) return;

    panel.querySelector('[data-action="show-all"]')?.addEventListener('click', () => {
      this.setAllVisible(true);
      this.renderTreePanel();
    });
    panel.querySelector('[data-action="hide-all"]')?.addEventListener('click', () => {
      this.setAllVisible(false);
      this.renderTreePanel();
    });
    panel.querySelector('[data-action="collapse-panel"]')?.addEventListener('click', () => {
      panel.style.flexBasis = '0px';
      panel.style.minWidth = '0px';
      const expandBtn = this.el!.querySelector('[data-action="expand-panel"]') as HTMLElement | null;
      if (expandBtn) expandBtn.hidden = false;
    });

    const expandBtn = this.el!.querySelector('[data-action="expand-panel"]') as HTMLElement | null;
    expandBtn?.addEventListener('click', () => {
      panel.style.flexBasis = '';
      panel.style.minWidth = '';
      if (expandBtn) expandBtn.hidden = true;
    });

    const searchInput = panel.querySelector('[data-role="tree-search"]') as HTMLInputElement | null;
    searchInput?.addEventListener('input', () => {
      this.searchQuery = searchInput.value;
      this.renderTreePanel();
    });

    const scrollArea = panel.querySelector('[data-role="tree-nodes"]') as HTMLElement | null;
    if (!scrollArea) return;

    scrollArea.addEventListener('click', (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      const btn = target.closest('[data-action]') as HTMLElement | null;
      if (!btn) return;
      const action = btn.dataset.action;
      const nodeId = btn.dataset.nodeId || (btn.closest('[data-node-id]') as HTMLElement | null)?.dataset.nodeId;
      if (!nodeId) return;

      if (action === 'expand') {
        e.stopPropagation();
        const node = this.nodeMap.get(nodeId);
        if (node && node.children.length > 0) {
          node.expanded = !node.expanded;
          this.renderTreePanel();
        }
        return;
      }
      if (action === 'vis') {
        e.stopPropagation();
        this.toggleVisibility(nodeId);
        return;
      }
      if (action === 'select' && this.threeModule) {
        this.selectNode(nodeId, this.threeModule);
      }
    });

    scrollArea.addEventListener('dblclick', (e: MouseEvent) => {
      const nodeEl = (e.target as HTMLElement).closest('[data-node-id]') as HTMLElement | null;
      if (!nodeEl || !this.threeModule) return;
      const nodeId = nodeEl.dataset.nodeId;
      if (nodeId) this.fitToNode(nodeId, this.threeModule);
    });
  }

  // ── Raycaster (viewport click → tree selection) ───────────────────────────

  private wireRaycaster(canvas: HTMLCanvasElement, three: ThreeModule): void {
    this.raycaster = new three.Raycaster();
    canvas.addEventListener('mousedown', (e: MouseEvent) => {
      this.mouseDownPos = { x: e.clientX, y: e.clientY };
    });
    canvas.addEventListener('mouseup', (e: MouseEvent) => {
      if (!this.mouseDownPos) return;
      const dx = e.clientX - this.mouseDownPos.x;
      const dy = e.clientY - this.mouseDownPos.y;
      this.mouseDownPos = undefined;
      if (dx * dx + dy * dy > 25) return;
      this.handleCanvasClick(e, canvas, three);
    });
  }

  private handleCanvasClick(e: MouseEvent, canvas: HTMLCanvasElement, three: ThreeModule): void {
    if (!this.raycaster || !this.scene || !this.camera) return;
    const rect = canvas.getBoundingClientRect();
    const ndc = new three.Vector2(
      ((e.clientX - rect.left) / rect.width) * 2 - 1,
      -((e.clientY - rect.top) / rect.height) * 2 + 1
    );
    this.raycaster.setFromCamera(ndc, this.camera);
    const hits = this.raycaster.intersectObjects(this.scene.children, false);
    if (hits.length === 0) {
      this.clearSelection(three);
      return;
    }
    const mesh = hits[0].object as import('three').Mesh;
    const meshIdx: number | undefined =
      typeof mesh.userData.meshIndex === 'number' ? (mesh.userData.meshIndex as number) : undefined;
    if (meshIdx === undefined) {
      this.clearSelection(three);
      return;
    }
    const ownerId = this.findNodeByMeshIndex(meshIdx);
    if (ownerId) this.selectNode(ownerId, three);
    else this.clearSelection(three);
  }

  private findNodeByMeshIndex(idx: number): string | undefined {
    const search = (nodes: StepTreeNode[]): string | undefined => {
      for (let i = 0; i < nodes.length; i++) {
        const n = nodes[i];
        for (let j = 0; j < n.meshIndices.length; j++) {
          if (n.meshIndices[j] === idx) return n.id;
        }
        const found = search(n.children);
        if (found) return found;
      }
      return undefined;
    };
    return search(this.rootNodes);
  }

  // ── Selection ─────────────────────────────────────────────────────────────

  private selectNode(nodeId: string, three: ThreeModule): void {
    if (nodeId === this.selectedNodeId) {
      this.clearSelection(three);
      return;
    }
    this.clearSelection(three);
    this.selectedNodeId = nodeId;

    const node = this.nodeMap.get(nodeId);
    if (!node) return;

    for (let i = 0; i < node.allMeshIndices.length; i++) {
      const mesh = this.meshObjects[node.allMeshIndices[i]];
      if (!mesh) continue;
      const mat = mesh.material as import('three').MeshLambertMaterial;
      this.highlightedMeshes.set(mesh, mat.color.clone());
      mat.color.setHex(0x4488ff);
    }

    const treeNodes = this.el!.querySelectorAll('[data-role="tree-node"]');
    for (let i = 0; i < treeNodes.length; i++) {
      const el = treeNodes[i] as HTMLElement;
      const sel = el.dataset.nodeId === nodeId;
      el.classList.toggle('stree-node--selected', sel);
      el.setAttribute('aria-selected', String(sel));
      if (sel) el.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }

    this.expandAncestors(this.rootNodes, nodeId);
  }

  private clearSelection(three: ThreeModule): void {
    this.selectedNodeId = undefined;
    this.highlightedMeshes.forEach((color, mesh) => {
      (mesh.material as import('three').MeshLambertMaterial).color.copy(color);
    });
    this.highlightedMeshes.clear();
    const treeNodes = this.el!.querySelectorAll('[data-role="tree-node"]');
    for (let i = 0; i < treeNodes.length; i++) {
      const el = treeNodes[i] as HTMLElement;
      el.classList.remove('stree-node--selected');
      el.setAttribute('aria-selected', 'false');
    }
  }

  private expandAncestors(nodes: StepTreeNode[], targetId: string): boolean {
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i];
      if (n.id === targetId) return true;
      if (this.expandAncestors(n.children, targetId)) {
        n.expanded = true;
        return true;
      }
    }
    return false;
  }

  // ── Visibility ────────────────────────────────────────────────────────────

  private toggleVisibility(nodeId: string): void {
    const node = this.nodeMap.get(nodeId);
    if (!node) return;
    this.setVisibilityRecursive(node, !node.visible);
    this.renderTreePanel();
  }

  private setVisibilityRecursive(node: StepTreeNode, visible: boolean): void {
    node.visible = visible;
    for (let i = 0; i < node.meshIndices.length; i++) {
      const mesh = this.meshObjects[node.meshIndices[i]];
      if (mesh) mesh.visible = visible;
    }
    for (let i = 0; i < node.children.length; i++) {
      this.setVisibilityRecursive(node.children[i], visible);
    }
  }

  private setAllVisible(visible: boolean): void {
    for (let i = 0; i < this.rootNodes.length; i++) {
      this.setVisibilityRecursive(this.rootNodes[i], visible);
    }
  }

  // ── Camera ────────────────────────────────────────────────────────────────

  private fitToNode(nodeId: string, three: ThreeModule): void {
    const node = this.nodeMap.get(nodeId);
    if (!node) return;
    const box = new three.Box3();
    for (let i = 0; i < node.allMeshIndices.length; i++) {
      const mesh = this.meshObjects[node.allMeshIndices[i]];
      if (mesh) box.expandByObject(mesh);
    }
    if (!box.isEmpty()) this.fitCameraToBox(three, box);
  }

  private fitCamera(three: ThreeModule): void {
    if (!this.scene || !this.camera) return;
    const box = new three.Box3().setFromObject(this.scene);
    if (!box.isEmpty()) this.fitCameraToBox(three, box);
  }

  private fitCameraToBox(three: ThreeModule, box: import('three').Box3): void {
    if (!this.camera) return;
    const center = new three.Vector3();
    const size = new three.Vector3();
    box.getCenter(center);
    box.getSize(size);
    const maxDim = Math.max(size.x, size.y, size.z) || 1;
    const fovRad = (this.camera.fov * Math.PI) / 180;
    const distance = Math.abs(maxDim / 2 / Math.tan(fovRad / 2)) * 1.5;
    this.camera.position.set(center.x + distance * 0.6, center.y + distance * 0.5, center.z + distance);
    this.camera.lookAt(center);
    this.camera.near = distance / 1000;
    this.camera.far = distance * 20;
    this.camera.updateProjectionMatrix();
    if (this.orbitControls) {
      this.orbitControls.target.copy(center);
      this.orbitControls.update();
    }
  }

  // ── Scene setup ───────────────────────────────────────────────────────────

  private initScene(three: ThreeModule): {
    scene: import('three').Scene;
    camera: import('three').PerspectiveCamera;
    renderer: import('three').WebGLRenderer;
  } {
    const canvasEl = this.el!.querySelector('[data-role="canvas"]') as HTMLElement;
    const scene = new three.Scene();
    scene.background = new three.Color(0x1a1f2e);
    scene.add(new three.AmbientLight(0xffffff, 0.5));
    const d1 = new three.DirectionalLight(0xffffff, 0.8);
    d1.position.set(1, 2, 3);
    scene.add(d1);
    const d2 = new three.DirectionalLight(0xffffff, 0.3);
    d2.position.set(-2, -1, -2);
    scene.add(d2);

    const w = canvasEl.clientWidth || 800;
    const h = canvasEl.clientHeight || 600;
    const camera = new three.PerspectiveCamera(60, w / h, 0.001, 100000);

    const renderer = new three.WebGLRenderer({ antialias: true });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.setSize(w, h);
    canvasEl.appendChild(renderer.domElement);
    renderer.domElement.style.cssText = 'display:block;width:100%;height:100%;cursor:crosshair;';

    this.resizeObserver = new ResizeObserver(() => {
      const cw = canvasEl.clientWidth;
      const ch = canvasEl.clientHeight;
      if (cw > 0 && ch > 0 && this.camera && this.renderer) {
        this.camera.aspect = cw / ch;
        this.camera.updateProjectionMatrix();
        this.renderer.setSize(cw, ch);
      }
    });
    this.resizeObserver.observe(canvasEl);

    return { scene, camera, renderer };
  }

  private startLoop(three: ThreeModule, orbitMod: OrbitControlsModule): void {
    if (!this.renderer || !this.camera || !this.scene) return;
    this.orbitControls = new orbitMod.OrbitControls(this.camera, this.renderer.domElement);
    this.orbitControls.enableDamping = true;
    this.orbitControls.dampingFactor = 0.05;
    this.orbitControls.screenSpacePanning = false;
    this.fitCamera(three);
    const animate = (): void => {
      this.animFrameId = window.requestAnimationFrame(animate);
      this.orbitControls?.update();
      if (this.renderer && this.scene && this.camera) {
        this.renderer.render(this.scene, this.camera);
      }
    };
    animate();
  }

  private teardownScene(): void {
    if (this.animFrameId) { window.cancelAnimationFrame(this.animFrameId); this.animFrameId = 0; }
    this.resizeObserver?.disconnect();
    this.orbitControls?.dispose();
    if (this.scene) {
      this.scene.traverse((obj) => {
        const mesh = obj as import('three').Mesh;
        mesh.geometry?.dispose();
        if (Array.isArray(mesh.material)) {
          (mesh.material as import('three').Material[]).forEach((m) => m.dispose());
        } else if (mesh.material) {
          (mesh.material as import('three').Material).dispose();
        }
      });
    }
    this.renderer?.dispose();
    this.scene = undefined;
    this.camera = undefined;
    this.renderer = undefined;
    this.orbitControls = undefined;
    this.threeModule = undefined;
    this.meshObjects = [];
    this.meshColors = [];
    this.highlightedMeshes.clear();
    this.rootNodes = [];
    this.nodeMap.clear();
    this.selectedNodeId = undefined;
  }

  // ── Dialog wire-up ────────────────────────────────────────────────────────

  private wireEvents(): void {
    this.el!.querySelector('[data-action="reload"]')?.addEventListener('click', () => {
      if (this.cancelled) return;
      this.teardownScene();
      this.rootNodes = [];
      this.nodeMap.clear();
      this.searchQuery = '';
      const treePanel = this.el!.querySelector('[data-role="tree-panel"]') as HTMLElement | null;
      if (treePanel) { treePanel.hidden = true; treePanel.style.flexBasis = ''; treePanel.style.minWidth = ''; }
      const expandBtn = this.el!.querySelector('[data-action="expand-panel"]') as HTMLElement | null;
      if (expandBtn) expandBtn.hidden = true;
      this.cancelled = false;
      this.load().catch((e: unknown) =>
        this.setError(e instanceof Error ? e.message : 'Could not reload.')
      );
    });

    this.el!.querySelector('[data-action="fit"]')?.addEventListener('click', () => {
      import(/* webpackChunkName: 'three' */ 'three')
        .then((three: ThreeModule) => this.fitCamera(three))
        .catch(() => undefined);
    });

    this.el!.querySelector('[data-action="fullscreen"]')?.addEventListener('click', () => {
      this.toggleFullscreen().catch((e: unknown) =>
        this.setError(e instanceof Error ? e.message : 'Could not toggle full screen.')
      );
    });

    document.addEventListener('fullscreenchange', this.onFullscreenChangeBound);

    this.el!.querySelector('.step-dialog__close')?.addEventListener('click', () => {
      this.closeDialog();
    });
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
    if (document.fullscreenElement) { await document.exitFullscreen(); this.updateFullscreenButton(); }
  }

  private updateFullscreenButton(): void {
    const btn = this.el!.querySelector('[data-action="fullscreen"]') as HTMLButtonElement | null;
    if (!btn) return;
    const isFs = Boolean(document.fullscreenElement);
    btn.innerHTML = renderIcon(isFs ? 'restore' : 'external');
    btn.setAttribute('aria-label', isFs ? 'Exit full screen' : 'Open full screen');
    btn.title = isFs ? 'Exit full screen' : 'Open full screen';
  }

  // ── Status helpers ────────────────────────────────────────────────────────

  private setBusy(isBusy: boolean, status: string): void {
    this.setStatus(status);
    this.el!.querySelectorAll('.step-dialog__button').forEach((btn) => {
      (btn as HTMLButtonElement).disabled = isBusy;
    });
  }

  private setStatus(status: string): void {
    const el = this.el!.querySelector('[data-role="status"]') as HTMLElement | null;
    if (el) el.textContent = status;
  }

  private setMessage(message: string): void {
    const el = this.el!.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!el) return;
    el.hidden = message.length === 0;
    el.textContent = message;
    el.classList.remove('step-dialog__message--error');
  }

  private setError(message: string): void {
    const el = this.el!.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!el) return;
    el.hidden = false;
    el.textContent = message.length > 400 ? message.slice(0, 400) + '…' : message;
    el.classList.add('step-dialog__message--error');
    this.setBusy(false, 'Error');
  }

  private renderMetadata(): void {
    const nameEl = this.el!.querySelector('.step-dialog__name') as HTMLElement | null;
    if (!nameEl || !this.metadata) return;
    const parts: string[] = [];
    if (this.metadata.timeLastModified) parts.push(`Modified ${new Date(this.metadata.timeLastModified).toLocaleString()}`);
    if (this.metadata.length) parts.push(formatBytes(Number(this.metadata.length)));
    parts.push('occt-import-js renderer (LGPL-2.1)');
    nameEl.title = `${this.fileName}\n${parts.join(' | ')}`;
  }

  // ── Styles ────────────────────────────────────────────────────────────────

  private ensureStyles(): void {
    if (this.el!.querySelector('style[data-bpf-preview-style="step"]')) return;
    const style = document.createElement('style');
    style.dataset.bpfPreviewStyle = 'step';
    style.textContent = `
      .step-dialog {
        background:#1a1f2e; color:#f5f5f5; display:flex; flex-direction:column;
        font-family:"Segoe UI",Arial,sans-serif; height:100%; min-height:0;
        overflow:hidden; width:100%;
      }
      .step-dialog__body {
        display:flex; flex:1 1 auto; min-height:0; overflow:hidden; position:relative;
      }
      .step-dialog__canvas {
        background:#1a1f2e; flex:1 1 auto; min-height:0; min-width:0;
        overflow:hidden; position:relative;
      }
      .step-dialog__canvas canvas { display:block; }
      .step-dialog__header {
        align-items:center; background:#1b1b1b;
        border-bottom:1px solid rgba(255,255,255,.12);
        display:flex; flex:0 0 52px; gap:16px; justify-content:space-between;
        min-height:52px; padding:6px 12px 6px 16px;
      }
      .step-dialog__title,.step-dialog__actions {
        align-items:center; display:flex; gap:8px; min-width:0;
      }
      .step-dialog__title { flex:1 1 auto; }
      .step-dialog__actions { flex:0 0 auto; flex-wrap:nowrap; justify-content:flex-end; }
      .step-dialog__badge {
        background:#3d2b1f; border:1px solid rgba(255,255,255,.18); border-radius:4px;
        color:#f4a261; flex:0 0 auto; font-size:13px; font-weight:700; padding:6px 8px;
      }
      .step-dialog__name {
        font-size:15px; font-weight:600; overflow:hidden;
        text-overflow:ellipsis; white-space:nowrap;
      }
      .step-dialog__status { color:#a6a6a6; font-size:12px; white-space:nowrap; }
      .step-dialog__button,.step-dialog__close {
        align-items:center; background:#242424; border:1px solid rgba(255,255,255,.2);
        color:#fff; cursor:pointer; display:inline-flex; font:inherit; justify-content:center;
      }
      .step-dialog__button {
        border-radius:4px; height:36px; min-width:36px; padding:0; width:36px;
      }
      .step-dialog__button svg { display:block; height:18px; width:18px; }
      .step-dialog__button:hover:not(:disabled),.step-dialog__close:hover {
        background:rgba(255,255,255,.1);
      }
      .step-dialog__button:disabled { color:#777; cursor:not-allowed; }
      .step-dialog__close {
        border-radius:4px; font-size:24px; height:36px;
        line-height:1; padding:0 0 3px; width:36px;
      }
      .step-dialog__message {
        flex:0 0 auto; border-bottom:1px solid rgba(0,0,0,.1);
        color:#c7343d; font-size:13px; max-height:8em; overflow-y:auto; padding:8px 20px;
      }
      .step-dialog__message--error { background:#3d1c1c; color:#f9aaa4; }

      /* ── Model tree panel ── */
      .step-tree {
        background:#13161e; border-right:1px solid rgba(255,255,255,.09);
        display:flex; flex:0 0 268px; flex-direction:column;
        min-height:0; min-width:268px; overflow:hidden;
        transition:flex-basis .18s ease, min-width .18s ease;
      }
      .step-tree__header {
        align-items:center; background:#1b1f2b;
        border-bottom:1px solid rgba(255,255,255,.07);
        display:flex; flex:0 0 40px; justify-content:space-between;
        min-height:40px; padding:0 4px 0 12px;
      }
      .step-tree__title {
        color:#7c8495; font-size:10.5px; font-weight:600;
        letter-spacing:.07em; text-transform:uppercase;
      }
      .step-tree__hdr-actions { display:flex; gap:2px; }
      .step-tree__hdr-btn {
        align-items:center; background:transparent; border:none; border-radius:4px;
        color:#7c8495; cursor:pointer; display:inline-flex;
        height:28px; justify-content:center; padding:0; width:28px;
      }
      .step-tree__hdr-btn:hover { background:rgba(255,255,255,.08); color:#e5e7eb; }
      .step-tree__hdr-btn svg { display:block; height:14px; width:14px; }
      .step-tree__search { flex:0 0 auto; padding:6px 8px 4px; }
      .step-tree__search-input {
        background:rgba(255,255,255,.06); border:1px solid rgba(255,255,255,.11);
        border-radius:4px; box-sizing:border-box; color:#e5e7eb;
        font:12px "Segoe UI",Arial,sans-serif; outline:none;
        padding:5px 8px; width:100%;
      }
      .step-tree__search-input::placeholder { color:#4b5563; }
      .step-tree__search-input:focus { border-color:rgba(68,136,255,.6); }
      .step-tree__scroll {
        flex:1 1 auto; min-height:0; overflow-x:hidden; overflow-y:auto; padding:4px 0;
      }
      .step-tree__scroll::-webkit-scrollbar { width:5px; }
      .step-tree__scroll::-webkit-scrollbar-track { background:transparent; }
      .step-tree__scroll::-webkit-scrollbar-thumb {
        background:rgba(255,255,255,.12); border-radius:3px;
      }
      .step-tree__footer {
        border-top:1px solid rgba(255,255,255,.05); color:#4b5563;
        flex:0 0 auto; font-size:10.5px; padding:5px 12px;
      }

      /* ── Tree nodes ── */
      .stree-node { user-select:none; }
      .stree-node__row {
        align-items:center; border-radius:3px; cursor:pointer;
        display:flex; gap:3px; height:26px;
        margin:1px 4px; padding-right:4px; position:relative;
      }
      .stree-node__row:hover { background:rgba(255,255,255,.06); }
      .stree-node--selected .stree-node__row { background:rgba(68,136,255,.18); }
      .stree-node--selected .stree-node__row:hover { background:rgba(68,136,255,.25); }
      .stree-node--faded .stree-node__label { color:#374151; }
      .stree-node--faded .stree-node__icon { opacity:.3; }
      .stree-node--faded .stree-node__swatch { opacity:.25; }

      .stree-node__toggle {
        align-items:center; background:transparent; border:none;
        border-radius:2px; color:#6b7280; cursor:pointer; display:inline-flex;
        flex:0 0 16px; height:18px; justify-content:center; padding:0; width:16px;
      }
      .stree-node__toggle:hover { color:#e5e7eb; }
      .stree-node__toggle svg { display:block; height:11px; width:11px; }
      .stree-node__toggle--leaf { cursor:default; opacity:0; pointer-events:none; }

      .stree-node__icon {
        align-items:center; color:#6b7280; display:inline-flex;
        flex:0 0 16px; height:16px; justify-content:center; width:16px;
      }
      .stree-node--selected .stree-node__icon { color:#7eb5ff; }
      .stree-node__icon svg { display:block; height:13px; width:13px; }

      .stree-node__swatch {
        border-radius:50%; display:inline-block;
        flex:0 0 8px; height:8px; width:8px;
      }
      .stree-node__label {
        color:#c9cdd6; flex:1 1 auto; font-size:12.5px;
        overflow:hidden; text-overflow:ellipsis; white-space:nowrap;
      }
      .stree-node--selected .stree-node__label { color:#fff; font-weight:500; }
      .stree-node__count {
        background:rgba(255,255,255,.07); border-radius:10px; color:#525c6b;
        flex:0 0 auto; font-size:10px; padding:1px 5px;
      }
      .stree-node__eye {
        align-items:center; background:transparent; border:none; border-radius:3px;
        color:rgba(255,255,255,.25); cursor:pointer; display:inline-flex;
        flex:0 0 20px; height:20px; justify-content:center;
        opacity:0; padding:0; width:20px;
      }
      .stree-node__eye svg { display:block; height:11px; width:11px; }
      .stree-node__row:hover .stree-node__eye { opacity:1; }
      .stree-node__eye:hover { color:#e5e7eb; }
      .stree-node__eye--off { color:rgba(248,113,113,.7); opacity:1; }
      .stree-node__eye--off:hover { color:#f87171; }

      /* ── Expand-panel stub ── */
      .step-tree__expand-btn {
        align-items:center; background:#1b1f2b;
        border:none; border-right:1px solid rgba(255,255,255,.09);
        color:#6b7280; cursor:pointer; display:flex;
        flex:0 0 22px; height:100%; justify-content:center;
        min-height:0; padding:0; width:22px;
      }
      .step-tree__expand-btn:hover { background:rgba(255,255,255,.07); color:#e5e7eb; }
      .step-tree__expand-btn svg { display:block; height:13px; width:13px; }

      .step-dialog:fullscreen { min-height:100dvh; }
    `;
    this.el!.appendChild(style);
  }
}

function capitalizeName(s: string): string {
  if (!s) return s;
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase().replace(/-/g, ' ').replace(/_/g, ' ');
}

function formatBytes(value: number): string {
  if (!Number.isFinite(value)) return '';
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${Math.round(value / 102.4) / 10} KB`;
  return `${Math.round(value / 1024 / 102.4) / 10} MB`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}
