import { SPHttpClient } from '@microsoft/sp-http';
import { BaseDialog, type IDialogConfiguration } from '@microsoft/sp-dialog';
import type { IFileExtensionSettings } from './previewSettings';
import { SharePointFileService, type ISharePointFileMetadata } from './sharePointFileService';
import { renderIcon } from '../../shared/icons';

// web-ifc and three are loaded dynamically to keep 3D libraries out of the main bundle.
// web-ifc license: MIT — https://github.com/ifcjs/web-ifc
// three license: MIT — https://github.com/mrdoob/three.js
type WebIFCModule = typeof import('web-ifc');
type ThreeModule = typeof import('three');
type OrbitControlsModule = typeof import('three/examples/jsm/controls/OrbitControls');
type BgUtils = typeof import('three/examples/jsm/utils/BufferGeometryUtils');

const WEB_IFC_CDN = 'https://cdn.jsdelivr.net/npm/web-ifc@0.0.77/';
const GEOMETRY_CHUNK = 250; // meshes to build per animation frame to keep UI responsive

interface RawGeom {
  vertices: Float32Array; // interleaved [x,y,z, nx,ny,nz] per vertex
  indices: Uint32Array;
  colorKey: string;
  r: number;
  g: number;
  b: number;
  a: number;
  transform: number[]; // column-major 4x4
  expressID: number;   // IFC element ID — kept for per-element mesh + property queries
}

export class IfcViewerDialog extends BaseDialog {
  private fileService: SharePointFileService;
  private metadata: ISharePointFileMetadata | undefined;
  private cancelled: boolean = false;
  private animFrameId: number = 0;
  private renderer: import('three').WebGLRenderer | undefined;
  private scene: import('three').Scene | undefined;
  private camera: import('three').PerspectiveCamera | undefined;
  private orbitControls: import('three/examples/jsm/controls/OrbitControls').OrbitControls | undefined;
  private resizeObserver: ResizeObserver | undefined;
  // Keep IFC model alive for property queries after geometry is built
  private ifcApiInstance: InstanceType<WebIFCModule['IfcAPI']> | undefined;
  private openModelID: number = -1;
  private webIfcModule: WebIFCModule | undefined;
  // Selection state
  private selectedMesh: import('three').Mesh | undefined;
  private originalMeshMaterial: import('three').Material | import('three').Material[] | undefined;
  private raycaster: import('three').Raycaster | undefined;
  private threeModule: ThreeModule | undefined;

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
    const badge = this.extensionSettings.extension.replace('.', '').toUpperCase();
    this.domElement.style.cssText =
      'box-sizing:border-box;display:flex;flex-direction:column;height:100dvh;inset:0;overflow:hidden;position:fixed;width:100vw;z-index:2147483647;';
    this.makeFullViewport();
    window.requestAnimationFrame(() => this.makeFullViewport());
    window.setTimeout(() => this.makeFullViewport(), 300);

    this.domElement.innerHTML = `
      <div class="ifc-dialog">
        <div class="ifc-dialog__header">
          <div class="ifc-dialog__title">
            <span class="ifc-dialog__badge">${escapeHtml(badge)}</span>
            <span class="ifc-dialog__name" title="${escapeHtml(this.fileName)}">${escapeHtml(this.fileName)}</span>
            <span class="ifc-dialog__status" data-role="status">Loading</span>
          </div>
          <div class="ifc-dialog__actions">
            <button class="ifc-dialog__button" data-action="reload" type="button" aria-label="Reload" title="Reload">${renderIcon('refresh')}</button>
            <button class="ifc-dialog__button" data-action="fit" type="button" aria-label="Fit to screen" title="Fit to screen">${renderIcon('maximize')}</button>
            <button class="ifc-dialog__button" data-action="fullscreen" type="button" aria-label="Open full screen" title="Open full screen">${renderIcon('external')}</button>
            <button class="ifc-dialog__close" type="button" aria-label="Close preview" title="Close">&times;</button>
          </div>
        </div>
        <div class="ifc-dialog__message" data-role="message" hidden></div>
        <div class="ifc-dialog__canvas" data-role="canvas">
          <div class="ifc-inspector" data-role="inspector" hidden>
            <div class="ifc-inspector__header">
              <span class="ifc-inspector__title">Properties</span>
              <button class="ifc-inspector__close" data-action="close-inspector" type="button" aria-label="Close properties panel" title="Close">&times;</button>
            </div>
            <div class="ifc-inspector__body" data-role="inspector-content"></div>
          </div>
          <div class="ifc-dialog__hint" data-role="hint" hidden>Click any element to view its properties</div>
        </div>
      </div>
    `;

    this.ensureStyles();
    this.wireEvents();
    this.load().catch((error: unknown) => {
      if (!this.cancelled) {
        this.setError(error instanceof Error ? error.message : 'Could not open IFC file.');
      }
    });
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: false };
  }

  protected onAfterClose(): void {
    this.cancelled = true;
    this.teardownScene();
    this.exitFullscreen().catch(() => undefined);
    super.onAfterClose();
  }

  private async load(): Promise<void> {
    this.teardownScene();
    this.cancelled = false;
    this.setBusy(true, 'Downloading file…');
    this.setMessage('');

    const [metadata, arrayBuffer] = await Promise.all([
      this.fileService.getMetadata(this.serverRelativeUrl),
      this.fileService.getContentAsArrayBuffer(this.serverRelativeUrl)
    ]);
    if (this.cancelled) {
      return;
    }

    this.metadata = metadata;
    this.renderMetadata();

    // Guard against files that would exhaust WASM memory.
    const fileSizeBytes = arrayBuffer.byteLength;
    if (fileSizeBytes > 300 * 1024 * 1024) {
      throw new Error(
        `This IFC file is ${formatBytes(fileSizeBytes)}, which exceeds the 300 MB browser processing limit. ` +
        'Use a dedicated BIM application to open very large files.'
      );
    }

    this.setBusy(true, 'Loading libraries…');
    const [webIfc, three, orbitMod, bgUtils] = await Promise.all([
      import(/* webpackChunkName: 'web-ifc' */ 'web-ifc') as Promise<WebIFCModule>,
      import(/* webpackChunkName: 'three' */ 'three') as Promise<ThreeModule>,
      import(/* webpackChunkName: 'three-orbit-controls' */ 'three/examples/jsm/controls/OrbitControls') as Promise<OrbitControlsModule>,
      import(/* webpackChunkName: 'three-bg-utils' */ 'three/examples/jsm/utils/BufferGeometryUtils') as Promise<BgUtils>
    ]);
    if (this.cancelled) {
      return;
    }

    this.threeModule = three;

    this.setBusy(true, 'Initialising IFC engine…');
    const ifcApi = new webIfc.IfcAPI();
    ifcApi.SetWasmPath(WEB_IFC_CDN, true);
    try {
      // forceSingleThread=true: SharePoint Online does not set COOP/COEP headers,
      // so SharedArrayBuffer is unavailable — the MT WASM would fail to init.
      await ifcApi.Init(undefined, true);
    } catch (wasmErr: unknown) {
      console.error('[IfcViewerDialog] WASM init failed:', wasmErr);
      throw new Error(
        'The IFC engine (web-ifc WASM) failed to initialise. ' +
        'Check that the browser can reach cdn.jsdelivr.net and that WebAssembly is not blocked by your security policy.'
      );
    }
    if (this.cancelled) {
      return;
    }

    this.setBusy(true, 'Parsing IFC…');
    const uint8 = new Uint8Array(arrayBuffer);
    // IFC/STEP (ISO-10303-21) is ASCII-only throughout.
    // Non-compliant exporters sometimes write raw bytes > 127 in string literals or token positions.
    // See sanitizeIfcData() doc-comment for the full strategy.
    const ifcData = sanitizeIfcData(uint8);
    let modelID: number;
    try {
      modelID = ifcApi.OpenModel(ifcData, { COORDINATE_TO_ORIGIN: true });
    } catch (parseErr: unknown) {
      console.error('[IfcViewerDialog] OpenModel failed:', parseErr);
      throw new Error(
        'The IFC file could not be parsed. Ensure the file is a valid IFC 2×3, IFC 4, or IFC 4.3 file.'
      );
    }
    if (modelID < 0) {
      throw new Error(
        'The IFC file could not be parsed. Ensure the file is a valid IFC 2×3, IFC 4, or IFC 4.3 file.'
      );
    }

    // Detect schema version for user-facing diagnostics (non-critical).
    let detectedSchema = '';
    try {
      detectedSchema = ifcApi.GetModelSchema(modelID) ?? '';
    } catch {
      // not critical — continue without schema info
    }

    if (this.cancelled) {
      ifcApi.CloseModel(modelID);
      return;
    }

    // Collect raw geometry synchronously from WASM memory.
    // We must copy typed arrays immediately — WASM heap is reused between calls.
    // Geometry is grouped per IFC element (expressID) so each element gets its own
    // Three.js Mesh with userData.expressID, enabling raycaster-based click-to-inspect.
    this.setBusy(true, 'Collecting geometry…');
    const rawGeoms: RawGeom[] = [];
    let streamError: unknown;
    try {
      ifcApi.StreamAllMeshes(modelID, (flatMesh) => {
        const eid = flatMesh.expressID;
        const geomCount = flatMesh.geometries.size();
        for (let j = 0; j < geomCount; j++) {
          const placed = flatMesh.geometries.get(j);
          if (!placed || !placed.geometryExpressID) {
            continue;
          }
          let geom: import('web-ifc').IfcGeometry | undefined;
          try {
            geom = ifcApi.GetGeometry(modelID, placed.geometryExpressID);
            if (!geom) {
              continue;
            }
            const vertexSize = geom.GetVertexDataSize();
            const indexSize = geom.GetIndexDataSize();
            if (vertexSize === 0 || indexSize === 0) {
              continue;
            }
            const vRaw = ifcApi.GetVertexArray(geom.GetVertexData(), vertexSize);
            const iRaw = ifcApi.GetIndexArray(geom.GetIndexData(), indexSize);
            if (!vRaw || !iRaw || vRaw.length === 0 || iRaw.length === 0) {
              continue;
            }
            const c = placed.color;
            rawGeoms.push({
              vertices: new Float32Array(vRaw),
              indices: new Uint32Array(iRaw),
              colorKey: colorKey(c.x, c.y, c.z, c.w),
              r: c.x,
              g: c.y,
              b: c.z,
              a: c.w,
              transform: placed.flatTransformation.slice(),
              expressID: eid
            });
          } catch (geomErr: unknown) {
            console.warn('[IfcViewerDialog] Skipping geometry', placed.geometryExpressID, geomErr);
          } finally {
            if (geom) {
              try { geom.delete(); } catch { /* ignore WASM object cleanup errors */ }
            }
          }
        }
        try { flatMesh.delete(); } catch { /* ignore */ }
      });
    } catch (err: unknown) {
      streamError = err;
    }

    if (streamError) {
      ifcApi.CloseModel(modelID);
      const msg = streamError instanceof Error ? streamError.message : String(streamError);
      const isMemError = /memory|out of bounds|RuntimeError/i.test(msg);
      throw new Error(
        isMemError
          ? 'A memory error occurred while processing this IFC file. ' +
            'The model may be too complex or contain geometry that exceeds browser limits.'
          : `Geometry streaming failed: ${msg}`
      );
    }

    if (this.cancelled) {
      ifcApi.CloseModel(modelID);
      return;
    }

    if (rawGeoms.length === 0) {
      ifcApi.CloseModel(modelID); // nothing to click; free WASM memory
      const schemaNote = detectedSchema ? ` (${detectedSchema})` : '';
      this.setBusy(false, `IFC viewer — no geometry found${schemaNote}`);
      this.setMessage(
        `No renderable 3D geometry was extracted from this file${schemaNote}. ` +
        'This is common with: ' +
        '(1) MEP / HVAC models — duct segments, pipe fittings, and equipment use swept-solid ' +
        'and profile representations that require a dedicated MEP viewer (e.g. Autodesk Viewer, ' +
        'BIMcollab Zoom); ' +
        '(2) Models exported with geometry visibility disabled; ' +
        '(3) Spatial-structure or annotation-only files.'
      );
      return;
    }

    // Keep the IFC model alive for property queries on element click
    this.ifcApiInstance = ifcApi;
    this.openModelID = modelID;
    this.webIfcModule = webIfc;

    // Build Three.js scene
    const { scene, camera, renderer } = this.initScene(three);
    this.scene = scene;
    this.camera = camera;
    this.renderer = renderer;

    // Group raw geoms by expressID — each IFC element gets one merged mesh.
    // This preserves element identity for raycaster picking while still merging
    // the (usually 1-3) placed geometries per element into a single draw call.
    const groups = new Map<number, { geoms: import('three').BufferGeometry[]; r: number; g: number; b: number; a: number }>();

    const total = rawGeoms.length;
    for (let i = 0; i < total; i += GEOMETRY_CHUNK) {
      if (this.cancelled) {
        return;
      }

      await yieldToMain();
      this.setStatus(`Building geometry… ${Math.round((i / total) * 100)}%`);

      const end = Math.min(i + GEOMETRY_CHUNK, total);
      for (let k = i; k < end; k++) {
        const raw = rawGeoms[k];
        const threeGeom = buildBufferGeometry(raw, three);
        let group = groups.get(raw.expressID);
        if (!group) {
          group = { geoms: [], r: raw.r, g: raw.g, b: raw.b, a: raw.a };
          groups.set(raw.expressID, group);
        }
        group.geoms.push(threeGeom);
      }
    }

    if (this.cancelled) {
      return;
    }

    this.setStatus('Merging geometry…');
    await yieldToMain();

    const groupEntries = Array.from(groups.entries());
    for (let gi = 0; gi < groupEntries.length; gi++) {
      if (this.cancelled) {
        return;
      }

      const [eid, group] = groupEntries[gi];
      const merged = bgUtils.mergeGeometries(group.geoms, false);
      for (let di = 0; di < group.geoms.length; di++) {
        group.geoms[di].dispose();
      }

      if (!merged) {
        continue;
      }

      const material = new three.MeshLambertMaterial({
        color: new three.Color(group.r, group.g, group.b),
        transparent: group.a < 0.99,
        opacity: group.a,
        side: three.DoubleSide
      });
      const mesh = new three.Mesh(merged, material);
      mesh.userData.expressID = eid;
      scene.add(mesh);
    }

    this.fitCamera(three);
    this.wireRaycaster(renderer.domElement, three);
    this.startLoop(three, orbitMod);

    // Show usage hint for element inspection
    const hintEl = this.domElement.querySelector('[data-role="hint"]') as HTMLElement | null;
    if (hintEl) {
      hintEl.hidden = false;
      window.setTimeout(() => { hintEl.hidden = true; }, 6000);
    }

    const schemaLabel = detectedSchema ? `, ${detectedSchema}` : '';
    const elementCount = groups.size;
    this.setBusy(false, `IFC viewer — ${elementCount.toLocaleString()} elements (web-ifc${schemaLabel}, MIT)`);
  }

  private wireRaycaster(canvas: HTMLCanvasElement, three: ThreeModule): void {
    this.raycaster = new three.Raycaster();
    let mouseDownX = 0;
    let mouseDownY = 0;

    canvas.addEventListener('mousedown', (e: MouseEvent) => {
      mouseDownX = e.clientX;
      mouseDownY = e.clientY;
    });

    canvas.addEventListener('click', (e: MouseEvent) => {
      // Ignore if user dragged (orbited) — only fire on actual clicks
      if (Math.abs(e.clientX - mouseDownX) > 5 || Math.abs(e.clientY - mouseDownY) > 5) {
        return;
      }
      this.handleCanvasClick(e, three);
    });
  }

  private handleCanvasClick(e: MouseEvent, three: ThreeModule): void {
    if (!this.raycaster || !this.camera || !this.scene || !this.renderer) {
      return;
    }

    const canvas = this.renderer.domElement;
    const rect = canvas.getBoundingClientRect();
    const mouse = new three.Vector2(
      ((e.clientX - rect.left) / rect.width) * 2 - 1,
      -((e.clientY - rect.top) / rect.height) * 2 + 1
    );

    this.raycaster.setFromCamera(mouse, this.camera);

    // Only test meshes that have an expressID (IFC elements — not lights etc.)
    const pickTargets: import('three').Mesh[] = [];
    this.scene.children.forEach((obj) => {
      const m = obj as import('three').Mesh;
      if (m.isMesh && m.userData.expressID !== undefined) {
        pickTargets.push(m);
      }
    });

    const intersects = this.raycaster.intersectObjects(pickTargets, false);
    if (intersects.length === 0) {
      this.clearInspector();
      return;
    }

    const hit = intersects[0].object as import('three').Mesh;
    const expressID = hit.userData.expressID as number;
    this.highlightMesh(hit, three);
    this.showInspector(expressID);
  }

  private highlightMesh(mesh: import('three').Mesh, three: ThreeModule): void {
    this.restoreHighlight();
    this.selectedMesh = mesh;
    this.originalMeshMaterial = mesh.material;
    const hlMat = new three.MeshLambertMaterial({
      color: new three.Color(0x4fc3f7),
      emissive: new three.Color(0x1565c0),
      emissiveIntensity: 0.4,
      side: three.DoubleSide
    });
    mesh.material = hlMat;
  }

  private restoreHighlight(): void {
    if (this.selectedMesh && this.originalMeshMaterial !== undefined) {
      const oldMat = this.selectedMesh.material;
      this.selectedMesh.material = this.originalMeshMaterial;
      if (oldMat && !Array.isArray(oldMat)) {
        (oldMat as import('three').Material).dispose();
      }
      this.selectedMesh = undefined;
      this.originalMeshMaterial = undefined;
    }
  }

  private showInspector(expressID: number): void {
    const inspectorEl = this.domElement.querySelector('[data-role="inspector"]') as HTMLElement | null;
    const contentEl = this.domElement.querySelector('[data-role="inspector-content"]') as HTMLElement | null;
    if (!inspectorEl || !contentEl) {
      return;
    }

    contentEl.innerHTML = this.renderProperties(expressID);
    inspectorEl.hidden = false;

    const hintEl = this.domElement.querySelector('[data-role="hint"]') as HTMLElement | null;
    if (hintEl) hintEl.hidden = true;
  }

  private clearInspector(): void {
    this.restoreHighlight();
    const inspectorEl = this.domElement.querySelector('[data-role="inspector"]') as HTMLElement | null;
    if (inspectorEl) {
      inspectorEl.hidden = true;
    }
  }

  private renderProperties(expressID: number): string {
    if (!this.ifcApiInstance || this.openModelID < 0) {
      return '<p class="ifc-inspector__error">Model not available.</p>';
    }

    let line: Record<string, unknown>;
    try {
      line = this.ifcApiInstance.GetLine(this.openModelID, expressID, false) as Record<string, unknown>;
    } catch {
      return '<p class="ifc-inspector__error">Could not load properties for this element.</p>';
    }

    const typeNum = typeof line.type === 'number' ? line.type : 0;
    const typeName = this.getIfcTypeName(typeNum);

    const rows: string[] = [];
    const lineKeys = Object.keys(line);
    for (let ki = 0; ki < lineKeys.length; ki++) {
      const key = lineKeys[ki];
      if (key === 'type' || key === 'expressID') continue;
      const val = extractIfcValue(line[key]);
      if (val === undefined) continue;
      rows.push(
        `<div class="ifc-inspector__prop">` +
        `<dt class="ifc-inspector__key">${escapeHtml(key)}</dt>` +
        `<dd class="ifc-inspector__val" title="${escapeHtml(String(val))}">${escapeHtml(String(val))}</dd>` +
        `</div>`
      );
    }

    const typeLabel = typeName || `Type ${typeNum}`;
    return `
      <div class="ifc-inspector__entity">
        <span class="ifc-inspector__type-badge">${escapeHtml(typeLabel)}</span>
        <span class="ifc-inspector__eid">#${expressID}</span>
      </div>
      ${rows.length > 0
        ? `<dl class="ifc-inspector__props">${rows.join('')}</dl>`
        : '<p class="ifc-inspector__empty">No displayable properties.</p>'}
    `;
  }

  private getIfcTypeName(typeNum: number): string {
    if (!this.webIfcModule || typeNum === 0) return '';
    const mod = this.webIfcModule as unknown as Record<string, unknown>;
    const keys = Object.keys(mod);
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      if (typeof mod[key] === 'number' && mod[key] === typeNum && key.startsWith('IFC')) {
        return key;
      }
    }
    return '';
  }

  private initScene(
    three: ThreeModule
  ): {
    scene: import('three').Scene;
    camera: import('three').PerspectiveCamera;
    renderer: import('three').WebGLRenderer;
  } {
    const canvasEl = this.domElement.querySelector('[data-role="canvas"]') as HTMLElement;

    const scene = new three.Scene();
    scene.background = new three.Color(0x1e1e1e);
    scene.add(new three.AmbientLight(0xffffff, 0.6));
    const dir = new three.DirectionalLight(0xffffff, 0.8);
    dir.position.set(1, 2, 3);
    scene.add(dir);

    const w = canvasEl.clientWidth || 800;
    const h = canvasEl.clientHeight || 600;
    const camera = new three.PerspectiveCamera(60, w / h, 0.001, 10000);
    camera.position.set(0, 5, 10);

    const renderer = new three.WebGLRenderer({ antialias: true });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
    renderer.setSize(w, h);
    canvasEl.appendChild(renderer.domElement);
    renderer.domElement.style.cssText = 'display:block;width:100%;height:100%;';

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

  private fitCamera(three: ThreeModule): void {
    if (!this.scene || !this.camera) {
      return;
    }

    const box = new three.Box3().setFromObject(this.scene);
    if (box.isEmpty()) {
      return;
    }

    const center = new three.Vector3();
    const size = new three.Vector3();
    box.getCenter(center);
    box.getSize(size);

    const maxDim = Math.max(size.x, size.y, size.z);
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

  private startLoop(three: ThreeModule, orbitMod: OrbitControlsModule): void {
    if (!this.renderer || !this.camera || !this.scene) {
      return;
    }

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
    // Close the IFC model first (frees WASM memory)
    if (this.ifcApiInstance !== undefined && this.openModelID >= 0) {
      try { this.ifcApiInstance.CloseModel(this.openModelID); } catch { /* ignore */ }
      this.ifcApiInstance = undefined;
      this.openModelID = -1;
      this.webIfcModule = undefined;
    }

    if (this.animFrameId) {
      window.cancelAnimationFrame(this.animFrameId);
      this.animFrameId = 0;
    }
    this.resizeObserver?.disconnect();
    this.orbitControls?.dispose();
    this.raycaster = undefined;
    this.selectedMesh = undefined;
    this.originalMeshMaterial = undefined;
    if (this.scene) {
      this.scene.traverse((obj) => {
        const mesh = obj as import('three').Mesh;
        mesh.geometry?.dispose();
        if (mesh.material) {
          if (Array.isArray(mesh.material)) {
            (mesh.material as import('three').Material[]).forEach((m) => m.dispose());
          } else {
            (mesh.material as import('three').Material).dispose();
          }
        }
      });
    }
    this.renderer?.dispose();
    this.scene = undefined;
    this.camera = undefined;
    this.renderer = undefined;
    this.orbitControls = undefined;
    this.threeModule = undefined;
  }

  private wireEvents(): void {
    this.domElement.querySelector('[data-action="reload"]')?.addEventListener('click', () => {
      this.cancelled = false;
      this.load().catch((e: unknown) => this.setError(e instanceof Error ? e.message : 'Could not reload.'));
    });

    this.domElement.querySelector('[data-action="fit"]')?.addEventListener('click', () => {
      import(/* webpackChunkName: 'three' */ 'three').then((three: ThreeModule) => this.fitCamera(three)).catch(() => undefined);
    });

    this.domElement.querySelector('[data-action="fullscreen"]')?.addEventListener('click', () => {
      this.toggleFullscreen().catch((e: unknown) => this.setError(e instanceof Error ? e.message : 'Could not toggle full screen.'));
    });

    this.domElement.querySelector('[data-action="close-inspector"]')?.addEventListener('click', () => {
      this.clearInspector();
    });

    document.addEventListener('fullscreenchange', () => this.updateFullscreenButton());

    this.domElement.querySelector('.ifc-dialog__close')?.addEventListener('click', () => {
      this.close().catch(() => undefined);
    });
  }

  private async toggleFullscreen(): Promise<void> {
    if (document.fullscreenElement) {
      await this.exitFullscreen();
    } else {
      await this.domElement.requestFullscreen();
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
    const btn = this.domElement.querySelector('[data-action="fullscreen"]') as HTMLButtonElement | null;
    if (!btn) {
      return;
    }

    const isFs = Boolean(document.fullscreenElement);
    btn.innerHTML = renderIcon(isFs ? 'restore' : 'external');
    btn.setAttribute('aria-label', isFs ? 'Exit full screen' : 'Open full screen');
    btn.title = isFs ? 'Exit full screen' : 'Open full screen';
  }

  private makeFullViewport(): void {
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
    el.style.removeProperty('transform');
    const rect = el.getBoundingClientRect();
    if (rect.left !== 0 || rect.top !== 0) {
      el.style.setProperty('transform', `translate(${-rect.left}px,${-rect.top}px)`, 'important');
    }
  }

  private ensureStyles(): void {
    if (this.domElement.querySelector('style[data-bpf-preview-style="ifc"]')) {
      return;
    }

    const style = document.createElement('style');
    style.dataset.bpfPreviewStyle = 'ifc';
    style.textContent = `
      .ifc-dialog {
        background: #1e1e1e;
        color: #f5f5f5;
        display: flex;
        flex-direction: column;
        font-family: "Segoe UI", Arial, sans-serif;
        height: 100%;
        min-height: 0;
        overflow: hidden;
        width: 100%;
      }
      .ifc-dialog__header {
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
      .ifc-dialog__title,
      .ifc-dialog__actions {
        align-items: center;
        display: flex;
        gap: 8px;
        min-width: 0;
      }
      .ifc-dialog__title { flex: 1 1 auto; }
      .ifc-dialog__actions { flex: 0 0 auto; flex-wrap: nowrap; justify-content: flex-end; }
      .ifc-dialog__badge {
        background: #1d3557;
        border: 1px solid rgba(255,255,255,.18);
        border-radius: 4px;
        color: #a8dadc;
        flex: 0 0 auto;
        font-size: 13px;
        font-weight: 700;
        padding: 6px 8px;
      }
      .ifc-dialog__name {
        font-size: 15px;
        font-weight: 600;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .ifc-dialog__status { color: #a6a6a6; font-size: 12px; }
      .ifc-dialog__button,
      .ifc-dialog__close {
        align-items: center;
        background: #242424;
        border: 1px solid rgba(255,255,255,.2);
        color: #ffffff;
        cursor: pointer;
        display: inline-flex;
        font: inherit;
        justify-content: center;
      }
      .ifc-dialog__button {
        border-radius: 4px;
        height: 36px;
        min-width: 36px;
        padding: 0;
        width: 36px;
      }
      .ifc-dialog__button svg { display: block; height: 18px; width: 18px; }
      .ifc-dialog__button:hover:not(:disabled),
      .ifc-dialog__close:hover { background: rgba(255,255,255,.1); }
      .ifc-dialog__button:disabled { color: #777; cursor: not-allowed; }
      .ifc-dialog__close {
        border-radius: 4px;
        font-size: 24px;
        height: 36px;
        line-height: 1;
        padding: 0 0 3px;
        width: 36px;
      }
      .ifc-dialog__message {
        flex: 0 0 auto;
        border-bottom: 1px solid rgba(0,0,0,.1);
        color: #c7343d;
        font-size: 13px;
        max-height: 8em;
        overflow-y: auto;
        padding: 8px 20px;
      }
      .ifc-dialog__message--error { background: #3d1c1c; color: #f9aaa4; }
      .ifc-dialog__canvas {
        background: #1e1e1e;
        flex: 1 1 auto;
        min-height: 0;
        overflow: hidden;
        position: relative;
        width: 100%;
      }
      .ifc-dialog__canvas canvas { display: block; }
      /* Usage hint overlay */
      .ifc-dialog__hint {
        background: rgba(0,0,0,.55);
        border-radius: 4px;
        bottom: 16px;
        color: #e0e0e0;
        font-size: 12px;
        left: 50%;
        padding: 6px 14px;
        pointer-events: none;
        position: absolute;
        transform: translateX(-50%);
        white-space: nowrap;
        z-index: 5;
      }
      /* Element inspector panel */
      .ifc-inspector {
        backdrop-filter: blur(6px);
        background: rgba(22, 22, 22, 0.94);
        border-left: 1px solid rgba(255,255,255,.1);
        bottom: 0;
        box-shadow: -4px 0 20px rgba(0,0,0,.4);
        display: flex;
        flex-direction: column;
        max-width: 300px;
        overflow: hidden;
        position: absolute;
        right: 0;
        top: 0;
        width: 280px;
        z-index: 10;
      }
      .ifc-inspector__header {
        align-items: center;
        background: rgba(255,255,255,.05);
        border-bottom: 1px solid rgba(255,255,255,.08);
        display: flex;
        flex: 0 0 auto;
        justify-content: space-between;
        padding: 10px 14px;
      }
      .ifc-inspector__title {
        color: #a8dadc;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: .06em;
        text-transform: uppercase;
      }
      .ifc-inspector__close {
        align-items: center;
        background: none;
        border: none;
        color: #888;
        cursor: pointer;
        display: inline-flex;
        font-size: 20px;
        height: 28px;
        justify-content: center;
        line-height: 1;
        padding: 0 0 2px;
        width: 28px;
      }
      .ifc-inspector__close:hover { color: #fff; }
      .ifc-inspector__body {
        flex: 1 1 auto;
        overflow-y: auto;
        padding: 12px 0;
      }
      .ifc-inspector__entity {
        align-items: baseline;
        display: flex;
        gap: 8px;
        padding: 4px 14px 12px;
      }
      .ifc-inspector__type-badge {
        background: #1d3557;
        border: 1px solid rgba(168,218,220,.3);
        border-radius: 3px;
        color: #a8dadc;
        font-size: 11px;
        font-weight: 700;
        max-width: 180px;
        overflow: hidden;
        padding: 3px 7px;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .ifc-inspector__eid {
        color: #666;
        font-family: Consolas, monospace;
        font-size: 11px;
        white-space: nowrap;
      }
      .ifc-inspector__props {
        border-top: 1px solid rgba(255,255,255,.06);
        display: grid;
        gap: 0;
        margin: 0;
        padding: 0;
      }
      .ifc-inspector__prop {
        border-bottom: 1px solid rgba(255,255,255,.04);
        display: grid;
        gap: 2px;
        padding: 7px 14px;
      }
      .ifc-inspector__prop:hover { background: rgba(255,255,255,.04); }
      .ifc-inspector__key {
        color: #888;
        font-size: 11px;
        font-weight: 600;
        letter-spacing: .02em;
        text-transform: uppercase;
      }
      .ifc-inspector__val {
        color: #e0e0e0;
        font-size: 13px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }
      .ifc-inspector__empty,
      .ifc-inspector__error {
        color: #888;
        font-size: 13px;
        padding: 8px 14px;
      }
      .ifc-inspector__error { color: #f9aaa4; }
      :fullscreen .ifc-dialog { min-height: 100dvh; }
    `;
    this.domElement.appendChild(style);
  }

  private setBusy(isBusy: boolean, status: string): void {
    this.setStatus(status);
    this.domElement.querySelectorAll('.ifc-dialog__button').forEach((btn) => {
      (btn as HTMLButtonElement).disabled = isBusy;
    });
  }

  private setStatus(status: string): void {
    const el = this.domElement.querySelector('[data-role="status"]') as HTMLElement | null;
    if (el) {
      el.textContent = status;
    }
  }

  private setMessage(message: string): void {
    const el = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!el) {
      return;
    }

    el.hidden = message.length === 0;
    el.textContent = message;
    el.classList.remove('ifc-dialog__message--error');
  }

  private setError(message: string): void {
    const el = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!el) {
      return;
    }

    const display = message.length > 400 ? message.slice(0, 400) + '…' : message;
    el.hidden = false;
    el.textContent = display;
    el.classList.add('ifc-dialog__message--error');
    this.setBusy(false, 'Error');
  }

  private renderMetadata(): void {
    const nameEl = this.domElement.querySelector('.ifc-dialog__name') as HTMLElement | null;
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
    parts.push('web-ifc renderer (MIT)');
    nameEl.title = `${this.fileName}\n${parts.join(' | ')}`;
  }
}

/**
 * Remove or neutralise all non-ASCII bytes (> 127) in an IFC/STEP file.
 *
 * IFC (ISO-10303-21 STEP) is ASCII-only throughout. Non-compliant exporters
 * sometimes embed raw Latin-1/Windows-1252 bytes in string literals (accented
 * author names, element descriptions). These confuse web-ifc's C++ tokenizer —
 * logged as "[ArgumentOffset()] unexpected line end <N>". In the worst case the
 * WASM parse fails entirely and returns modelID = -1, crashing the JS wrapper
 * inside OpenModel at GetHeaderLine(-1, FILE_SCHEMA).arguments[0][0].value.
 *
 * Strategy — STEP string-boundary aware:
 *   • INSIDE a single-quoted string literal ('...'): replace bytes > 127 with
 *     ASCII space (0x20). A space is valid STEP whitespace inside a string and
 *     preserves the enclosing quote boundary.
 *   • OUTSIDE a string literal: REMOVE the byte entirely. Bytes > 127 outside
 *     strings are invalid in any STEP token (references, numbers, entity type
 *     names). Inserting a space would split the token — e.g. a raw byte inside
 *     '#155' would give '#1 55' (broken reference) — whereas removal gives '#155'
 *     (the intended value), allowing the C++ parser to continue.
 *
 * STEP escape-quote rule: '' inside a string is a literal apostrophe, not the
 * end of the string. The scanner handles this correctly.
 */
export function sanitizeIfcData(data: Uint8Array): Uint8Array {
  // Fast path: bail immediately if every byte is ASCII.
  let hasDirtyBytes = false;
  for (let i = 0; i < data.length; i++) {
    if (data[i] > 127) { hasDirtyBytes = true; break; }
  }
  if (!hasDirtyBytes) {
    return data;
  }

  // Allocate at most data.length bytes; the actual output may be shorter
  // because bytes outside strings are removed rather than substituted.
  const out = new Uint8Array(data.length);
  let outLen = 0;
  let inString = false; // are we currently inside a STEP single-quoted string?

  for (let i = 0; i < data.length; i++) {
    const b = data[i];

    if (inString) {
      if (b === 39) { // 39 = ASCII single-quote '
        out[outLen++] = b;
        if (i + 1 < data.length && data[i + 1] === 39) {
          // Escaped quote '' — two consecutive quotes represent one literal
          // apostrophe inside the string; consume both and stay in string mode.
          out[outLen++] = 39;
          i++;
        } else {
          inString = false; // closing quote — leave string mode
        }
      } else if (b > 127) {
        out[outLen++] = 32; // replace with space — valid inside a STEP string
      } else {
        out[outLen++] = b;
      }
    } else {
      if (b === 39) { // opening quote — enter string mode
        inString = true;
        out[outLen++] = b;
      } else if (b > 127) {
        // Outside a string: remove the byte entirely.
        // Do NOT write it to out — this is intentional.
      } else {
        out[outLen++] = b;
      }
    }
  }

  return out.subarray(0, outLen);
}

/** Round color components to 2 decimal places for grouping similar colours. */
export function colorKey(r: number, g: number, b: number, a: number): string {
  return `${r.toFixed(2)},${g.toFixed(2)},${b.toFixed(2)},${a.toFixed(2)}`;
}

/** Build an indexed BufferGeometry with applied placement transform from raw interleaved WASM data. */
function buildBufferGeometry(raw: RawGeom, three: ThreeModule): import('three').BufferGeometry {
  const vertexCount = raw.vertices.length / 6;
  const positions = new Float32Array(vertexCount * 3);
  const normals = new Float32Array(vertexCount * 3);

  for (let k = 0; k < vertexCount; k++) {
    positions[k * 3] = raw.vertices[k * 6];
    positions[k * 3 + 1] = raw.vertices[k * 6 + 1];
    positions[k * 3 + 2] = raw.vertices[k * 6 + 2];
    normals[k * 3] = raw.vertices[k * 6 + 3];
    normals[k * 3 + 1] = raw.vertices[k * 6 + 4];
    normals[k * 3 + 2] = raw.vertices[k * 6 + 5];
  }

  const geom = new three.BufferGeometry();
  geom.setAttribute('position', new three.BufferAttribute(positions, 3));
  geom.setAttribute('normal', new three.BufferAttribute(normals, 3));
  geom.setIndex(new three.BufferAttribute(raw.indices, 1));

  const matrix = new three.Matrix4().fromArray(raw.transform);
  geom.applyMatrix4(matrix);
  return geom;
}

/** Extract a displayable primitive value from a raw IFC property field. Returns undefined when nothing displayable. */
function extractIfcValue(val: unknown): string | number | undefined {
  if (val === undefined || val === null) return undefined;
  if (typeof val === 'string') return val || undefined;
  if (typeof val === 'number') return val;
  if (typeof val === 'boolean') return String(val);
  if (typeof val === 'object') {
    const obj = val as { value?: unknown };
    const inner = obj.value;
    if (inner === undefined || inner === null) return undefined;
    if (typeof inner === 'string') return inner || undefined;
    if (typeof inner === 'number') return inner;
    if (typeof inner === 'boolean') return String(inner);
  }
  return undefined;
}

function yieldToMain(): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, 0));
}

export function formatBytes(value: number): string {
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

export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
