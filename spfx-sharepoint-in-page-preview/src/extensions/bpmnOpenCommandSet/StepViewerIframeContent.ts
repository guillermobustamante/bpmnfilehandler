/**
 * Self-contained 3D STEP viewer page rendered inside a blob-URL iframe.
 *
 * A blob-URL iframe creates an isolated CSP context separate from the SharePoint
 * host page. This allows 'wasm-unsafe-eval' which Emscripten-compiled libraries
 * (occt-import-js) require — without changing the tenant's SharePoint CSP policy.
 *
 * All heavy libraries (Three.js, occt-import-js) are loaded directly from the
 * jsdelivr CDN inside the iframe. The parent dialog downloads the STEP file bytes
 * and transfers them in via postMessage.
 *
 * postMessage protocol (bpf:'step' namespace):
 *   parent → iframe : { bpf:'step', type:'load',   buffer: ArrayBuffer }
 *   parent → iframe : { bpf:'step', type:'fit' }
 *   iframe → parent : { bpf:'step', type:'ready' }
 *   iframe → parent : { bpf:'step', type:'status',  text: string }
 *   iframe → parent : { bpf:'step', type:'loaded',  meshCount: number }
 *   iframe → parent : { bpf:'step', type:'error',   message: string }
 */
export const STEP_VIEWER_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy"
  content="default-src 'none';
           script-src 'wasm-unsafe-eval' https://cdn.jsdelivr.net blob: 'unsafe-inline';
           connect-src https://cdn.jsdelivr.net;
           worker-src blob:;
           style-src 'unsafe-inline';">
<style>
  *{margin:0;padding:0;box-sizing:border-box}
  html,body{width:100%;height:100%;overflow:hidden;background:#1a1f2e}
  canvas{display:block;width:100%!important;height:100%!important}
  #ov{
    position:absolute;inset:0;display:flex;flex-direction:column;
    align-items:center;justify-content:center;gap:12px;
    color:#c0c8e0;font:14px/1.6 "Segoe UI",Arial,sans-serif;
    text-align:center;padding:24px;pointer-events:none
  }
  #ov.err{color:#f9aaa4}
  .spin{
    width:28px;height:28px;
    border:3px solid rgba(192,200,224,.2);
    border-top-color:#c0c8e0;
    border-radius:50%;
    animation:sp .8s linear infinite
  }
  .err .spin{display:none}
  @keyframes sp{to{transform:rotate(360deg)}}
</style>
</head>
<body>
<div id="ov"><div class="spin"></div><span id="msg">Waiting for file…</span></div>
<script>
// Classic (non-module) script so 'unsafe-inline' in the meta CSP applies.
// W3C CSP Level 3 §8.2: 'unsafe-inline' is silently ignored for <script type="module">,
// which caused the inline module to be blocked and the 3D libraries to never load.
// Dynamic import() still works from classic scripts — it does not require a module context.
//
// NOTE: Chrome DevTools does NOT expose sandboxed null-origin frames (sandbox="allow-scripts"
// without allow-same-origin) in the frame selector dropdown. console.* calls from this frame
// are also invisible in DevTools. All diagnostic logging therefore goes via postMessage so it
// appears in the parent's [BPF-TOP] console output instead.

var OCCT = 'https://cdn.jsdelivr.net/npm/occt-import-js@0.0.23/dist/';
var THREE_BASE = 'https://cdn.jsdelivr.net/npm/three@0.168.0/';

function setMsg(txt) {
  var el = document.getElementById('msg');
  if (el) el.textContent = txt;
}
function toParent(type, extra) {
  try { window.parent.postMessage(Object.assign({ bpf: 'step', type: type }, extra || {}), '*'); } catch(e) { /* isolated */ }
}
// All debug from this frame routes through postMessage → visible as [BPF-TOP] in parent console.
function dbg(msg) { toParent('debug', { msg: msg }); }
function status(txt) { setMsg(txt); toParent('status', { text: txt }); dbg('status: ' + txt); }
function die(msg) {
  var ov = document.getElementById('ov');
  if (ov) ov.className = 'err';
  setMsg('Error: ' + msg);
  toParent('error', { message: msg });
}

// Catch any uncaught errors and route them back to the parent.
window.onerror = function(msg, src, line, col, err) {
  toParent('debug', { msg: 'window.onerror: ' + msg + ' (' + src + ':' + line + ')' });
  toParent('error', { message: String(msg) });
  return true;
};
window.onunhandledrejection = function(ev) {
  var msg = (ev.reason instanceof Error) ? ev.reason.message : String(ev.reason);
  toParent('debug', { msg: 'unhandledRejection: ' + msg });
  toParent('error', { message: msg });
};

dbg('script start — origin:' + location.origin + ' | href:' + location.href);

// Notify parent that the frame is ready to receive the file bytes.
dbg('sending ready');
toParent('ready');
dbg('ready sent — waiting for load message');

window.addEventListener('message', function(ev) {
  dbg('message received — bpf:' + (ev.data && ev.data.bpf) + ' | type:' + (ev.data && ev.data.type) + ' | hasBuffer:' + !!(ev.data && ev.data.buffer));
  if (!ev.data || ev.data.bpf !== 'step' || ev.data.type !== 'load') return;

  var buf = ev.data.buffer;
  dbg('load message — buf byteLength:' + (buf && buf.byteLength));

  // Async IIFE — equivalent to top-level await in a module, works in classic scripts.
  (async function() {
    try {
      status('Loading 3D libraries…');
      dbg('starting import() of Three.js modules');
      var mods = await Promise.all([
        import(THREE_BASE + 'build/three.module.js'),
        import(THREE_BASE + 'examples/jsm/controls/OrbitControls.js'),
        import(THREE_BASE + 'examples/jsm/utils/BufferGeometryUtils.js')
      ]);
      dbg('Three.js modules loaded');
      var T = mods[0];
      var OrbitControls = mods[1].OrbitControls;
      var mergeGeometries = mods[2].mergeGeometries;

      status('Loading WASM engine…');
      dbg('starting import() of occt-import-js');
      var occtMod = await import(OCCT + 'occt-import-js.js');
      dbg('occt-import-js module loaded, initialising WASM');
      var occt = await occtMod.default({ locateFile: function(p) { return OCCT + p; } });
      dbg('WASM engine ready');

      status('Parsing STEP file…');
      dbg('calling ReadStepFile, buf byteLength:' + buf.byteLength);
      var result = occt.ReadStepFile(new Uint8Array(buf), null);
      dbg('ReadStepFile done — success:' + (result && result.success) + ' | meshes:' + (result && result.meshes && result.meshes.length));
      if (!result || !result.success || !result.meshes || result.meshes.length === 0) {
        throw new Error(
          'No renderable geometry found. ' +
          'Supported schemas: AP203, AP214, AP242 basic solids. ' +
          'Verify the file contains 3D solid bodies in a CAD application.'
        );
      }

      status('Building geometry…');

      // Scene setup
      var scene = new T.Scene();
      scene.background = new T.Color(0x1a1f2e);
      scene.add(new T.AmbientLight(0xffffff, 0.5));
      var d1 = new T.DirectionalLight(0xffffff, 0.8);
      d1.position.set(1, 2, 3);
      scene.add(d1);
      var d2 = new T.DirectionalLight(0xffffff, 0.3);
      d2.position.set(-2, -1, -2);
      scene.add(d2);

      var camera = new T.PerspectiveCamera(60, innerWidth / innerHeight, 0.001, 100000);

      var cvs = document.createElement('canvas');
      document.body.appendChild(cvs);

      var renderer = new T.WebGLRenderer({ canvas: cvs, antialias: true });
      renderer.setPixelRatio(Math.min(devicePixelRatio, 2));
      renderer.setSize(innerWidth, innerHeight);

      // Group meshes by colour to minimise draw calls
      var groups = {};
      for (var i = 0; i < result.meshes.length; i++) {
        var m = result.meshes[i];
        var geo = new T.BufferGeometry();
        geo.setAttribute('position', new T.BufferAttribute(new Float32Array(m.attributes.position.array), 3));
        if (m.attributes.normal) {
          geo.setAttribute('normal', new T.BufferAttribute(new Float32Array(m.attributes.normal.array), 3));
        }
        if (m.index) { geo.setIndex(m.index.array); }
        if (!m.attributes.normal) { geo.computeVertexNormals(); }
        var col = m.color || [0.6, 0.7, 0.8];
        var key = col[0].toFixed(2) + ',' + col[1].toFixed(2) + ',' + col[2].toFixed(2);
        if (!groups[key]) { groups[key] = { geoms: [], r: col[0], g: col[1], b: col[2] }; }
        groups[key].geoms.push(geo);
      }

      var allGeos = [];
      Object.values(groups).forEach(function(grp) {
        var merged = grp.geoms.length > 1 ? mergeGeometries(grp.geoms, false) : grp.geoms[0];
        if (!merged) return;
        allGeos.push(merged);
        scene.add(new T.Mesh(
          merged,
          new T.MeshLambertMaterial({ color: new T.Color(grp.r, grp.g, grp.b), side: T.DoubleSide })
        ));
      });

      // Fit camera to geometry bounding box
      var box = new T.Box3();
      allGeos.forEach(function(g) { g.computeBoundingBox(); box.union(g.boundingBox); });
      var center = box.getCenter(new T.Vector3());
      var size = box.getSize(new T.Vector3());
      var maxDim = Math.max(size.x, size.y, size.z) || 1;
      var fovRad = camera.fov * Math.PI / 180;
      var dist = Math.abs(maxDim / 2 / Math.tan(fovRad / 2)) * 1.5;

      function fitCam() {
        camera.position.set(center.x + dist * 0.6, center.y + dist * 0.5, center.z + dist);
        camera.lookAt(center);
        controls.target.copy(center);
        controls.update();
      }

      camera.position.set(center.x + dist * 0.6, center.y + dist * 0.5, center.z + dist);
      camera.lookAt(center);
      camera.near = dist / 1000;
      camera.far = dist * 20;
      camera.updateProjectionMatrix();

      var controls = new OrbitControls(camera, renderer.domElement);
      controls.target.copy(center);
      controls.enableDamping = true;
      controls.dampingFactor = 0.05;
      controls.update();

      // Remove loading overlay
      var ov = document.getElementById('ov');
      if (ov) ov.remove();

      // Handle resize
      new ResizeObserver(function() {
        camera.aspect = innerWidth / innerHeight;
        camera.updateProjectionMatrix();
        renderer.setSize(innerWidth, innerHeight);
      }).observe(document.documentElement);

      // Handle commands from parent (fit camera)
      window.addEventListener('message', function(e) {
        if (!e.data || e.data.bpf !== 'step') return;
        if (e.data.type === 'fit') { fitCam(); }
      });

      // Render loop
      (function loop() {
        requestAnimationFrame(loop);
        controls.update();
        renderer.render(scene, camera);
      })();

      toParent('loaded', { meshCount: result.meshes.length });

    } catch (err) {
      dbg('caught error: ' + (err instanceof Error ? err.stack || err.message : String(err)));
      die(err instanceof Error ? err.message : String(err));
    }
  })();
});
</script>
</body>
</html>`;
