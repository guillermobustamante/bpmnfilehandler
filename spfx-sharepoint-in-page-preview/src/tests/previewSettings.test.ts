// Pure-function tests for previewSettings.ts.
// These functions do not invoke SharePoint APIs, so no network setup is required.
jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } }
}));

import {
  createDefaultPreviewSettings,
  findExtensionSettings,
  normalizeBaseUrl,
  normalizeExtension,
  normalizeSettings,
  type IPreviewSettings
} from '../extensions/bpmnOpenCommandSet/previewSettings';

// ── createDefaultPreviewSettings ──────────────────────────────────────────────

describe('createDefaultPreviewSettings', () => {
  it('includes .bpmn enabled', () => {
    const defaults = createDefaultPreviewSettings();
    const bpmn = defaults.extensions.find((e) => e.extension === '.bpmn');
    expect(bpmn?.enabled).toBe(true);
    expect(bpmn?.renderer).toBe('bpmn-js');
  });

  it('includes .drawio disabled by default', () => {
    const defaults = createDefaultPreviewSettings();
    const drawio = defaults.extensions.find((e) => e.extension === '.drawio');
    expect(drawio?.enabled).toBe(false);
    expect(drawio?.renderer).toBe('diagrams-net-embed');
  });

  it('includes Mermaid extensions disabled by default', () => {
    const defaults = createDefaultPreviewSettings();
    const mmd = defaults.extensions.find((e) => e.extension === '.mmd');
    const mermaid = defaults.extensions.find((e) => e.extension === '.mermaid');
    expect(mmd?.enabled).toBe(false);
    expect(mmd?.renderer).toBe('mermaid-js');
    expect(mermaid?.enabled).toBe(false);
    expect(mermaid?.renderer).toBe('mermaid-js');
  });

  it('includes .ifc disabled by default', () => {
    const defaults = createDefaultPreviewSettings();
    const ifc = defaults.extensions.find((e) => e.extension === '.ifc');
    expect(ifc?.enabled).toBe(false);
    expect(ifc?.renderer).toBe('web-ifc');
  });

  it('includes .step and .stp disabled by default', () => {
    const defaults = createDefaultPreviewSettings();
    const step = defaults.extensions.find((e) => e.extension === '.step');
    const stp = defaults.extensions.find((e) => e.extension === '.stp');
    expect(step?.enabled).toBe(false);
    expect(step?.renderer).toBe('occt-step');
    expect(stp?.enabled).toBe(false);
    expect(stp?.renderer).toBe('occt-step');
  });

  it('sets schemaVersion to 1', () => {
    expect(createDefaultPreviewSettings().schemaVersion).toBe(1);
  });

  it('uses the provided appBaseUrl', () => {
    expect(createDefaultPreviewSettings('https://example.com').appBaseUrl).toBe('https://example.com');
  });
});

// ── normalizeSettings ─────────────────────────────────────────────────────────

describe('normalizeSettings', () => {
  it('returns defaults when called with undefined', () => {
    const result = normalizeSettings(undefined, '');
    expect(result.schemaVersion).toBe(1);
    expect(result.extensions.length).toBeGreaterThan(0);
    expect(result.extensions.find((e) => e.extension === '.bpmn')?.enabled).toBe(true);
  });

  it('merges missing default extensions into a saved config', () => {
    const partial: Partial<IPreviewSettings> = {
      schemaVersion: 1,
      extensions: [{ displayName: 'BPMN', enabled: true, extension: '.bpmn', mode: 'modeler', renderer: 'bpmn-js' }]
    };
    const result = normalizeSettings(partial, '');
    // New extensions should be appended with disabled defaults
    const ifc = result.extensions.find((e) => e.extension === '.ifc');
    expect(ifc).toBeDefined();
    expect(ifc?.enabled).toBe(false);
  });

  it('preserves user-set enabled flag on optional extensions', () => {
    const partial: Partial<IPreviewSettings> = {
      schemaVersion: 1,
      extensions: [
        { displayName: 'BPMN', enabled: true, extension: '.bpmn', mode: 'modeler', renderer: 'bpmn-js' },
        { displayName: 'diagrams.net', enabled: true, extension: '.drawio', mode: 'modeler', renderer: 'diagrams-net-embed' }
      ]
    };
    const result = normalizeSettings(partial, '');
    expect(result.extensions.find((e) => e.extension === '.drawio')?.enabled).toBe(true);
  });

  it('uses fallbackAppBaseUrl when appBaseUrl is missing', () => {
    const result = normalizeSettings({}, 'https://fallback.example.com');
    expect(result.appBaseUrl).toBe('https://fallback.example.com');
  });

  it('strips trailing slashes from appBaseUrl', () => {
    const result = normalizeSettings({ appBaseUrl: 'https://example.com///' }, '');
    expect(result.appBaseUrl).toBe('https://example.com');
  });

  it('coerces unknown renderer to bpmn-js', () => {
    const partial: Partial<IPreviewSettings> = {
      extensions: [
        {
          displayName: 'X',
          enabled: true,
          extension: '.xyz',
          mode: 'viewer',
          renderer: 'unknown-renderer' as unknown as 'bpmn-js'
        }
      ]
    };
    const result = normalizeSettings(partial, '');
    expect(result.extensions.find((e) => e.extension === '.xyz')?.renderer).toBe('bpmn-js');
  });

  it('correctly normalizes known renderer values', () => {
    const renderers = ['diagrams-net-embed', 'mermaid-js', 'web-ifc', 'occt-step', 'coming-soon'] as const;
    renderers.forEach((renderer) => {
      const partial: Partial<IPreviewSettings> = {
        extensions: [{ displayName: 'X', enabled: false, extension: '.x', mode: 'viewer', renderer }]
      };
      const result = normalizeSettings(partial, '');
      expect(result.extensions.find((e) => e.extension === '.x')?.renderer).toBe(renderer);
    });
  });
});

// ── findExtensionSettings ─────────────────────────────────────────────────────

describe('findExtensionSettings', () => {
  const settings = createDefaultPreviewSettings();

  it('returns the extension for an enabled file', () => {
    const result = findExtensionSettings(settings, 'process.bpmn');
    expect(result?.extension).toBe('.bpmn');
  });

  it('returns undefined for a disabled extension', () => {
    expect(findExtensionSettings(settings, 'diagram.drawio')).toBeUndefined();
    expect(findExtensionSettings(settings, 'model.ifc')).toBeUndefined();
    expect(findExtensionSettings(settings, 'diagram.mmd')).toBeUndefined();
  });

  it('returns undefined for coming-soon even if enabled', () => {
    const withEnabled: IPreviewSettings = {
      ...settings,
      extensions: settings.extensions.map((e) => (e.renderer === 'coming-soon' ? { ...e, enabled: true } : e))
    };
    const jt = findExtensionSettings(withEnabled, 'model.jt');
    expect(jt).toBeUndefined();
  });

  it('is case-insensitive for filenames', () => {
    const result = findExtensionSettings(settings, 'PROCESS.BPMN');
    expect(result?.extension).toBe('.bpmn');
  });

  it('returns undefined for an unrecognised extension', () => {
    expect(findExtensionSettings(settings, 'file.txt')).toBeUndefined();
  });

  it('matches when new renderers are explicitly enabled', () => {
    const enabledSettings: IPreviewSettings = {
      ...settings,
      extensions: settings.extensions.map((e) =>
        e.extension === '.mmd' || e.extension === '.ifc' || e.extension === '.step' ? { ...e, enabled: true } : e
      )
    };
    expect(findExtensionSettings(enabledSettings, 'diagram.mmd')?.renderer).toBe('mermaid-js');
    expect(findExtensionSettings(enabledSettings, 'building.ifc')?.renderer).toBe('web-ifc');
    expect(findExtensionSettings(enabledSettings, 'part.step')?.renderer).toBe('occt-step');
  });
});

// ── normalizeBaseUrl ──────────────────────────────────────────────────────────

describe('normalizeBaseUrl', () => {
  it('strips trailing slashes', () => {
    expect(normalizeBaseUrl('https://example.com/')).toBe('https://example.com');
    expect(normalizeBaseUrl('https://example.com///')).toBe('https://example.com');
  });

  it('leaves URLs without trailing slashes unchanged', () => {
    expect(normalizeBaseUrl('https://example.com')).toBe('https://example.com');
  });

  it('trims whitespace', () => {
    expect(normalizeBaseUrl('  https://example.com  ')).toBe('https://example.com');
  });

  it('handles empty string', () => {
    expect(normalizeBaseUrl('')).toBe('');
  });
});

// ── normalizeExtension ────────────────────────────────────────────────────────

describe('normalizeExtension', () => {
  it('lowercases the extension', () => {
    expect(normalizeExtension('.BPMN')).toBe('.bpmn');
  });

  it('prepends a dot if missing', () => {
    expect(normalizeExtension('bpmn')).toBe('.bpmn');
  });

  it('does not double-prepend dots', () => {
    expect(normalizeExtension('.bpmn')).toBe('.bpmn');
  });

  it('trims whitespace', () => {
    expect(normalizeExtension('  .bpmn  ')).toBe('.bpmn');
  });

  it('returns empty string for blank input', () => {
    expect(normalizeExtension('')).toBe('');
    expect(normalizeExtension('   ')).toBe('');
  });
});
