import { BaseDialog, type IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  CONFIG_LIST_TITLE,
  DEFAULT_APP_BASE_URL,
  FREE_USER_LIMIT,
  type IFileExtensionSettings,
  type IPreviewSettings,
  type LicenseTier,
  normalizeBaseUrl,
  normalizeExtension,
  normalizeSettings,
  PreviewSettingsService
} from './previewSettings';
import { APP_VERSION } from '../../shared/appConstants';

type SaveCallback = (settings: IPreviewSettings) => void;

export interface IAdminScriptContext {
  componentId: string;
  configSiteUrl: string;
  currentSiteUrl: string;
  tenantHostName: string;
  tenantId: string;
}

// Renderers for which the Modeler/Viewer mode toggle has no effect at runtime
const VIEW_ONLY_RENDERERS = new Set<IFileExtensionSettings['renderer']>(['web-ifc', 'occt-step']);

export class PreviewSettingsDialog extends BaseDialog {
  private draftSettings: IPreviewSettings;

  public constructor(
    private readonly settingsService: PreviewSettingsService,
    settings: IPreviewSettings,
    private readonly scriptContext: IAdminScriptContext,
    private readonly onSaved: SaveCallback
  ) {
    super({ isBlocking: false });
    this.draftSettings = normalizeSettings(settings, DEFAULT_APP_BASE_URL);
  }

  public render(): void {
    const registerScript = buildRegisterFileHandlerScript(this.draftSettings, this.scriptContext);
    const cleanupScript = buildCleanupFileHandlerScript(this.draftSettings, this.scriptContext);
    this.domElement.innerHTML = `
      <div class="bpf-admin">
        <div class="bpf-admin__header">
          <div class="bpf-admin__header-text">
            <h2 class="bpf-admin__title">File Preview Settings</h2>
            <p class="bpf-admin__subtitle">Configure which file types are enabled across your tenant. Settings are stored in <strong>${escapeHtml(this.scriptContext.configSiteUrl)}</strong>.</p>
          </div>
          <button class="bpf-admin__close" type="button" aria-label="Close settings" title="Close">&times;</button>
        </div>

        <div class="bpf-admin__body">

          <section class="bpf-admin__card">
            <h3 class="bpf-admin__card-title">Extensions</h3>
            <p class="bpf-admin__card-desc">Enable the file types you want users to preview. At least one must be enabled to save.</p>
            <div class="bpf-admin__ext-table" role="table" aria-label="File extension settings">
              <div class="bpf-admin__ext-head" role="row">
                <span>On</span>
                <span>Extension</span>
                <span>Display name</span>
                <span>Mode</span>
                <span>Renderer</span>
                <span></span>
              </div>
              ${this.draftSettings.extensions.map((ext, idx) => renderExtensionRow(ext, idx)).join('')}
            </div>
            <button class="bpf-admin__add-btn" data-action="add-extension" type="button">+ Add extension</button>
          </section>

          <section class="bpf-admin__card">
            <h3 class="bpf-admin__card-title">License</h3>
            <div class="bpf-admin__two-col">
              <label class="bpf-admin__field">
                <span class="bpf-admin__label">Tier</span>
                <select class="bpf-admin__select" data-field="licenseTier">
                  ${renderTierOption('Free', this.draftSettings.license.tier, 'Free (up to 20 users)')}
                  ${renderTierOption('Professional', this.draftSettings.license.tier, 'Professional')}
                  ${renderTierOption('Business', this.draftSettings.license.tier, 'Business')}
                  ${renderTierOption('Enterprise', this.draftSettings.license.tier, 'Enterprise')}
                </select>
              </label>
              <label class="bpf-admin__field">
                <span class="bpf-admin__label">Total users</span>
                <input class="bpf-admin__input" data-field="declaredUserCount" min="0" step="1" type="number" value="${this.draftSettings.license.declaredUserCount}" />
              </label>
            </div>
            <label class="bpf-admin__field">
              <span class="bpf-admin__label">License key</span>
              <input class="bpf-admin__input" data-field="licenseKey" type="password" autocomplete="off" value="${escapeAttribute(this.draftSettings.license.key)}" />
            </label>
            <p class="bpf-admin__hint">License keys are stored in the SharePoint configuration list visible to site administrators. This is a validation token, not an external-system credential.</p>
          </section>

          <section class="bpf-admin__card">
            <h3 class="bpf-admin__card-title">Configuration scope</h3>
            <dl class="bpf-admin__facts">
              <div><dt>Config list</dt><dd>${escapeHtml(CONFIG_LIST_TITLE)}</dd></div>
              <div><dt>Config site</dt><dd>${escapeHtml(this.scriptContext.configSiteUrl)}</dd></div>
              <div><dt>Current site</dt><dd>${escapeHtml(this.scriptContext.currentSiteUrl)}</dd></div>
            </dl>
          </section>

          <section class="bpf-admin__card">
            <h3 class="bpf-admin__card-title">Optional native Microsoft 365 File Handler</h3>
            <label class="bpf-admin__field">
              <span class="bpf-admin__label">File Handler endpoint URL</span>
              <input class="bpf-admin__input" data-field="appBaseUrl" type="url" placeholder="https://your-handler.example.com" value="${escapeAttribute(this.draftSettings.appBaseUrl)}" />
            </label>
            <label class="bpf-admin__toggle-row">
              <span class="bpf-admin__toggle">
                <input data-field="fileHandlerEnabled" type="checkbox" ${this.draftSettings.fileHandlerEnabled ? 'checked' : ''} />
                <span class="bpf-admin__toggle-track"></span>
              </span>
              <span class="bpf-admin__toggle-label">Native File Handler has been registered by an admin</span>
            </label>
            <p class="bpf-admin__hint">The SharePoint command (Open BPMN / Open DrawIO) is self-contained in this SPFx package and does not require Azure. Native File Handler registration is optional and controls Microsoft's built-in file preview flow.</p>
          </section>

          <div class="bpf-admin__warning" role="alert">
            <strong>Tenant Administrator action required.</strong> The scripts below require Entra <em>Application Administrator</em> or <em>Privileged Role Administrator</em> permissions. Do not share publicly.
          </div>

          <details class="bpf-admin__card bpf-admin__collapsible">
            <summary class="bpf-admin__collapsible-summary">
              <span class="bpf-admin__card-title">Native File Handler cleanup script</span>
              <button class="bpf-admin__copy-btn" data-action="copy-cleanup-script" type="button">Copy</button>
            </summary>
            <p class="bpf-admin__hint">Use this if Microsoft's built-in preview still opens an old Azure File Handler. Removes matching add-ins from Entra; does not affect the SPFx renderer.</p>
            <textarea class="bpf-admin__script" data-role="cleanup-script" readonly>${escapeHtml(cleanupScript)}</textarea>
          </details>

          <details class="bpf-admin__card bpf-admin__collapsible">
            <summary class="bpf-admin__collapsible-summary">
              <span class="bpf-admin__card-title">Optional native File Handler registration script</span>
              <button class="bpf-admin__copy-btn" data-action="copy-register-script" type="button">Copy</button>
            </summary>
            <p class="bpf-admin__hint">Optional only. Use this if you want native Microsoft 365 File Handler integration for OneDrive / File Handler launch flows. The SharePoint preview experience does not require this script.</p>
            <textarea class="bpf-admin__script" data-role="register-script" readonly>${escapeHtml(registerScript)}</textarea>
          </details>

          <p class="bpf-admin__message" data-role="message" aria-live="polite"></p>
        </div>

        <div class="bpf-admin__footer">
          <span class="bpf-admin__version">v${APP_VERSION}</span>
          <div class="bpf-admin__footer-actions">
            <button class="bpf-admin__btn-secondary" data-action="cancel" type="button">Cancel</button>
            <button class="bpf-admin__btn-primary" data-action="save" type="button">Save settings</button>
          </div>
        </div>
      </div>
    `;

    this.applyDialogChrome();
    this.wireEvents();
  }

  public getConfig(): IDialogConfiguration {
    return { isBlocking: false };
  }

  protected onAfterClose(): void {
    this.domElement.innerHTML = '';
    super.onAfterClose();
  }

  private wireEvents(): void {
    this.domElement.querySelector('.bpf-admin__close')?.addEventListener('click', () => {
      this.close().catch(() => undefined);
    });
    this.domElement.querySelector('[data-action="cancel"]')?.addEventListener('click', () => {
      this.close().catch(() => undefined);
    });
    this.domElement.querySelector('[data-action="add-extension"]')?.addEventListener('click', () => {
      const current = this.readForm();
      current.extensions.push({
        displayName: 'New BPMN-compatible file type',
        enabled: false,
        extension: '.new',
        mode: 'viewer',
        renderer: 'bpmn-js'
      });
      this.draftSettings = current;
      this.render();
    });
    this.domElement.querySelectorAll('[data-action="remove-extension"]').forEach((button) => {
      button.addEventListener('click', () => {
        const index = Number((button as HTMLElement).dataset.index);
        const current = this.readForm();
        current.extensions.splice(index, 1);
        this.draftSettings = current;
        this.render();
      });
    });
    this.domElement.querySelector('[data-action="copy-cleanup-script"]')?.addEventListener('click', (e) => {
      e.stopPropagation(); // prevent details toggle
      this.copyScript('cleanup-script', 'Cleanup script copied.');
    });
    this.domElement.querySelector('[data-action="copy-register-script"]')?.addEventListener('click', (e) => {
      e.stopPropagation();
      this.copyScript('register-script', 'Registration script copied.');
    });
    this.domElement.querySelector('[data-action="save"]')?.addEventListener('click', () => {
      this.save().catch((error: unknown) => {
        this.setMessage(error instanceof Error ? error.message : 'Could not save settings.', true);
      });
    });

    // Warn before enabling diagrams.net (external renderer)
    this.domElement.querySelectorAll('[data-extension-row]').forEach((row) => {
      const rowEl = row as HTMLElement;
      if (rowEl.dataset.renderer !== 'diagrams-net-embed') return;
      const checkbox = rowEl.querySelector('[data-row-field="enabled"]') as HTMLInputElement | null;
      if (!checkbox) return;
      checkbox.addEventListener('change', () => {
        if (!checkbox.checked) return;
        const confirmed = window.confirm(
          'Enabling diagrams.net will send drawing XML from SharePoint to the external\n' +
          'diagrams.net service (https://embed.diagrams.net) in the user\'s browser.\n\n' +
          'Your organization\'s diagram content will be processed by an external service.\n' +
          'Confirm this is acceptable under your data handling policies before enabling.'
        );
        if (!confirmed) checkbox.checked = false;
      });
    });
  }

  private async save(): Promise<void> {
    const saveButton = this.domElement.querySelector('[data-action="save"]') as HTMLButtonElement | null;
    const settings = this.readForm();
    const validationMessage = validateSettings(settings);
    if (validationMessage) {
      this.setMessage(validationMessage, true);
      return;
    }

    if (saveButton) saveButton.disabled = true;
    this.setMessage('Saving settings…', false);

    try {
      const savedSettings = await this.settingsService.saveSettings(settings);
      this.onSaved(savedSettings);
      this.setMessage('Settings saved.', false);
      this.close().catch(() => undefined);
    } catch (error) {
      this.setMessage(error instanceof Error ? error.message : 'Could not save settings.', true);
      if (saveButton) saveButton.disabled = false;
    }
  }

  private readForm(): IPreviewSettings {
    const appBaseUrl = readInputValue(this.domElement, 'appBaseUrl') || DEFAULT_APP_BASE_URL;
    const declaredUserCount = Number(readInputValue(this.domElement, 'declaredUserCount'));
    const tier = readInputValue(this.domElement, 'licenseTier') as LicenseTier;
    const extensions = Array.from(this.domElement.querySelectorAll('[data-extension-row]')).map((row) => {
      const rowElement = row as HTMLElement;
      const renderer = readRenderer(rowElement.dataset.renderer);
      return {
        displayName: readRowInputValue(rowElement, 'displayName') || 'File type',
        enabled: renderer === 'coming-soon' ? false : readRowCheckboxValue(rowElement, 'enabled'),
        extension: normalizeExtension(readRowInputValue(rowElement, 'extension')),
        mode: readRowInputValue(rowElement, 'mode') === 'modeler' ? 'modeler' : 'viewer',
        renderer
      } as IFileExtensionSettings;
    });

    return normalizeSettings(
      {
        appBaseUrl: normalizeBaseUrl(appBaseUrl),
        extensions,
        fileHandlerEnabled: readCheckboxValue(this.domElement, 'fileHandlerEnabled'),
        license: {
          declaredUserCount: Number.isFinite(declaredUserCount) ? declaredUserCount : FREE_USER_LIMIT,
          freeUserLimit: FREE_USER_LIMIT,
          key: readInputValue(this.domElement, 'licenseKey'),
          tier
        },
        schemaVersion: 1
      },
      DEFAULT_APP_BASE_URL
    );
  }

  private copyScript(role: string, successMessage: string): void {
    const script = (this.domElement.querySelector(`[data-role="${role}"]`) as HTMLTextAreaElement | null)?.value || '';
    if (!script) return;
    navigator.clipboard
      .writeText(script)
      .then(() => this.setMessage(successMessage, false))
      .catch(() => this.setMessage('Could not copy the script. Select the text and copy it manually.', true));
  }

  private setMessage(message: string, isError: boolean): void {
    const el = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!el) return;
    el.textContent = message;
    el.classList.toggle('bpf-admin__message--error', isError);
  }

  private applyDialogChrome(): void {
    const dialogMain = this.domElement.closest('.ms-Dialog-main') as HTMLElement | null;
    if (dialogMain) {
      dialogMain.style.width = 'min(1080px, calc(100vw - 48px))';
      dialogMain.style.maxWidth = 'min(1080px, calc(100vw - 48px))';
      dialogMain.style.maxHeight = 'calc(100vh - 48px)';
      dialogMain.style.borderRadius = '8px';
      dialogMain.style.overflow = 'hidden';
      dialogMain.style.boxShadow = '0 8px 40px rgba(0,0,0,.2)';
    }

    const modalScroll = this.domElement.closest('.ms-Modal-scrollableContent') as HTMLElement | null;
    if (modalScroll) {
      modalScroll.style.maxHeight = 'calc(100vh - 48px)';
      modalScroll.style.overflow = 'hidden';
    }

    const style = document.createElement('style');
    style.textContent = `
      :root {
        --bpf-blue: #0078d4;
        --bpf-blue-hover: #106ebe;
        --bpf-blue-light: #deecf9;
        --bpf-text: #242424;
        --bpf-text-secondary: #605e5c;
        --bpf-text-disabled: #a19f9d;
        --bpf-surface: #ffffff;
        --bpf-bg: #f5f5f5;
        --bpf-border: #e1dfdd;
        --bpf-border-strong: #c8c6c4;
        --bpf-red: #a4262c;
        --bpf-green: #107c10;
        --bpf-warn-bg: #fff4ce;
        --bpf-warn-border: #f2c811;
        --bpf-radius: 6px;
      }
      .bpf-admin {
        background: var(--bpf-bg);
        color: var(--bpf-text);
        display: flex;
        flex-direction: column;
        font-family: "Segoe UI Variable", "Segoe UI", system-ui, Arial, sans-serif;
        font-size: 14px;
        max-height: calc(100vh - 48px);
      }
      /* ── Header ── */
      .bpf-admin__header {
        align-items: flex-start;
        background: var(--bpf-surface);
        border-bottom: 1px solid var(--bpf-border);
        display: flex;
        gap: 16px;
        justify-content: space-between;
        padding: 20px 24px 16px;
      }
      .bpf-admin__title {
        font-size: 20px;
        font-weight: 600;
        line-height: 1.2;
        margin: 0 0 4px;
      }
      .bpf-admin__subtitle {
        color: var(--bpf-text-secondary);
        font-size: 13px;
        margin: 0;
      }
      .bpf-admin__close {
        align-items: center;
        background: none;
        border: none;
        border-radius: 50%;
        color: var(--bpf-text);
        cursor: pointer;
        display: inline-flex;
        flex: 0 0 auto;
        font-size: 22px;
        height: 36px;
        justify-content: center;
        line-height: 1;
        margin-top: -4px;
        padding: 0;
        width: 36px;
      }
      .bpf-admin__close:hover { background: var(--bpf-border); }
      /* ── Scrollable body ── */
      .bpf-admin__body {
        display: flex;
        flex-direction: column;
        gap: 16px;
        overflow: auto;
        padding: 20px 24px;
      }
      /* ── Cards ── */
      .bpf-admin__card {
        background: var(--bpf-surface);
        border-radius: var(--bpf-radius);
        box-shadow: 0 1px 4px rgba(0,0,0,.07);
        display: flex;
        flex-direction: column;
        gap: 14px;
        padding: 20px;
      }
      .bpf-admin__card-title {
        font-size: 15px;
        font-weight: 600;
        margin: 0;
      }
      .bpf-admin__card-desc {
        color: var(--bpf-text-secondary);
        font-size: 13px;
        margin: -8px 0 0;
      }
      /* ── Collapsible sections (PS scripts) ── */
      .bpf-admin__collapsible { gap: 0; padding: 0; }
      .bpf-admin__collapsible-summary {
        align-items: center;
        cursor: pointer;
        display: flex;
        gap: 12px;
        justify-content: space-between;
        list-style: none;
        padding: 16px 20px;
        user-select: none;
      }
      .bpf-admin__collapsible-summary::-webkit-details-marker { display: none; }
      .bpf-admin__collapsible-summary::before {
        border: 5px solid transparent;
        border-left: 8px solid var(--bpf-text-secondary);
        content: '';
        flex: 0 0 auto;
        margin-right: 2px;
        transition: transform .15s;
      }
      .bpf-admin__collapsible[open] > .bpf-admin__collapsible-summary::before {
        transform: rotate(90deg);
      }
      .bpf-admin__collapsible[open] > .bpf-admin__card-title,
      .bpf-admin__collapsible[open] > p,
      .bpf-admin__collapsible[open] > textarea {
        display: block;
      }
      .bpf-admin__collapsible > p,
      .bpf-admin__collapsible > textarea {
        margin: 0 20px;
      }
      .bpf-admin__collapsible > p { margin-bottom: 8px; }
      .bpf-admin__collapsible > textarea { margin-bottom: 20px; }
      /* ── Extension table ── */
      .bpf-admin__ext-table {
        border: 1px solid var(--bpf-border);
        border-radius: 4px;
        display: grid;
        overflow: hidden;
      }
      .bpf-admin__ext-head {
        background: var(--bpf-bg);
        border-bottom: 1px solid var(--bpf-border);
        color: var(--bpf-text-secondary);
        display: grid;
        font-size: 12px;
        font-weight: 600;
        gap: 10px;
        grid-template-columns: 52px 110px minmax(160px,1fr) 130px 120px 36px;
        letter-spacing: .03em;
        padding: 8px 12px;
        text-transform: uppercase;
      }
      .bpf-admin__ext-row {
        border-top: 1px solid var(--bpf-border);
        display: grid;
        gap: 10px;
        grid-template-columns: 52px 110px minmax(160px,1fr) 130px 120px 36px;
        padding: 8px 12px;
        align-items: center;
      }
      .bpf-admin__ext-row:first-of-type { border-top: 0; }
      .bpf-admin__ext-row--disabled { opacity: .55; }
      /* Toggle switch */
      .bpf-admin__toggle {
        display: inline-flex;
        height: 20px;
        position: relative;
        width: 40px;
      }
      .bpf-admin__toggle input[type="checkbox"] {
        height: 0;
        opacity: 0;
        position: absolute;
        width: 0;
      }
      .bpf-admin__toggle-track {
        background: var(--bpf-border-strong);
        border-radius: 10px;
        bottom: 0;
        cursor: pointer;
        left: 0;
        position: absolute;
        right: 0;
        top: 0;
        transition: background .15s;
      }
      .bpf-admin__toggle-track::after {
        background: #fff;
        border-radius: 50%;
        bottom: 2px;
        box-shadow: 0 1px 3px rgba(0,0,0,.3);
        content: '';
        height: 16px;
        left: 2px;
        position: absolute;
        transition: left .15s;
        width: 16px;
      }
      .bpf-admin__toggle input:checked + .bpf-admin__toggle-track { background: var(--bpf-blue); }
      .bpf-admin__toggle input:checked + .bpf-admin__toggle-track::after { left: 22px; }
      .bpf-admin__toggle input:disabled + .bpf-admin__toggle-track { cursor: not-allowed; opacity: .5; }
      .bpf-admin__toggle-row {
        align-items: center;
        display: flex;
        gap: 10px;
      }
      .bpf-admin__toggle-label { color: var(--bpf-text); font-size: 14px; }
      /* Renderer badge */
      .bpf-admin__renderer-badge {
        border-radius: 3px;
        display: inline-block;
        font-size: 11px;
        font-weight: 600;
        padding: 2px 7px;
        white-space: nowrap;
      }
      /* Form controls */
      .bpf-admin__input,
      .bpf-admin__select {
        background: var(--bpf-surface);
        border: 1px solid var(--bpf-border-strong);
        border-radius: 4px;
        box-sizing: border-box;
        color: var(--bpf-text);
        font: inherit;
        min-height: 34px;
        padding: 5px 10px;
        width: 100%;
      }
      .bpf-admin__input:focus,
      .bpf-admin__select:focus {
        border-color: var(--bpf-blue);
        box-shadow: 0 0 0 1px var(--bpf-blue);
        outline: none;
      }
      .bpf-admin__input:disabled,
      .bpf-admin__select:disabled {
        background: var(--bpf-bg);
        color: var(--bpf-text-disabled);
        cursor: not-allowed;
      }
      /* Inline row inputs (ext table) */
      .bpf-admin__ext-row input[type="text"],
      .bpf-admin__ext-row input:not([type]),
      .bpf-admin__ext-row select {
        background: var(--bpf-surface);
        border: 1px solid var(--bpf-border-strong);
        border-radius: 4px;
        box-sizing: border-box;
        color: var(--bpf-text);
        font: inherit;
        font-size: 13px;
        min-height: 30px;
        padding: 4px 8px;
        width: 100%;
      }
      .bpf-admin__ext-row input:read-only,
      .bpf-admin__ext-row select:disabled {
        background: var(--bpf-bg);
        color: var(--bpf-text-disabled);
        cursor: default;
      }
      .bpf-admin__ext-row select:disabled { cursor: not-allowed; }
      .bpf-admin__mode-hint {
        color: var(--bpf-text-disabled);
        font-size: 11px;
        font-style: italic;
        margin-top: 2px;
      }
      /* Script textarea */
      .bpf-admin__script {
        background: #1e1e1e;
        border: 1px solid var(--bpf-border);
        border-radius: 4px;
        box-sizing: border-box;
        color: #d4d4d4;
        display: block;
        font-family: Consolas, "Cascadia Code", "Courier New", monospace;
        font-size: 12px;
        height: 200px;
        resize: vertical;
        white-space: pre;
        width: 100%;
      }
      /* Layout helpers */
      .bpf-admin__two-col {
        display: grid;
        gap: 12px;
        grid-template-columns: minmax(0,1fr) 160px;
      }
      .bpf-admin__field { display: flex; flex-direction: column; gap: 5px; }
      .bpf-admin__label { color: var(--bpf-text); font-size: 13px; font-weight: 600; }
      .bpf-admin__hint { color: var(--bpf-text-secondary); font-size: 13px; margin: 0; }
      /* DL facts */
      .bpf-admin__facts { display: grid; gap: 8px; margin: 0; }
      .bpf-admin__facts div { display: grid; gap: 8px; grid-template-columns: 140px minmax(0,1fr); }
      .bpf-admin__facts dt { color: var(--bpf-text-secondary); font-size: 13px; font-weight: 600; }
      .bpf-admin__facts dd { overflow-wrap: anywhere; }
      /* Warning banner */
      .bpf-admin__warning {
        background: var(--bpf-warn-bg);
        border: 1px solid var(--bpf-warn-border);
        border-radius: var(--bpf-radius);
        color: var(--bpf-text);
        font-size: 13px;
        padding: 12px 16px;
      }
      /* Add extension */
      .bpf-admin__add-btn {
        align-self: flex-start;
        background: none;
        border: 1px dashed var(--bpf-border-strong);
        border-radius: 4px;
        color: var(--bpf-blue);
        cursor: pointer;
        font: inherit;
        font-size: 13px;
        padding: 6px 14px;
      }
      .bpf-admin__add-btn:hover { background: var(--bpf-blue-light); border-color: var(--bpf-blue); }
      /* Remove extension */
      .bpf-admin__remove-btn {
        align-items: center;
        background: none;
        border: 1px solid transparent;
        border-radius: 4px;
        color: var(--bpf-text-secondary);
        cursor: pointer;
        display: inline-flex;
        font: inherit;
        font-size: 18px;
        height: 30px;
        justify-content: center;
        padding: 0;
        width: 30px;
      }
      .bpf-admin__remove-btn:hover:not(:disabled) { border-color: var(--bpf-border-strong); color: var(--bpf-red); }
      .bpf-admin__remove-btn:disabled { color: var(--bpf-text-disabled); cursor: not-allowed; }
      /* Copy button in collapsible header */
      .bpf-admin__copy-btn {
        background: none;
        border: 1px solid var(--bpf-border-strong);
        border-radius: 4px;
        color: var(--bpf-text);
        cursor: pointer;
        flex: 0 0 auto;
        font: inherit;
        font-size: 12px;
        padding: 4px 12px;
      }
      .bpf-admin__copy-btn:hover { background: var(--bpf-bg); }
      /* Footer */
      .bpf-admin__footer {
        align-items: center;
        background: var(--bpf-surface);
        border-top: 1px solid var(--bpf-border);
        display: flex;
        gap: 12px;
        justify-content: space-between;
        padding: 14px 24px;
      }
      .bpf-admin__footer-actions { display: flex; gap: 10px; }
      .bpf-admin__version { color: var(--bpf-text-disabled); font-size: 12px; }
      .bpf-admin__btn-primary,
      .bpf-admin__btn-secondary {
        border-radius: 4px;
        cursor: pointer;
        font: inherit;
        font-size: 14px;
        min-height: 36px;
        padding: 0 20px;
      }
      .bpf-admin__btn-primary {
        background: var(--bpf-blue);
        border: 1px solid var(--bpf-blue);
        color: #fff;
      }
      .bpf-admin__btn-primary:hover:not(:disabled) { background: var(--bpf-blue-hover); border-color: var(--bpf-blue-hover); }
      .bpf-admin__btn-primary:disabled { background: var(--bpf-border-strong); border-color: var(--bpf-border-strong); cursor: not-allowed; }
      .bpf-admin__btn-secondary {
        background: var(--bpf-surface);
        border: 1px solid var(--bpf-border-strong);
        color: var(--bpf-text);
      }
      .bpf-admin__btn-secondary:hover { background: var(--bpf-bg); }
      /* Status message */
      .bpf-admin__message { color: var(--bpf-text-secondary); font-size: 13px; min-height: 18px; margin: 0; }
      .bpf-admin__message--error { color: var(--bpf-red); }
      @media (max-width: 800px) {
        .bpf-admin__two-col,
        .bpf-admin__facts div { grid-template-columns: 1fr; }
        .bpf-admin__ext-head,
        .bpf-admin__ext-row { grid-template-columns: 44px minmax(80px, 1fr) minmax(100px, 1fr) 100px 90px 30px; }
      }
    `;
    this.domElement.prepend(style);
  }
}

function renderTierOption(value: LicenseTier, selected: LicenseTier, label: string): string {
  return `<option value="${value}" ${value === selected ? 'selected' : ''}>${escapeHtml(label)}</option>`;
}

function renderExtensionRow(extension: IFileExtensionSettings, index: number): string {
  const comingSoon = extension.renderer === 'coming-soon';
  const viewOnly = VIEW_ONLY_RENDERERS.has(extension.renderer);
  const modeDisabled = comingSoon || viewOnly;
  const modeTitle = viewOnly ? 'This renderer is view-only — mode has no effect' : '';

  const rendererBadge = getRendererBadge(extension.renderer);

  return `
    <div class="bpf-admin__ext-row ${comingSoon ? 'bpf-admin__ext-row--disabled' : ''}"
         data-extension-row="${index}" data-renderer="${extension.renderer}" role="row">
      <label class="bpf-admin__toggle" title="${comingSoon ? 'Not yet available' : ''}">
        <input data-row-field="enabled" type="checkbox" ${extension.enabled ? 'checked' : ''} ${comingSoon ? 'disabled' : ''} />
        <span class="bpf-admin__toggle-track"></span>
      </label>
      <input data-row-field="extension" aria-label="Extension" type="text" value="${escapeAttribute(extension.extension)}" ${comingSoon ? 'readonly' : ''} />
      <input data-row-field="displayName" aria-label="Display name" type="text" value="${escapeAttribute(extension.displayName)}" ${comingSoon ? 'readonly' : ''} />
      <div>
        <select data-row-field="mode" aria-label="Mode" ${modeDisabled ? 'disabled' : ''} title="${modeTitle}">
          <option value="modeler" ${extension.mode === 'modeler' ? 'selected' : ''}>Modeler</option>
          <option value="viewer" ${extension.mode === 'viewer' ? 'selected' : ''}>Viewer</option>
        </select>
        ${viewOnly ? `<div class="bpf-admin__mode-hint">View-only renderer</div>` : ''}
      </div>
      <span class="bpf-admin__renderer-badge" style="background:${rendererBadge.bg};color:${rendererBadge.color}">${escapeHtml(rendererBadge.label)}</span>
      <button class="bpf-admin__remove-btn" data-action="remove-extension" data-index="${index}" type="button"
              aria-label="Remove extension" title="Remove" ${comingSoon ? 'disabled' : ''}>&times;</button>
    </div>
  `;
}

function getRendererBadge(renderer: IFileExtensionSettings['renderer']): { label: string; color: string; bg: string } {
  const map: Record<IFileExtensionSettings['renderer'], { label: string; color: string; bg: string }> = {
    'bpmn-js':           { label: 'BPMN-JS',    color: '#107c10', bg: '#dff6dd' },
    'diagrams-net-embed':{ label: 'Draw.io',     color: '#c55a11', bg: '#fce4d6' },
    'mermaid-js':        { label: 'Mermaid',     color: '#5c2d91', bg: '#edebf9' },
    'web-ifc':           { label: 'web-ifc',     color: '#004578', bg: '#cce4f7' },
    'occt-step':         { label: 'OCCT',        color: '#003f5c', bg: '#cde8f7' },
    'coming-soon':       { label: 'Coming soon', color: '#a19f9d', bg: '#f3f2f1' }
  };
  return map[renderer] ?? { label: renderer, color: '#605e5c', bg: '#f3f2f1' };
}

function validateSettings(settings: IPreviewSettings): string {
  if (settings.fileHandlerEnabled && !settings.appBaseUrl) {
    return 'Enter the optional File Handler endpoint URL or turn off native File Handler registration.';
  }

  if (settings.appBaseUrl) {
    try {
      const parsed = new URL(settings.appBaseUrl);
      if (parsed.protocol !== 'https:') {
        return 'The File Handler endpoint URL must use HTTPS.';
      }
    } catch {
      return 'Enter a valid File Handler endpoint URL.';
    }
  }

  const extensions = new Set<string>();
  for (const extension of settings.extensions) {
    if (!extension.extension || extension.extension === '.') {
      return 'Every extension row must have a valid file extension.';
    }
    if (extensions.has(extension.extension)) {
      return `Extension ${extension.extension} is listed more than once.`;
    }
    extensions.add(extension.extension);
  }

  if (!settings.extensions.some((extension) => extension.enabled && extension.renderer !== 'coming-soon')) {
    return 'Enable at least one available file extension.';
  }

  if (settings.license.tier === 'Free' && settings.license.declaredUserCount > settings.license.freeUserLimit) {
    return `The free tier is limited to ${settings.license.freeUserLimit} total users.`;
  }

  if (settings.license.tier !== 'Free' && settings.license.key.trim().length === 0) {
    return 'Enter a license key for paid tiers.';
  }

  return '';
}

function buildRegisterFileHandlerScript(settings: IPreviewSettings, context: IAdminScriptContext): string {
  const enabledExtensions = settings.extensions
    .filter((extension) => extension.enabled && extension.renderer !== 'coming-soon')
    .map((extension) => extension.extension)
    .join(',');
  const endpoint = settings.appBaseUrl || '<optional-handler-endpoint-url>';
  const iconBaseUrl = `${context.configSiteUrl}/SiteAssets/M365FilePreviewIcons`;

  return `# Optional Microsoft 365 File Handler registration
# App version: ${APP_VERSION}
# The SPFx SharePoint preview experience does not require this script.
# Run only if you want native Microsoft 365 File Handler integration.
# Required Azure roles: Application Administrator OR Privileged Role Administrator
# Handle this script with appropriate security controls. Do not share publicly.

$TenantId = "${context.tenantId || '<tenant-id>'}"
$EnabledExtensions = "${enabledExtensions}"
$FileHandlerEndpointUrl = "${endpoint}"
$FileHandlerIconBaseUrl = "${iconBaseUrl}"

if ($FileHandlerEndpointUrl -eq "<optional-handler-endpoint-url>") {
  throw "Set FileHandlerEndpointUrl to your HTTPS File Handler endpoint before running this script."
}

az login --tenant $TenantId
$app = az ad app create --display-name "Microsoft 365 File Preview Handler" --sign-in-audience AzureADMyOrg | ConvertFrom-Json

function Get-ExtensionSpec {
  param([string]$Extension)
  switch ($Extension) {
    ".bpmn" {
      return @{
        assetPrefix = "bpmn"
        actionMenuDisplayName = "Open BPMN"
        fileTypeDisplayName = "BPMN process diagram"
        openLabel = "Open BPMN"
        openMode = "modeler"
        previewMode = "viewer"
      }
    }
    ".drawio" {
      return @{
        assetPrefix = "drawio"
        actionMenuDisplayName = "Open DrawIO"
        fileTypeDisplayName = "DrawIO diagram"
        openLabel = "Open DrawIO"
        openMode = "modeler"
        previewMode = "viewer"
      }
    }
    default {
      throw "Native File Handler registration is only available for .bpmn and .drawio. Unsupported: $Extension"
    }
  }
}

function New-IconJson {
  param([string]$AssetPrefix, [string]$IconType)
  @{
    svg = "$FileHandlerIconBaseUrl/$AssetPrefix-$IconType.svg"
    png1x = "$FileHandlerIconBaseUrl/$AssetPrefix-$IconType-32.png"
    "png1.5x" = "$FileHandlerIconBaseUrl/$AssetPrefix-$IconType-48.png"
    png2x = "$FileHandlerIconBaseUrl/$AssetPrefix-$IconType-64.png"
  } | ConvertTo-Json -Compress
}

function New-FileHandler {
  param([string]$Extension)
  $spec = Get-ExtensionSpec -Extension $Extension
  $encodedExtension = [uri]::EscapeDataString($Extension)
  $actions = @(
    @{
      type = "preview"
      url = "$FileHandlerEndpointUrl/filehandler/preview?extension=$encodedExtension&mode=$($spec.previewMode)"
      availableOn = @{ file = @{ extensions = @($Extension) }; web = @{} }
    },
    @{
      type = "open"
      url = "$FileHandlerEndpointUrl/filehandler/open?extension=$encodedExtension&mode=$($spec.openMode)"
      displayName = $spec.openLabel
      shortDisplayName = $spec.openLabel
      availableOn = @{ file = @{ extensions = @($Extension) }; web = @{} }
    }
  ) | ConvertTo-Json -Depth 20 -Compress

  return @{
    id = [guid]::NewGuid().ToString()
    type = "FileHandler"
    properties = @(
      @{ key = "version"; value = "2" },
      @{ key = "fileTypeDisplayName"; value = $spec.fileTypeDisplayName },
      @{ key = "actionMenuDisplayName"; value = $spec.actionMenuDisplayName },
      @{ key = "fileTypeIcon"; value = (New-IconJson -AssetPrefix $spec.assetPrefix -IconType "file") },
      @{ key = "appIcon"; value = (New-IconJson -AssetPrefix $spec.assetPrefix -IconType "app") },
      @{ key = "actions"; value = $actions }
    )
  }
}

$extensions = $EnabledExtensions.Split(",") | ForEach-Object {
  $extension = $_.Trim().ToLowerInvariant()
  if ($extension -and -not $extension.StartsWith(".")) { ".$extension" } else { $extension }
} | Where-Object { $_ } | Select-Object -Unique

$addIns = @()
foreach ($extension in $extensions) {
  $addIns += New-FileHandler -Extension $extension
}

$manifestPath = Join-Path $env:TEMP "m365-file-preview-addins.json"
@{ addIns = $addIns } | ConvertTo-Json -Depth 20 | Set-Content -Path $manifestPath -Encoding UTF8
az ad app update --id $app.appId --set addIns=@$manifestPath

Write-Host "Created File Handler app registration:" $app.appId
Write-Host "File Handler icon base URL:" $FileHandlerIconBaseUrl
Write-Host "Enabled extensions:" $EnabledExtensions`;
}

function buildCleanupFileHandlerScript(settings: IPreviewSettings, context: IAdminScriptContext): string {
  const enabledExtensions = settings.extensions
    .filter((extension) => extension.renderer !== 'coming-soon')
    .map((extension) => extension.extension)
    .join(',');
  const endpoint = settings.appBaseUrl || '<your-file-handler-endpoint-url>';

  return `# Native Microsoft 365 File Handler cleanup
# App version: ${APP_VERSION}
# The SPFx SharePoint renderer does not use Azure. This script removes stale native
# File Handler add-ins that route Microsoft's built-in preview/open flow to an old endpoint.
# Required Azure roles: Application Administrator OR Privileged Role Administrator
# Handle this script with appropriate security controls. Do not share publicly.

$TenantId = "${context.tenantId || '<tenant-id>'}"
$EndpointContains = "${endpoint}"
$ExtensionCsv = "${enabledExtensions || '.bpmn,.drawio,.jt,.step'}"
$RemoveAllMatchingFileTypes = $false

az login --tenant $TenantId

$extensions = $ExtensionCsv.Split(",") | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ }
$apps = az ad app list --all --query "[].{appId:appId,displayName:displayName,addIns:addIns}" | ConvertFrom-Json
$updated = 0

foreach ($app in $apps) {
  if (-not $app.addIns) {
    continue
  }

  $keptAddIns = @()
  $removedAddIns = @()

  foreach ($addIn in $app.addIns) {
    if ($addIn.type -ne "FileHandler") {
      $keptAddIns += $addIn
      continue
    }

    $properties = @{}
    foreach ($property in ($addIn.properties | Where-Object { $_ })) {
      $properties[$property.key] = [string]$property.value
    }

    $url = $properties["url"]
    $fileType = ($properties["fileType"] + "," + $properties["fileTypeDisplayName"]).ToLowerInvariant()
    $matchesEndpoint = $EndpointContains -and $url -and $url.ToLowerInvariant().Contains($EndpointContains.ToLowerInvariant())
    $matchesOldAzureBpmnEndpoint = $url -and $url.ToLowerInvariant().Contains("bpmn-file-handler") -and $url.ToLowerInvariant().Contains("azurewebsites.net")
    $matchesExtension = $false

    foreach ($extension in $extensions) {
      if ($fileType.Contains($extension)) {
        $matchesExtension = $true
      }
    }

    if (($matchesEndpoint -or $matchesOldAzureBpmnEndpoint) -or ($RemoveAllMatchingFileTypes -and $matchesExtension)) {
      $removedAddIns += $addIn
    } else {
      $keptAddIns += $addIn
    }
  }

  if ($removedAddIns.Count -eq 0) {
    continue
  }

  $manifestPath = Join-Path $env:TEMP ("m365-file-preview-cleanup-" + $app.appId + ".json")
  @{ addIns = $keptAddIns } | ConvertTo-Json -Depth 30 | Set-Content -Path $manifestPath -Encoding UTF8
  az ad app update --id $app.appId --set addIns=@$manifestPath
  $updated++
  Write-Host "Removed" $removedAddIns.Count "File Handler add-in(s) from" $app.displayName "(" $app.appId ")"
}

Write-Host "Updated app registrations:" $updated
Write-Host "Microsoft 365 File Handler changes can still take time to expire from SharePoint/OneDrive caches."`;
}

function readRenderer(value: unknown): IFileExtensionSettings['renderer'] {
  switch (value) {
    case 'bpmn-js':
    case 'coming-soon':
    case 'diagrams-net-embed':
    case 'mermaid-js':
    case 'web-ifc':
    case 'occt-step':
      return value;
    default:
      return 'bpmn-js';
  }
}

function readInputValue(root: HTMLElement, fieldName: string): string {
  const input = root.querySelector(`[data-field="${fieldName}"]`) as HTMLInputElement | HTMLSelectElement | null;
  return input?.value.trim() || '';
}

function readCheckboxValue(root: HTMLElement, fieldName: string): boolean {
  const input = root.querySelector(`[data-field="${fieldName}"]`) as HTMLInputElement | null;
  return Boolean(input?.checked);
}

function readRowInputValue(root: HTMLElement, fieldName: string): string {
  const input = root.querySelector(`[data-row-field="${fieldName}"]`) as HTMLInputElement | HTMLSelectElement | null;
  return input?.value.trim() || '';
}

function readRowCheckboxValue(root: HTMLElement, fieldName: string): boolean {
  const input = root.querySelector(`[data-row-field="${fieldName}"]`) as HTMLInputElement | null;
  return Boolean(input?.checked);
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function escapeAttribute(value: string): string {
  return escapeHtml(value);
}
