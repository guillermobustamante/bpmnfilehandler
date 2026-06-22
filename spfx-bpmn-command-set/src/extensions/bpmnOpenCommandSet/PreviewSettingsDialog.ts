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
          <div>
            <h2>File preview settings</h2>
            <p>Tenant defaults should live in one central SharePoint configuration site. This instance is using ${escapeHtml(
              this.scriptContext.configSiteUrl
            )}.</p>
          </div>
          <button class="bpf-admin__close" type="button" aria-label="Close settings" title="Close">&times;</button>
        </div>

        <div class="bpf-admin__body">
          <section class="bpf-admin__section">
            <h3>Configuration scope</h3>
            <dl class="bpf-admin__facts">
              <div><dt>Config list</dt><dd>${escapeHtml(CONFIG_LIST_TITLE)}</dd></div>
              <div><dt>Config site</dt><dd>${escapeHtml(this.scriptContext.configSiteUrl)}</dd></div>
              <div><dt>Current site</dt><dd>${escapeHtml(this.scriptContext.currentSiteUrl)}</dd></div>
            </dl>
          </section>

          <section class="bpf-admin__section">
            <h3>Optional native Microsoft 365 File Handler</h3>
            <label class="bpf-admin__field">
              <span>File Handler endpoint URL</span>
              <input data-field="appBaseUrl" type="url" placeholder="https://your-handler.example.com" value="${escapeAttribute(
                this.draftSettings.appBaseUrl
              )}" />
            </label>
            <label class="bpf-admin__check">
              <input data-field="fileHandlerEnabled" type="checkbox" ${this.draftSettings.fileHandlerEnabled ? 'checked' : ''} />
              <span>Native File Handler has been registered by an admin</span>
            </label>
            <p class="bpf-admin__hint">The SharePoint command named "File Preview app" is self-contained in this SPFx package and does not use Azure. Native File Handler registration is optional and controls Microsoft's built-in file preview/open flow, which can still route to an external endpoint if an old Entra add-in remains registered.</p>
          </section>

          <section class="bpf-admin__section">
            <h3>How users should open files</h3>
            <p class="bpf-admin__hint">Select one supported file, then use the command bar or item menu action named "File Preview app". Do not use Microsoft's built-in file-name click preview to test this app; SharePoint's native previewer can still show "Can't preview this file" for extensions such as .drawio because that previewer is not controlled by SPFx.</p>
          </section>

          <section class="bpf-admin__section">
            <h3>External renderer disclosure</h3>
            <p class="bpf-admin__hint">The .drawio renderer is disabled by default. If enabled, the "File Preview app" command loads the drawing XML from SharePoint into the diagrams.net embedded runtime at https://embed.diagrams.net in the user's browser session. This does not change Microsoft's built-in SharePoint previewer.</p>
          </section>

          <section class="bpf-admin__section">
            <h3>License</h3>
            <div class="bpf-admin__grid">
              <label class="bpf-admin__field">
                <span>Tier</span>
                <select data-field="licenseTier">
                  ${renderTierOption('Free', this.draftSettings.license.tier, 'Free - 20 users or less')}
                  ${renderTierOption('Professional', this.draftSettings.license.tier, 'Professional')}
                  ${renderTierOption('Business', this.draftSettings.license.tier, 'Business')}
                  ${renderTierOption('Enterprise', this.draftSettings.license.tier, 'Enterprise')}
                </select>
              </label>
              <label class="bpf-admin__field">
                <span>Total users</span>
                <input data-field="declaredUserCount" min="0" step="1" type="number" value="${this.draftSettings.license.declaredUserCount}" />
              </label>
            </div>
            <label class="bpf-admin__field">
              <span>License key</span>
              <input data-field="licenseKey" type="password" autocomplete="off" value="${escapeAttribute(this.draftSettings.license.key)}" />
            </label>
          </section>

          <section class="bpf-admin__section">
            <div class="bpf-admin__section-heading">
              <h3>Extensions</h3>
              <button class="bpf-admin__secondary" data-action="add-extension" type="button">Add extension</button>
            </div>
            <div class="bpf-admin__table" role="table" aria-label="File extension settings">
              <div class="bpf-admin__row bpf-admin__row--head" role="row">
                <span>Enabled</span>
                <span>Extension</span>
                <span>Name</span>
                <span>Mode</span>
                <span>Status</span>
                <span></span>
              </div>
              ${this.draftSettings.extensions.map((extension, index) => renderExtensionRow(extension, index)).join('')}
            </div>
          </section>

          <section class="bpf-admin__section">
            <div class="bpf-admin__section-heading">
              <h3>Native File Handler cleanup script</h3>
              <button class="bpf-admin__secondary" data-action="copy-cleanup-script" type="button">Copy cleanup script</button>
            </div>
            <p class="bpf-admin__hint">Use this if Microsoft's built-in preview still opens an old Azure File Handler. It removes matching native File Handler add-ins from Entra app registrations; it does not affect the SPFx SharePoint renderer.</p>
            <textarea class="bpf-admin__script" data-role="cleanup-script" readonly>${escapeHtml(cleanupScript)}</textarea>
          </section>

          <section class="bpf-admin__section">
            <div class="bpf-admin__section-heading">
              <h3>Optional native File Handler registration script</h3>
              <button class="bpf-admin__secondary" data-action="copy-register-script" type="button">Copy registration script</button>
            </div>
            <p class="bpf-admin__hint">Optional only. The SharePoint preview experience does not require PowerShell. Use this only if you intentionally want native Microsoft 365 File Handler integration for OneDrive/File Handler launch flows.</p>
            <textarea class="bpf-admin__script" data-role="register-script" readonly>${escapeHtml(registerScript)}</textarea>
          </section>

          <p class="bpf-admin__message" data-role="message" aria-live="polite"></p>
        </div>

        <div class="bpf-admin__footer">
          <button class="bpf-admin__secondary" data-action="cancel" type="button">Cancel</button>
          <button class="bpf-admin__primary" data-action="save" type="button">Save settings</button>
        </div>
      </div>
    `;

    this.applyDialogChrome();
    this.wireEvents();
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
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
    this.domElement.querySelector('[data-action="copy-cleanup-script"]')?.addEventListener('click', () => {
      this.copyScript('cleanup-script', 'Cleanup script copied.');
    });
    this.domElement.querySelector('[data-action="copy-register-script"]')?.addEventListener('click', () => {
      this.copyScript('register-script', 'Registration script copied.');
    });
    this.domElement.querySelector('[data-action="save"]')?.addEventListener('click', () => {
      this.save().catch((error: unknown) => {
        this.setMessage(error instanceof Error ? error.message : 'Could not save settings.', true);
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

    if (saveButton) {
      saveButton.disabled = true;
    }

    this.setMessage('Saving settings...', false);

    try {
      const savedSettings = await this.settingsService.saveSettings(settings);
      this.onSaved(savedSettings);
      this.setMessage('Settings saved.', false);
      this.close().catch(() => undefined);
    } catch (error) {
      this.setMessage(error instanceof Error ? error.message : 'Could not save settings.', true);
      if (saveButton) {
        saveButton.disabled = false;
      }
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
    if (!script) {
      return;
    }

    navigator.clipboard
      .writeText(script)
      .then(() => this.setMessage(successMessage, false))
      .catch(() => this.setMessage('Could not copy the script. Select the text and copy it manually.', true));
  }

  private setMessage(message: string, isError: boolean): void {
    const messageElement = this.domElement.querySelector('[data-role="message"]') as HTMLElement | null;
    if (!messageElement) {
      return;
    }

    messageElement.textContent = message;
    messageElement.classList.toggle('bpf-admin__message--error', isError);
  }

  private applyDialogChrome(): void {
    const dialogMain = this.domElement.closest('.ms-Dialog-main') as HTMLElement | null;
    if (dialogMain) {
      dialogMain.style.width = 'min(1080px, calc(100vw - 48px))';
      dialogMain.style.maxWidth = 'min(1080px, calc(100vw - 48px))';
      dialogMain.style.maxHeight = 'calc(100vh - 48px)';
      dialogMain.style.borderRadius = '0';
      dialogMain.style.overflow = 'hidden';
    }

    const modalScroll = this.domElement.closest('.ms-Modal-scrollableContent') as HTMLElement | null;
    if (modalScroll) {
      modalScroll.style.maxHeight = 'calc(100vh - 48px)';
      modalScroll.style.overflow = 'hidden';
    }

    const style = document.createElement('style');
    style.textContent = `
      .bpf-admin {
        background: #ffffff;
        color: #242424;
        display: flex;
        flex-direction: column;
        font-family: "Segoe UI", Arial, sans-serif;
        max-height: calc(100vh - 48px);
      }
      .bpf-admin__header,
      .bpf-admin__footer {
        align-items: center;
        display: flex;
        justify-content: space-between;
        padding: 18px 24px;
      }
      .bpf-admin__header {
        border-bottom: 1px solid #e1dfdd;
      }
      .bpf-admin__footer {
        border-top: 1px solid #e1dfdd;
        gap: 12px;
        justify-content: flex-end;
      }
      .bpf-admin h2,
      .bpf-admin h3,
      .bpf-admin p,
      .bpf-admin dl {
        margin: 0;
      }
      .bpf-admin h2 {
        font-size: 22px;
        font-weight: 600;
      }
      .bpf-admin h3 {
        font-size: 16px;
        font-weight: 600;
      }
      .bpf-admin__header p,
      .bpf-admin__hint {
        color: #605e5c;
        font-size: 13px;
      }
      .bpf-admin__header p {
        margin-top: 4px;
      }
      .bpf-admin__body {
        display: grid;
        gap: 18px;
        overflow: auto;
        padding: 20px 24px;
      }
      .bpf-admin__section {
        border: 1px solid #edebe9;
        display: grid;
        gap: 14px;
        padding: 16px;
      }
      .bpf-admin__section-heading {
        align-items: center;
        display: flex;
        justify-content: space-between;
      }
      .bpf-admin__facts {
        display: grid;
        gap: 8px;
      }
      .bpf-admin__facts div {
        display: grid;
        gap: 8px;
        grid-template-columns: 140px minmax(0, 1fr);
      }
      .bpf-admin__facts dt {
        color: #605e5c;
        font-size: 13px;
        font-weight: 600;
      }
      .bpf-admin__facts dd {
        overflow-wrap: anywhere;
      }
      .bpf-admin__grid {
        display: grid;
        gap: 12px;
        grid-template-columns: minmax(0, 1fr) 160px;
      }
      .bpf-admin__field {
        display: grid;
        gap: 6px;
      }
      .bpf-admin__field span,
      .bpf-admin__check span {
        color: #323130;
        font-size: 13px;
        font-weight: 600;
      }
      .bpf-admin__field input,
      .bpf-admin__field select,
      .bpf-admin__row input,
      .bpf-admin__row select,
      .bpf-admin__script {
        border: 1px solid #8a8886;
        box-sizing: border-box;
        font: inherit;
        min-height: 34px;
        padding: 6px 8px;
        width: 100%;
      }
      .bpf-admin__script {
        font-family: Consolas, "Courier New", monospace;
        font-size: 12px;
        height: 220px;
        resize: vertical;
        white-space: pre;
      }
      .bpf-admin__check {
        align-items: center;
        display: flex;
        gap: 8px;
      }
      .bpf-admin__table {
        border: 1px solid #edebe9;
        display: grid;
      }
      .bpf-admin__row {
        align-items: center;
        border-top: 1px solid #edebe9;
        display: grid;
        gap: 10px;
        grid-template-columns: 72px 120px minmax(180px, 1fr) 130px 120px 44px;
        padding: 8px;
      }
      .bpf-admin__row:first-child {
        border-top: 0;
      }
      .bpf-admin__row--head {
        background: #faf9f8;
        color: #605e5c;
        font-size: 12px;
        font-weight: 600;
      }
      .bpf-admin__row--disabled {
        background: #fafafa;
        color: #777777;
      }
      .bpf-admin__status {
        color: #605e5c;
        font-size: 12px;
      }
      .bpf-admin__primary,
      .bpf-admin__secondary,
      .bpf-admin__icon,
      .bpf-admin__close {
        border: 1px solid #8a8886;
        cursor: pointer;
        font: inherit;
      }
      .bpf-admin__primary {
        background: #0078d4;
        border-color: #0078d4;
        color: #ffffff;
        min-height: 36px;
        padding: 0 18px;
      }
      .bpf-admin__secondary {
        background: #ffffff;
        color: #242424;
        min-height: 36px;
        padding: 0 14px;
      }
      .bpf-admin__icon,
      .bpf-admin__close {
        align-items: center;
        background: #ffffff;
        display: inline-flex;
        justify-content: center;
      }
      .bpf-admin__icon {
        height: 32px;
        width: 32px;
      }
      .bpf-admin__close {
        border-radius: 50%;
        font-size: 28px;
        height: 40px;
        line-height: 1;
        padding: 0 0 4px;
        width: 40px;
      }
      .bpf-admin__message {
        color: #605e5c;
        min-height: 20px;
      }
      .bpf-admin__message--error {
        color: #a4262c;
      }
      @media (max-width: 800px) {
        .bpf-admin__grid,
        .bpf-admin__row,
        .bpf-admin__facts div {
          grid-template-columns: 1fr;
        }
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
  const external = extension.renderer === 'diagrams-net-embed';
  return `
    <div class="bpf-admin__row ${comingSoon ? 'bpf-admin__row--disabled' : ''}" data-extension-row="${index}" data-renderer="${
      extension.renderer
    }" role="row">
      <label>
        <input data-row-field="enabled" type="checkbox" ${extension.enabled ? 'checked' : ''} ${comingSoon ? 'disabled' : ''} />
      </label>
      <input data-row-field="extension" aria-label="Extension" value="${escapeAttribute(extension.extension)}" ${
        comingSoon ? 'readonly' : ''
      } />
      <input data-row-field="displayName" aria-label="Display name" value="${escapeAttribute(extension.displayName)}" ${
        comingSoon ? 'readonly' : ''
      } />
      <select data-row-field="mode" aria-label="Mode" ${comingSoon ? 'disabled' : ''}>
        <option value="modeler" ${extension.mode === 'modeler' ? 'selected' : ''}>Modeler</option>
        <option value="viewer" ${extension.mode === 'viewer' ? 'selected' : ''}>Viewer</option>
      </select>
      <span class="bpf-admin__status">${comingSoon ? 'Coming soon' : external ? 'External renderer' : 'Available'}</span>
      <button class="bpf-admin__icon" data-action="remove-extension" data-index="${index}" type="button" aria-label="Remove extension" title="Remove" ${
        comingSoon ? 'disabled' : ''
      }>&times;</button>
    </div>
  `;
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

  return `# Optional Microsoft 365 File Handler registration
# App version: ${APP_VERSION}
# The SPFx SharePoint preview experience does not require this script.
# Run only if you want native Microsoft 365 File Handler integration.
# Required: Azure CLI and Entra application administrator permissions.

$TenantId = "${context.tenantId || '<tenant-id>'}"
$EnabledExtensions = "${enabledExtensions}"
$FileHandlerEndpointUrl = "${endpoint}"

if ($FileHandlerEndpointUrl -eq "<optional-handler-endpoint-url>") {
  throw "Set FileHandlerEndpointUrl to your HTTPS File Handler endpoint before running this script."
}

az login --tenant $TenantId
$app = az ad app create --display-name "Microsoft 365 File Preview Handler" --sign-in-audience AzureADMyOrg | ConvertFrom-Json
$handlerId = [guid]::NewGuid().ToString()
$addIns = @(
  @{
    id = $handlerId
    type = "FileHandler"
    properties = @(
      @{ key = "version"; value = "2" },
      @{ key = "fileTypeDisplayName"; value = "Process and architecture file" },
      @{ key = "fileType"; value = $EnabledExtensions },
      @{ key = "action"; value = "preview" },
      @{ key = "url"; value = "$FileHandlerEndpointUrl/filehandler/preview" }
    )
  }
)
$manifestPath = Join-Path $env:TEMP "m365-file-preview-addins.json"
@{ addIns = $addIns } | ConvertTo-Json -Depth 20 | Set-Content -Path $manifestPath -Encoding UTF8
az ad app update --id $app.appId --set addIns=@$manifestPath

Write-Host "Created File Handler app registration:" $app.appId
Write-Host "Handler ID:" $handlerId
Write-Host "Enabled extensions:" $EnabledExtensions`;
}

function buildCleanupFileHandlerScript(settings: IPreviewSettings, context: IAdminScriptContext): string {
  const enabledExtensions = settings.extensions
    .filter((extension) => extension.renderer !== 'coming-soon')
    .map((extension) => extension.extension)
    .join(',');
  const endpoint = settings.appBaseUrl || 'https://bpmn-file-handler-2f18b433.azurewebsites.net';

  return `# Native Microsoft 365 File Handler cleanup
# App version: ${APP_VERSION}
# The SPFx SharePoint renderer does not use Azure. This script removes stale native
# File Handler add-ins that route Microsoft's built-in preview/open flow to an old endpoint.
# Required: Azure CLI and Entra application administrator permissions.

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
  if (value === 'coming-soon' || value === 'diagrams-net-embed') {
    return value;
  }

  return 'bpmn-js';
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
