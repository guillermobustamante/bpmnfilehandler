import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

export type FilePreviewMode = 'modeler' | 'viewer';
export type FilePreviewRenderer = 'bpmn-js' | 'coming-soon' | 'diagrams-net-embed' | 'mermaid-js' | 'web-ifc' | 'occt-step';
export type LicenseTier = 'Free' | 'Professional' | 'Business' | 'Enterprise';

export interface IFileExtensionSettings {
  displayName: string;
  enabled: boolean;
  extension: string;
  mode: FilePreviewMode;
  renderer: FilePreviewRenderer;
}

export interface ILicenseSettings {
  declaredUserCount: number;
  freeUserLimit: number;
  key: string;
  tier: LicenseTier;
}

export interface IPreviewSettings {
  appBaseUrl: string;
  extensions: IFileExtensionSettings[];
  fileHandlerEnabled: boolean;
  license: ILicenseSettings;
  schemaVersion: 1;
}

export interface IPreviewConfigurationStatus {
  configFieldExists: boolean;
  configItemExists: boolean;
  configListExists: boolean;
}

interface IConfigItem {
  BpfConfigJson?: string;
  Id?: number;
}

class SharePointRequestError extends Error {
  public constructor(public readonly status: number, public readonly body: string) {
    super(formatSharePointError(status, body));
  }
}

export const DEFAULT_APP_BASE_URL: string = '';
export const CONFIG_LIST_TITLE: string = 'M365 File Preview Settings';
export const CONFIG_ITEM_TITLE: string = 'Configuration';
export const CONFIG_FIELD_NAME: string = 'BpfConfigJson';
export const FREE_USER_LIMIT: number = 20;

export function createDefaultPreviewSettings(appBaseUrl: string = DEFAULT_APP_BASE_URL): IPreviewSettings {
  return {
    appBaseUrl,
    extensions: [
      {
        displayName: 'BPMN process diagram',
        enabled: true,
        extension: '.bpmn',
        mode: 'modeler',
        renderer: 'bpmn-js'
      },
      {
        displayName: 'JT 3D model',
        enabled: false,
        extension: '.jt',
        mode: 'viewer',
        renderer: 'coming-soon'
      },
      {
        displayName: 'diagrams.net drawing',
        enabled: false,
        extension: '.drawio',
        mode: 'modeler',
        renderer: 'diagrams-net-embed'
      },
      {
        displayName: 'Mermaid diagram',
        enabled: false,
        extension: '.mmd',
        mode: 'viewer',
        renderer: 'mermaid-js'
      },
      {
        displayName: 'Mermaid diagram (.mermaid)',
        enabled: false,
        extension: '.mermaid',
        mode: 'viewer',
        renderer: 'mermaid-js'
      },
      {
        displayName: 'IFC building model',
        enabled: false,
        extension: '.ifc',
        mode: 'viewer',
        renderer: 'web-ifc'
      },
      {
        displayName: 'STEP CAD model (.step)',
        enabled: false,
        extension: '.step',
        mode: 'viewer',
        renderer: 'occt-step'
      },
      {
        displayName: 'STEP CAD model (.stp)',
        enabled: false,
        extension: '.stp',
        mode: 'viewer',
        renderer: 'occt-step'
      }
    ],
    fileHandlerEnabled: false,
    license: {
      declaredUserCount: FREE_USER_LIMIT,
      freeUserLimit: FREE_USER_LIMIT,
      key: '',
      tier: 'Free'
    },
    schemaVersion: 1
  };
}

export function normalizeSettings(settings: Partial<IPreviewSettings> | undefined, fallbackAppBaseUrl: string): IPreviewSettings {
  const defaults = createDefaultPreviewSettings(fallbackAppBaseUrl);
  if (!settings) {
    return defaults;
  }

  const extensions = mergeDefaultExtensions(
    Array.isArray(settings.extensions)
    ? settings.extensions
        .map(normalizeExtensionSettings)
        .filter((extension): extension is IFileExtensionSettings => Boolean(extension))
      : defaults.extensions,
    defaults.extensions
  );

  return {
    appBaseUrl: normalizeBaseUrl(settings.appBaseUrl || fallbackAppBaseUrl || defaults.appBaseUrl),
    extensions: extensions.length > 0 ? extensions : defaults.extensions,
    fileHandlerEnabled: Boolean(settings.fileHandlerEnabled),
    license: {
      declaredUserCount: toSafeNumber(settings.license?.declaredUserCount, FREE_USER_LIMIT),
      freeUserLimit: FREE_USER_LIMIT,
      key: settings.license?.key || '',
      tier: normalizeLicenseTier(settings.license?.tier)
    },
    schemaVersion: 1
  };
}

export function findExtensionSettings(settings: IPreviewSettings, fileName: string): IFileExtensionSettings | undefined {
  const normalizedFileName = fileName.toLowerCase();
  return settings.extensions.find(
    (extension) => extension.enabled && extension.renderer !== 'coming-soon' && normalizedFileName.endsWith(extension.extension)
  );
}

export function normalizeExtension(value: string): string {
  const trimmed = value.trim().toLowerCase();
  if (!trimmed) {
    return '';
  }

  return trimmed.startsWith('.') ? trimmed : `.${trimmed}`;
}

export function normalizeBaseUrl(value: string): string {
  return value.trim().replace(/\/+$/, '');
}

export class PreviewSettingsService {
  public constructor(private readonly spHttpClient: SPHttpClient, private readonly webAbsoluteUrl: string) {}

  public async getConfigurationStatus(): Promise<IPreviewConfigurationStatus> {
    const configListExists = await this.configListExists();
    if (!configListExists) {
      return {
        configFieldExists: false,
        configItemExists: false,
        configListExists: false
      };
    }

    return {
      configFieldExists: await this.configFieldExists(),
      configItemExists: Boolean(await this.tryGetConfigItem(false)),
      configListExists: true
    };
  }

  public async getSettings(fallbackAppBaseUrl: string): Promise<IPreviewSettings> {
    const item = await this.tryGetConfigItem(true);
    if (!item?.BpfConfigJson) {
      return createDefaultPreviewSettings(fallbackAppBaseUrl);
    }

    try {
      return normalizeSettings(JSON.parse(item.BpfConfigJson) as Partial<IPreviewSettings>, fallbackAppBaseUrl);
    } catch {
      return createDefaultPreviewSettings(fallbackAppBaseUrl);
    }
  }

  public async saveSettings(settings: IPreviewSettings): Promise<IPreviewSettings> {
    const normalized = normalizeSettings(settings, settings.appBaseUrl);
    await this.ensureConfigList();
    await this.waitForConfigFieldQueryable();

    const item = await this.tryGetConfigItem(false);
    const body = {
      [CONFIG_FIELD_NAME]: JSON.stringify(normalized),
      Title: CONFIG_ITEM_TITLE
    };

    if (item?.Id) {
      await this.postJson(`${this.configItemsUrl()}(${item.Id})`, body, {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      });
    } else {
      await this.postJson(this.configItemsUrl(), body);
    }

    return normalized;
  }

  public async initializeConfiguration(defaultSettings: IPreviewSettings): Promise<IPreviewConfigurationStatus> {
    await this.ensureConfigList();
    await this.waitForConfigFieldQueryable();

    const item = await this.tryGetConfigItem(false);
    if (!item?.Id) {
      const normalized = normalizeSettings(defaultSettings, defaultSettings.appBaseUrl);
      await this.postJson(this.configItemsUrl(), {
        [CONFIG_FIELD_NAME]: JSON.stringify(normalized),
        Title: CONFIG_ITEM_TITLE
      });
    }

    return this.getConfigurationStatus();
  }

  public async ensureConfigList(): Promise<void> {
    if (!(await this.configListExists())) {
      await this.postJson(`${this.webApiUrl()}/lists`, {
        AllowContentTypes: false,
        BaseTemplate: 100,
        Description: 'Configuration for the Microsoft 365 file preview framework.',
        Title: CONFIG_LIST_TITLE
      });
    }

    if (!(await this.configFieldExists())) {
      await this.createConfigField();
    }
  }

  private async tryGetConfigItem(includeConfigJson: boolean): Promise<IConfigItem | undefined> {
    const selectFields = includeConfigJson ? `Id,Title,${CONFIG_FIELD_NAME}` : 'Id,Title';
    const response = await this.spHttpClient.get(
      `${this.configItemsUrl()}?$select=${selectFields}&$filter=Title eq '${encodeODataString(
        CONFIG_ITEM_TITLE
      )}'&$top=1`,
      SPHttpClient.configurations.v1
    );

    if (response.status === 404) {
      return undefined;
    }

    if (!response.ok && includeConfigJson && isMissingConfigFieldErrorText(await readErrorBody(response))) {
      return this.tryGetConfigItem(false);
    }

    await assertOk(response);
    const payload = (await response.json()) as { value?: IConfigItem[] };
    return payload.value?.[0];
  }

  private async configListExists(): Promise<boolean> {
    const response = await this.spHttpClient.get(this.configListUrl(), SPHttpClient.configurations.v1);
    if (response.status === 404) {
      return false;
    }

    await assertOk(response);
    return true;
  }

  private async configFieldExists(): Promise<boolean> {
    const response = await this.spHttpClient.get(
      `${this.configListUrl()}/fields?$select=InternalName&$filter=InternalName eq '${encodeODataString(CONFIG_FIELD_NAME)}'&$top=1`,
      SPHttpClient.configurations.v1
    );
    if (response.status === 404) {
      return false;
    }

    await assertOk(response);
    const payload = (await response.json()) as { value?: Array<{ InternalName?: string }> };
    return Boolean(payload.value?.some((field) => field.InternalName === CONFIG_FIELD_NAME));
  }

  private async waitForConfigFieldQueryable(): Promise<void> {
    for (let attempt = 0; attempt < 5; attempt++) {
      const response = await this.spHttpClient.get(
        `${this.configItemsUrl()}?$select=Id,${CONFIG_FIELD_NAME}&$top=1`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        return;
      }

      const errorBody = await readErrorBody(response);
      if (!isMissingConfigFieldErrorText(errorBody) || attempt === 4) {
        throw new SharePointRequestError(response.status, errorBody);
      }

      await delay(400 * (attempt + 1));
      if (!(await this.configFieldExists())) {
        await this.ensureConfigList();
      }
    }
  }

  private async createConfigField(): Promise<void> {
    const schemaXml = `<Field Type="Note" DisplayName="${CONFIG_FIELD_NAME}" StaticName="${CONFIG_FIELD_NAME}" Name="${CONFIG_FIELD_NAME}" RichText="FALSE" NumLines="20" />`;
    const createFieldUrl = `${this.configListUrl()}/fields/createfieldasxml`;

    try {
      await this.postJson(createFieldUrl, {
        parameters: {
          Options: 0,
          SchemaXml: schemaXml
        }
      });
      return;
    } catch (error) {
      if (!(error instanceof SharePointRequestError) || !isCreateFieldPayloadError(error.body)) {
        throw error;
      }
    }

    try {
      await this.postVerboseJson(createFieldUrl, {
        parameters: {
          Options: 0,
          SchemaXml: schemaXml
        }
      });
    } catch (error) {
      if (await this.configFieldExists()) {
        return;
      }

      throw new Error(
        `The central configuration list exists, but SharePoint could not create the ${CONFIG_FIELD_NAME} field. ` +
          'Open the tenant root site as a SharePoint admin and try File Preview Admin again. ' +
          getErrorMessage(error)
      );
    }
  }

  private async postJson(url: string, body: unknown, extraHeaders: Record<string, string> = {}): Promise<void> {
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(body),
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        ...extraHeaders
      }
    });

    await assertOk(response);
  }

  private async postVerboseJson(url: string, body: unknown): Promise<void> {
    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(body),
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      }
    });

    await assertOk(response);
  }

  private configItemsUrl(): string {
    return `${this.configListUrl()}/items`;
  }

  private configListUrl(): string {
    return `${this.webApiUrl()}/lists/getbytitle('${encodeODataString(CONFIG_LIST_TITLE)}')`;
  }

  private webApiUrl(): string {
    return `${this.webAbsoluteUrl.replace(/\/+$/, '')}/_api/web`;
  }
}

function normalizeExtensionSettings(value: Partial<IFileExtensionSettings> | undefined): IFileExtensionSettings | undefined {
  const extension = normalizeExtension(value?.extension || '');
  if (!extension) {
    return undefined;
  }

  return {
    displayName: value?.displayName?.trim() || `${extension.toUpperCase()} file`,
    enabled: Boolean(value?.enabled),
    extension,
    mode: value?.mode === 'viewer' ? 'viewer' : 'modeler',
    renderer: normalizeRenderer(value?.renderer)
  };
}

function mergeDefaultExtensions(
  configuredExtensions: IFileExtensionSettings[],
  defaultExtensions: IFileExtensionSettings[]
): IFileExtensionSettings[] {
  // Upgrade saved extensions that have an incorrect renderer:
  // (a) renderer was 'coming-soon' but a real implementation now exists in defaults.
  // (b) renderer is 'bpmn-js' for an extension whose correct default is NOT 'bpmn-js'.
  //     This handles the case where occt-step / web-ifc were added as valid renderers AFTER
  //     the user last saved their settings — normalizeRenderer() previously fell back to
  //     'bpmn-js' for any unrecognised string, so .step/.stp got that fallback value.
  const upgraded = configuredExtensions.map((configured) => {
    const defaultMatch = defaultExtensions.find((d) => d.extension === configured.extension);
    if (!defaultMatch || defaultMatch.renderer === 'coming-soon') {
      return configured;
    }
    const shouldUpgrade =
      configured.renderer === 'coming-soon' ||
      (configured.renderer === 'bpmn-js' && defaultMatch.renderer !== 'bpmn-js');
    return shouldUpgrade ? { ...configured, renderer: defaultMatch.renderer } : configured;
  });

  // Add any default extensions not already present in the configured list.
  for (const defaultExtension of defaultExtensions) {
    if (!upgraded.some((extension) => extension.extension === defaultExtension.extension)) {
      upgraded.push(defaultExtension);
    }
  }

  return upgraded;
}

function normalizeRenderer(value: unknown): FilePreviewRenderer {
  if (
    value === 'coming-soon' ||
    value === 'diagrams-net-embed' ||
    value === 'mermaid-js' ||
    value === 'web-ifc' ||
    value === 'occt-step'
  ) {
    return value;
  }

  return 'bpmn-js';
}

function normalizeLicenseTier(value: unknown): LicenseTier {
  return value === 'Professional' || value === 'Business' || value === 'Enterprise' ? value : 'Free';
}

function toSafeNumber(value: unknown, fallback: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return fallback;
  }

  return Math.max(0, Math.floor(value));
}

function encodeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

async function assertOk(response: SPHttpClientResponse): Promise<void> {
  if (response.ok) {
    return;
  }

  const body = await readErrorBody(response);
  throw new SharePointRequestError(response.status, body);
}

async function readErrorBody(response: SPHttpClientResponse): Promise<string> {
  return response.text().catch(() => '');
}

function isMissingConfigFieldErrorText(body: string): boolean {
  return body.indexOf(CONFIG_FIELD_NAME) !== -1 && body.toLowerCase().indexOf('does not exist') !== -1;
}

function isCreateFieldPayloadError(body: string): boolean {
  const normalized = body.toLowerCase();
  return (
    normalized.indexOf('__metadata') !== -1 ||
    normalized.indexOf('xmlschemafieldcreationinformation') !== -1 ||
    normalized.indexOf('odataexception') !== -1
  );
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error && error.message ? error.message : 'SharePoint returned an unexpected provisioning error.';
}

function formatSharePointError(status: number, body: string): string {
  const parsedMessage = tryReadSharePointErrorMessage(body);
  if (parsedMessage) {
    return parsedMessage;
  }

  return body || `SharePoint returned ${status}.`;
}

function tryReadSharePointErrorMessage(body: string): string {
  if (!body) {
    return '';
  }

  try {
    const parsed = JSON.parse(body) as { error?: { message?: string | { value?: string } } };
    const message = parsed.error?.message;
    if (typeof message === 'string') {
      return message;
    }

    return message?.value || '';
  } catch {
    return '';
  }
}

function delay(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}
