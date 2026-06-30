import { Log } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
  type RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { BpmnViewerDialog } from './BpmnViewerDialog';
import { DrawioViewerDialog } from './DrawioViewerDialog';
import { MermaidViewerDialog } from './MermaidViewerDialog';
import { IfcViewerDialog } from './IfcViewerDialog';
import { StepViewerDialog } from './StepViewerDialog';
import { PreviewSettingsDialog } from './PreviewSettingsDialog';
import {
  createDefaultPreviewSettings,
  DEFAULT_APP_BASE_URL,
  findExtensionSettings,
  type IFileExtensionSettings,
  type IPreviewSettings,
  normalizeBaseUrl,
  PreviewSettingsService
} from './previewSettings';
import { getServerRelativeUrlFromRowValue } from './sharePointFileService';
import { COMMAND_SET_COMPONENT_ID } from '../../shared/appConstants';

export interface IBpmnOpenCommandSetCommandSetProperties {
  configSiteUrl?: string;
  showSettingsCommand?: boolean;
}

const LOG_SOURCE: string = 'BpmnOpenCommandSet';
const OPEN_FILE_COMMAND_ID: string = 'OPEN_BPMN';
const SETTINGS_COMMAND_ID: string = 'BPMN_SETTINGS';

export default class BpmnOpenCommandSetCommandSet extends BaseListViewCommandSet<IBpmnOpenCommandSetCommandSetProperties> {
  private settings: IPreviewSettings = createDefaultPreviewSettings(DEFAULT_APP_BASE_URL);
  private settingsService: PreviewSettingsService | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized configurable file preview command set');

    this.settings = createDefaultPreviewSettings(DEFAULT_APP_BASE_URL);
    this.settingsService = new PreviewSettingsService(this.context.spHttpClient, this.getConfigWebAbsoluteUrl());

    this.hideCommand(OPEN_FILE_COMMAND_ID);
    this.hideCommand(SETTINGS_COMMAND_ID);
    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewStateChanged);

    this.loadSettings().catch((error: unknown) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error('Could not load preview settings.'));
    });
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId === SETTINGS_COMMAND_ID) {
      this.openSettingsDialog();
      return;
    }

    if (event.itemId !== OPEN_FILE_COMMAND_ID) {
      throw new Error('Unknown command');
    }

    this.openSelectedFile().catch((error: unknown) => {
      // console.error makes the real root-cause visible in DevTools (Log.error only goes to SP internal logging)
      console.error('[BpfPreview] openSelectedFile rejected:', error);
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error('Could not open preview dialog.'));
      try {
        Dialog.alert(error instanceof Error ? error.message : 'Could not open the selected file.').catch(() => undefined);
      } catch (alertErr: unknown) {
        // Dialog.alert() throws synchronously when the dialog service host element is gone (domElement undefined).
        // This happens when show() fails and the framework tears down the host before we reach this catch handler.
        console.error('[BpfPreview] Dialog.alert threw (dialog service lost):', alertErr);
      }
    });
  }

  private async openSelectedFile(): Promise<void> {
    const selectedRow = this.context.listView.selectedRows?.[0];
    let selectedFile = selectedRow ? getSelectedFile(selectedRow, this.settings) : undefined;
    if (selectedRow && !selectedFile?.extensionSettings) {
      this.settings = await this.loadSettingsForExecute();
      selectedFile = getSelectedFile(selectedRow, this.settings);
    }

    if (!selectedRow || !selectedFile?.extensionSettings) {
      throw new Error('Select one enabled file type to open.');
    }

    const serverRelativeUrl = getSelectedFileServerRelativeUrl(selectedRow, this.context.pageContext.web.absoluteUrl);
    if (!serverRelativeUrl) {
      throw new Error('Could not read the selected SharePoint file path.');
    }

    const { spHttpClient } = this.context;
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const { renderer } = selectedFile.extensionSettings;

    let dialog;
    if (renderer === 'diagrams-net-embed') {
      dialog = new DrawioViewerDialog(spHttpClient, webUrl, serverRelativeUrl, selectedFile.fileName, selectedFile.extensionSettings);
    } else if (renderer === 'mermaid-js') {
      dialog = new MermaidViewerDialog(spHttpClient, webUrl, serverRelativeUrl, selectedFile.fileName, selectedFile.extensionSettings);
    } else if (renderer === 'web-ifc') {
      dialog = new IfcViewerDialog(spHttpClient, webUrl, serverRelativeUrl, selectedFile.fileName, selectedFile.extensionSettings);
    } else if (renderer === 'occt-step') {
      dialog = new StepViewerDialog(spHttpClient, webUrl, serverRelativeUrl, selectedFile.fileName, selectedFile.extensionSettings);
    } else {
      dialog = new BpmnViewerDialog(spHttpClient, webUrl, serverRelativeUrl, selectedFile.fileName, selectedFile.extensionSettings);
    }
    await dialog.show();
  }

  private onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    this.updateCommandVisibility();
  };

  private async loadSettings(): Promise<void> {
    try {
      this.settings = await this.getSettingsService().getSettings(DEFAULT_APP_BASE_URL);
    } catch (error) {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error('Could not load preview settings.'));
      this.settings = createDefaultPreviewSettings(DEFAULT_APP_BASE_URL);
    }

    this.updateCommandVisibility();
  }

  private async loadSettingsForExecute(): Promise<IPreviewSettings> {
    try {
      return await this.getSettingsService().getSettings(DEFAULT_APP_BASE_URL);
    } catch (error) {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error('Could not refresh preview settings.'));
      return this.settings;
    }
  }

  private openSettingsDialog(): void {
    if (!this.userCanManageSettings()) {
      Dialog.alert('You need Manage Web permissions to update file preview settings.').catch(() => undefined);
      return;
    }

    const dialog = new PreviewSettingsDialog(this.getSettingsService(), this.settings, this.getAdminScriptContext(), (savedSettings) => {
      this.settings = savedSettings;
      this.updateCommandVisibility();
    });
    dialog.show().catch((error: unknown) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error('Could not open settings dialog.'));
    });
  }

  private updateCommandVisibility(): void {
    const openCommand: Command = this.tryGetCommand(OPEN_FILE_COMMAND_ID);
    if (openCommand) {
      const selectedRows = this.context.listView.selectedRows || [];
      const selectedFile = selectedRows.length === 1 ? getSelectedFile(selectedRows[0], this.settings) : undefined;
      const extension = selectedFile?.extensionSettings?.extension ?? '';
      openCommand.visible = selectedRows.length === 1 && Boolean(extension);
      openCommand.title = extension ? getOpenCommandTitle(extension) : 'Open file';
    }

    const settingsCommand: Command = this.tryGetCommand(SETTINGS_COMMAND_ID);
    if (settingsCommand) {
      settingsCommand.visible = Boolean(this.properties.showSettingsCommand) && this.userCanManageSettings();
    }

    this.raiseOnChange();
  }

  private hideCommand(commandId: string): void {
    const command = this.tryGetCommand(commandId);
    if (command) {
      command.visible = false;
    }
  }

  private userCanManageSettings(): boolean {
    return this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
  }

  private getSettingsService(): PreviewSettingsService {
    if (!this.settingsService) {
      this.settingsService = new PreviewSettingsService(this.context.spHttpClient, this.getConfigWebAbsoluteUrl());
    }

    return this.settingsService;
  }

  private getConfigWebAbsoluteUrl(): string {
    return normalizeBaseUrl(this.properties.configSiteUrl || getTenantRootSiteUrl(this.context.pageContext.web.absoluteUrl));
  }

  private getAdminScriptContext(): {
    componentId: string;
    configSiteUrl: string;
    currentSiteUrl: string;
    tenantHostName: string;
    tenantId: string;
  } {
    return {
      componentId: COMMAND_SET_COMPONENT_ID,
      configSiteUrl: this.getConfigWebAbsoluteUrl(),
      currentSiteUrl: this.context.pageContext.web.absoluteUrl,
      tenantHostName: new URL(this.context.pageContext.web.absoluteUrl).hostname,
      tenantId: getTenantId(this.context.pageContext.aadInfo)
    };
  }
}

function getSelectedFile(
  row: RowAccessor,
  settings: IPreviewSettings = createDefaultPreviewSettings(DEFAULT_APP_BASE_URL)
): { extensionSettings?: IFileExtensionSettings; fileName: string } {
  const fileName = getSelectedFileName(row);
  return {
    extensionSettings: findExtensionSettings(settings, fileName),
    fileName
  };
}

function getSelectedFileName(row: RowAccessor): string {
  return String(row.getValueByName('FileLeafRef') || row.getValueByName('LinkFilename') || '');
}

function getOpenCommandTitle(extension: string): string {
  const labels: Record<string, string> = {
    '.drawio': 'DrawIO',
    '.mmd': 'Mermaid',
    '.mermaid': 'Mermaid',
    '.ifc': 'IFC',
    '.step': 'STEP',
    '.stp': 'STP'
  };
  const label = labels[extension.toLowerCase()] ?? extension.replace(/^\./, '').toUpperCase();
  return `Open ${label}`;
}

function getSelectedFileServerRelativeUrl(row: RowAccessor, webAbsoluteUrl: string): string {
  const fileRef = row.getValueByName('FileRef');
  const fromFileRef = getServerRelativeUrlFromRowValue(fileRef, webAbsoluteUrl);
  if (fromFileRef) {
    return fromFileRef;
  }

  return getServerRelativeUrlFromRowValue(row.getValueByName('EncodedAbsUrl'), webAbsoluteUrl);
}

function getTenantId(aadInfo: unknown): string {
  const candidate = aadInfo as { tenantId?: { toString: () => string } };
  return candidate.tenantId?.toString() || '';
}

function getTenantRootSiteUrl(webAbsoluteUrl: string): string {
  const parsed = new URL(webAbsoluteUrl);
  return `${parsed.protocol}//${parsed.hostname}`;
}
