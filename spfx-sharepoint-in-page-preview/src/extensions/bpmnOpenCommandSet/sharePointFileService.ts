import { SPHttpClient, type SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISharePointFileMetadata {
  eTag?: string;
  length?: string;
  name?: string;
  serverRelativeUrl: string;
  timeLastModified?: string;
}

export class SharePointFileService {
  public constructor(private readonly spHttpClient: SPHttpClient, private readonly webAbsoluteUrl: string) {}

  public async getMetadata(serverRelativeUrl: string): Promise<ISharePointFileMetadata> {
    const response = await this.spHttpClient.get(
      `${this.fileApiUrl(serverRelativeUrl)}?$select=ETag,Length,Name,ServerRelativeUrl,TimeLastModified`,
      SPHttpClient.configurations.v1
    );

    await assertOk(response);
    const payload = (await response.json()) as {
      ETag?: string;
      Length?: string;
      Name?: string;
      ServerRelativeUrl?: string;
      TimeLastModified?: string;
    };

    return {
      eTag: payload.ETag,
      length: payload.Length,
      name: payload.Name,
      serverRelativeUrl: payload.ServerRelativeUrl || serverRelativeUrl,
      timeLastModified: payload.TimeLastModified
    };
  }

  public async getContent(serverRelativeUrl: string): Promise<string> {
    const response = await this.spHttpClient.get(`${this.fileApiUrl(serverRelativeUrl)}/$value`, SPHttpClient.configurations.v1);
    await assertOk(response);
    return await response.text();
  }

  public async getContentAsArrayBuffer(serverRelativeUrl: string): Promise<ArrayBuffer> {
    const response = await this.spHttpClient.get(`${this.fileApiUrl(serverRelativeUrl)}/$value`, SPHttpClient.configurations.v1);
    await assertOk(response);
    return await response.arrayBuffer();
  }

  public async saveContent(serverRelativeUrl: string, xml: string, eTag?: string): Promise<ISharePointFileMetadata> {
    const response = await this.spHttpClient.post(`${this.fileApiUrl(serverRelativeUrl)}/$value`, SPHttpClient.configurations.v1, {
      body: xml,
      headers: {
        'Content-Type': 'application/xml; charset=utf-8',
        'IF-MATCH': eTag || '*',
        'X-HTTP-Method': 'PUT'
      }
    });

    await assertOk(response);
    return await this.getMetadata(serverRelativeUrl);
  }

  private fileApiUrl(serverRelativeUrl: string): string {
    return `${this.webApiUrl()}/GetFileByServerRelativePath(decodedUrl='${encodeODataString(serverRelativeUrl)}')`;
  }

  private webApiUrl(): string {
    return `${this.webAbsoluteUrl.replace(/\/+$/, '')}/_api/web`;
  }
}

export function getServerRelativeUrlFromRowValue(value: unknown, webAbsoluteUrl: string): string {
  if (typeof value !== 'string' || value.length === 0) {
    return '';
  }

  if (value.startsWith('/')) {
    return value;
  }

  try {
    const parsed = new URL(value, webAbsoluteUrl);
    return decodeURIComponent(parsed.pathname);
  } catch {
    return '';
  }
}

async function assertOk(response: SPHttpClientResponse): Promise<void> {
  if (response.ok) {
    return;
  }

  const body = await response.text().catch(() => '');
  throw new Error(body || `SharePoint returned ${response.status}.`);
}

function encodeODataString(value: string): string {
  return value.replace(/'/g, "''");
}
