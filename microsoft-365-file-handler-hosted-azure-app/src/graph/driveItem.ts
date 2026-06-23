export type DriveItemMetadata = {
  id: string;
  name: string;
  eTag?: string;
  webUrl?: string;
  size?: number;
  parentReference?: {
    driveId?: string;
    id?: string;
    path?: string;
    siteId?: string;
  };
  lastModifiedDateTime?: string;
  lastModifiedBy?: {
    user?: {
      displayName?: string;
      email?: string;
    };
  };
};

export async function getDriveItemMetadata(itemUrl: string, accessToken: string): Promise<DriveItemMetadata> {
  const response = await graphFetch(itemUrl, accessToken, {
    headers: {
      Accept: "application/json"
    }
  });

  return (await response.json()) as DriveItemMetadata;
}

export async function getDriveItemContent(itemUrl: string, accessToken: string): Promise<string> {
  const response = await graphFetch(`${trimTrailingSlash(itemUrl)}/content`, accessToken, {
    headers: {
      Accept: "application/xml,text/xml,text/plain,*/*"
    }
  });

  return await response.text();
}

export async function putDriveItemContent(
  itemUrl: string,
  accessToken: string,
  xml: string,
  eTag?: string
): Promise<DriveItemMetadata> {
  const headers: Record<string, string> = {
    "Content-Type": "application/xml; charset=utf-8"
  };

  if (eTag) {
    headers["If-Match"] = eTag;
  }

  const response = await graphFetch(`${trimTrailingSlash(itemUrl)}/content`, accessToken, {
    body: xml,
    headers,
    method: "PUT"
  });

  return (await response.json()) as DriveItemMetadata;
}

export async function resolveDriveItemInput(
  input: string,
  accessToken: string,
  allowedExtensions: string[] = [".bpmn"]
): Promise<{ itemUrl: string; metadata: DriveItemMetadata }> {
  const trimmed = input.trim();
  if (!trimmed) {
    throw new Error("Enter a SharePoint or OneDrive BPMN file link.");
  }

  if (trimmed.startsWith("https://graph.microsoft.com/v1.0/")) {
    const metadata = await getDriveItemMetadata(trimmed, accessToken);
    assertAllowedFile(metadata.name, allowedExtensions);
    return { itemUrl: trimmed, metadata };
  }

  const metadata = await resolveSharingUrl(trimmed, accessToken);
  assertAllowedFile(metadata.name, allowedExtensions);

  const driveId = metadata.parentReference?.driveId;
  if (!driveId) {
    throw new Error("Microsoft Graph did not return a drive ID for this file.");
  }

  return {
    itemUrl: `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(metadata.id)}`,
    metadata
  };
}

async function graphFetch(url: string, accessToken: string, init: RequestInit): Promise<Response> {
  const response = await fetch(url, {
    ...init,
    cache: "no-store",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(init.headers || {})
    }
  });

  if (!response.ok) {
    const message = await readGraphError(response);
    throw new Error(message);
  }

  return response;
}

async function readGraphError(response: Response): Promise<string> {
  const body = await response.text().catch(() => "");
  if (!body) {
    return `Microsoft Graph returned ${response.status}.`;
  }

  try {
    const parsed = JSON.parse(body) as { error?: { message?: string } };
    return parsed.error?.message || body;
  } catch {
    return body;
  }
}

function trimTrailingSlash(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

async function resolveSharingUrl(webUrl: string, accessToken: string): Promise<DriveItemMetadata> {
  let parsed: URL;
  try {
    parsed = new URL(webUrl);
  } catch {
    throw new Error("Enter a valid SharePoint or OneDrive file URL.");
  }

  if (!parsed.hostname.endsWith(".sharepoint.com") && !parsed.hostname.endsWith(".sharepoint-df.com")) {
    throw new Error("Only SharePoint or OneDrive for Business links are supported.");
  }

  const shareId = `u!${base64UrlEncode(webUrl)}`;
  const response = await graphFetch(
    `https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem?$select=id,name,eTag,webUrl,size,parentReference,lastModifiedDateTime,lastModifiedBy`,
    accessToken,
    {
      headers: {
        Accept: "application/json"
      }
    }
  );

  return (await response.json()) as DriveItemMetadata;
}

function assertAllowedFile(name: string, allowedExtensions: string[]): void {
  const normalizedName = name.toLowerCase();
  const normalizedExtensions = allowedExtensions.map((extension) => normalizeExtension(extension)).filter(Boolean);
  if (!normalizedExtensions.some((extension) => normalizedName.endsWith(extension))) {
    throw new Error(`The selected file is not supported by this viewer. Supported extension: ${normalizedExtensions.join(", ")}.`);
  }
}

function normalizeExtension(value: string): string {
  const trimmed = value.trim().toLowerCase();
  if (!trimmed) {
    return "";
  }

  return trimmed.startsWith(".") ? trimmed : `.${trimmed}`;
}

function base64UrlEncode(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  const chunkSize = 0x8000;

  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...chunk);
  }

  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}
