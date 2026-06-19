export type DriveItemMetadata = {
  id: string;
  name: string;
  eTag?: string;
  webUrl?: string;
  size?: number;
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

