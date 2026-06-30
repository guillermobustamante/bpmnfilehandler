// Pure-function tests for sharePointFileService.ts.
// getServerRelativeUrlFromRowValue is a pure function — no SharePoint mocks needed.
jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: { configurations: { v1: {} } }
}));

import { getServerRelativeUrlFromRowValue } from '../extensions/bpmnOpenCommandSet/sharePointFileService';

const WEB_URL = 'https://contoso.sharepoint.com/sites/mysite';

describe('getServerRelativeUrlFromRowValue', () => {
  it('returns value directly when it starts with /', () => {
    expect(getServerRelativeUrlFromRowValue('/sites/mysite/Documents/file.bpmn', WEB_URL)).toBe(
      '/sites/mysite/Documents/file.bpmn'
    );
  });

  it('extracts pathname from an absolute URL', () => {
    const absoluteUrl = 'https://contoso.sharepoint.com/sites/mysite/Documents/file.bpmn';
    expect(getServerRelativeUrlFromRowValue(absoluteUrl, WEB_URL)).toBe('/sites/mysite/Documents/file.bpmn');
  });

  it('returns empty string for empty input', () => {
    expect(getServerRelativeUrlFromRowValue('', WEB_URL)).toBe('');
  });

  it('returns empty string for non-string input', () => {
    expect(getServerRelativeUrlFromRowValue(null, WEB_URL)).toBe('');
    expect(getServerRelativeUrlFromRowValue(undefined, WEB_URL)).toBe('');
    expect(getServerRelativeUrlFromRowValue(42, WEB_URL)).toBe('');
  });

  it('resolves a relative string against the base URL when it does not start with /', () => {
    // The URL constructor treats 'not a url' as a relative URL resolved against the base.
    // The function returns the resolved pathname, not an empty string.
    const result = getServerRelativeUrlFromRowValue('not a url', WEB_URL);
    expect(typeof result).toBe('string');
    expect(result.startsWith('/')).toBe(true);
  });

  it('handles URL-encoded characters in pathname', () => {
    const encoded = 'https://contoso.sharepoint.com/sites/mysite/Documents/my%20file.bpmn';
    expect(getServerRelativeUrlFromRowValue(encoded, WEB_URL)).toBe('/sites/mysite/Documents/my file.bpmn');
  });
});
