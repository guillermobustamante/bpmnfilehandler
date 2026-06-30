/**
 * Unit tests for pure utility functions exported from IfcViewerDialog.ts.
 *
 * Covers:
 *  - sanitizeIfcData  — STEP-aware sanitizer: removes bytes > 127 outside string
 *                       literals, replaces with space inside them.
 *  - colorKey         — geometry grouping by colour for draw-call reduction
 *  - formatBytes      — human-readable file size display
 *  - escapeHtml       — XSS prevention in dialog HTML
 */

// ── SPFx module stubs ─────────────────────────────────────────────────────────
jest.mock('@microsoft/sp-http', () => ({ SPHttpClient: { configurations: { v1: {} } } }));
jest.mock('@microsoft/sp-dialog', () => ({ BaseDialog: class {} }));
jest.mock('../shared/icons', () => ({ renderIcon: () => '' }));
jest.mock('../extensions/bpmnOpenCommandSet/previewSettings', () => ({}));
jest.mock('../extensions/bpmnOpenCommandSet/sharePointFileService', () => ({}));

import {
  sanitizeIfcData,
  colorKey,
  formatBytes,
  escapeHtml
} from '../extensions/bpmnOpenCommandSet/IfcViewerDialog';

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Build a Uint8Array from a plain ASCII string. */
function fromAscii(str: string): Uint8Array {
  const arr = new Uint8Array(str.length);
  for (let i = 0; i < str.length; i++) arr[i] = str.charCodeAt(i);
  return arr;
}

/** Concatenate multiple Uint8Arrays and plain byte values into one Uint8Array. */
function concatBytes(...parts: Array<Uint8Array | number>): Uint8Array {
  let total = 0;
  for (const p of parts) total += typeof p === 'number' ? 1 : p.length;
  const out = new Uint8Array(total);
  let pos = 0;
  for (const p of parts) {
    if (typeof p === 'number') {
      out[pos++] = p;
    } else {
      out.set(p, pos);
      pos += p.length;
    }
  }
  return out;
}

/** Decode a Uint8Array back to a string (printable ASCII kept; others shown as [N]). */
function toAscii(data: Uint8Array): string {
  let s = '';
  for (let i = 0; i < data.length; i++) {
    const b = data[i];
    s += b >= 32 && b <= 126 ? String.fromCharCode(b) : '[' + b + ']';
  }
  return s;
}

/** Return a copy of data with value injected at index. */
function withByteAt(data: Uint8Array, index: number, value: number): Uint8Array {
  const copy = new Uint8Array(data);
  copy[index] = value;
  return copy;
}

/**
 * Build a minimal structurally-valid IFC/STEP byte sequence:
 *   ISO-10303-21;\nHEADER;\n<headerExtra>ENDSEC;\nDATA;\n<dataExtra>ENDSEC;\nEND-ISO-10303-21;
 */
function makeIfcBytes(headerExtra: string = '', dataExtra: string = ''): Uint8Array {
  const content =
    'ISO-10303-21;\nHEADER;\n' + headerExtra +
    'ENDSEC;\nDATA;\n' + dataExtra +
    'ENDSEC;\nEND-ISO-10303-21;\n';
  return fromAscii(content);
}

// ── sanitizeIfcData ───────────────────────────────────────────────────────────

describe('sanitizeIfcData', () => {

  // ── Fast path ──────────────────────────────────────────────────────────────

  it('returns the SAME reference for an all-ASCII file (no copy)', () => {
    const data = makeIfcBytes();
    expect(sanitizeIfcData(data)).toBe(data);
  });

  it('returns the SAME reference for an empty array', () => {
    const empty = new Uint8Array(0);
    expect(sanitizeIfcData(empty)).toBe(empty);
  });

  it('handles a large all-ASCII file without allocating (same reference)', () => {
    const large = new Uint8Array(1024 * 1024).fill(65);
    expect(sanitizeIfcData(large)).toBe(large);
  });

  // ── Outside-string removal ─────────────────────────────────────────────────
  // Bytes > 127 outside single-quoted STEP strings must be REMOVED entirely.
  // Inserting a space would split tokens (references, numbers, type names).

  it('removes a byte > 127 outside a string literal (not replaced with space)', () => {
    // Pure bytes with no single-quotes — everything is "outside a string"
    const dirty = new Uint8Array([65, 155, 66]); // A, 0x9B, B
    const result = sanitizeIfcData(dirty);
    expect(Array.from(result)).toEqual([65, 66]); // byte removed, length is 2
  });

  it('removes ALL bytes > 127 that are outside string literals', () => {
    const data = new Uint8Array([128, 200, 255, 65, 127]);
    const result = sanitizeIfcData(data);
    expect(Array.from(result)).toEqual([65, 127]); // 128/200/255 removed
  });

  it('preserves byte value 127 (last 7-bit ASCII)', () => {
    const data = new Uint8Array([127, 128]);
    const result = sanitizeIfcData(data);
    expect(result.length).toBe(1);
    expect(result[0]).toBe(127);
  });

  it('does NOT mutate the original array', () => {
    const original = new Uint8Array([65, 200, 66]);
    const snapshot = new Uint8Array(original);
    sanitizeIfcData(original);
    expect(Array.from(original)).toEqual(Array.from(snapshot));
  });

  // ── Key scenario: broken STEP reference ───────────────────────────────────
  // Hypothesis: byte 0x9B was inserted INSIDE a reference token like #155,
  // making the raw bytes '#1' + 0x9B + '55'. Space-replacement gives '#1 55'
  // (broken). Removal restores '#155' (the intended reference).

  it('removes 0x9B from a STEP entity reference, restoring "#155"', () => {
    const withBad = concatBytes(fromAscii('#1'), 0x9b, fromAscii('55=IFCWALL($);'));
    const result = sanitizeIfcData(withBad);
    expect(toAscii(result)).toBe('#155=IFCWALL($);');
  });

  it('removes 0x9B from an entity type name outside a string', () => {
    const dirty = concatBytes(fromAscii('IFCWALL'), 0x9b, fromAscii('TYPE('));
    const result = sanitizeIfcData(dirty);
    expect(toAscii(result)).toBe('IFCWALLTYPE(');
  });

  // ── Inside-string replacement ──────────────────────────────────────────────
  // Bytes > 127 inside STEP single-quoted strings are replaced with space (0x20).

  it('replaces a byte > 127 with space INSIDE a string literal', () => {
    const dirty = concatBytes(fromAscii("'Stra"), 0x9b, fromAscii("e'"));
    const result = sanitizeIfcData(dirty);
    expect(toAscii(result)).toBe("'Stra e'");
  });

  it('replaces ALL bytes > 127 inside a string with spaces', () => {
    const dirty = concatBytes(fromAscii("'"), 200, 201, 202, fromAscii("'"));
    const result = sanitizeIfcData(dirty);
    expect(toAscii(result)).toBe("'   '");
  });

  it('leaves string boundaries (the single-quote chars) unchanged', () => {
    const data = concatBytes(fromAscii("'hello'"));
    const result = sanitizeIfcData(data);
    expect(result).toBe(data); // all-ASCII fast path
  });

  // ── Mixed contexts ─────────────────────────────────────────────────────────

  it('removes byte outside string but replaces byte inside string in same input', () => {
    // FILE_NAME('[0x9B]',$);
    //           ^inside    ^ outside (after closing quote) - no dirty byte here
    const dirty = concatBytes(fromAscii("FILE_NAME('"), 0x9b, fromAscii("',$);"));
    const result = sanitizeIfcData(dirty);
    expect(toAscii(result)).toBe("FILE_NAME(' ',$);");
    // Length stays the same because the byte inside the string was replaced (not removed)
    expect(result.length).toBe(dirty.length);
  });

  it('correctly switches context across multiple strings', () => {
    // #1[0x9B]=[0xAA]'abc[0xBB]def'[0xCC];
    // outside: 0x9B removed, 0xAA removed, 0xCC removed
    // inside:  0xBB -> space
    const dirty = new Uint8Array([
      35, 49,         // #1
      0x9b,           // outside string -> remove
      61,             // =
      0xaa,           // outside string -> remove
      39,             // '  (open string)
      97, 98, 99,    // abc
      0xbb,           // inside string -> space
      100, 101, 102, // def
      39,             // '  (close string)
      0xcc,           // outside string -> remove
      59              // ;
    ]);
    const result = sanitizeIfcData(dirty);
    expect(toAscii(result)).toBe("#1='abc def';");
  });

  // ── Escaped-quote handling ─────────────────────────────────────────────────
  // STEP rule: '' inside a string represents a literal apostrophe.
  // The scanner must NOT exit string mode when it sees ''.

  it("handles escaped quotes ('') inside a string without premature exit", () => {
    // 'it''s' followed by 0x9B outside the string
    const dirty = concatBytes(fromAscii("'it''s'"), 0x9b);
    const result = sanitizeIfcData(dirty);
    // 0x9B was outside the final closing quote -> removed
    expect(toAscii(result)).toBe("'it''s'");
    expect(result.length).toBe(7); // the 0x9B was removed
  });

  it('replaces a dirty byte that follows an escaped quote (still inside string)', () => {
    // 'it''[0x9B]s' — the '' is escaped apostrophe; 0x9B is still inside string
    const dirty = concatBytes(fromAscii("'it''"), 0x9b, fromAscii("s'"));
    const result = sanitizeIfcData(dirty);
    expect(toAscii(result)).toBe("'it'' s'");
    // Length unchanged (replaced not removed)
    expect(result.length).toBe(dirty.length);
  });

  // ── Real-world crash scenario ──────────────────────────────────────────────

  it('fixes the wall-with-opening-and-window.ifc scenario: 0x9B outside string in DATA section', () => {
    const base = makeIfcBytes(
      "FILE_DESCRIPTION(('IFC4'),'2;1');\n",
      "#1=IFCPROJECT('0HY9P$MnT7$h38C89YIjmM',$,'WallProject',$,$,$,$,(#2),#3);\n"
    );

    // Find DATA;\n and inject 0x9B right at the start of the DATA section
    // (before the first '#', so it's outside any string literal)
    const marker = fromAscii('DATA;\n');
    let dataOffset = -1;
    for (let i = 0; i <= base.length - marker.length; i++) {
      let match = true;
      for (let k = 0; k < marker.length; k++) {
        if (base[i + k] !== marker[k]) { match = false; break; }
      }
      if (match) { dataOffset = i + marker.length; break; }
    }
    expect(dataOffset).toBeGreaterThan(0);

    const dirty = withByteAt(base, dataOffset, 155); // inject before '#'
    expect(dirty[dataOffset]).toBe(155);

    const result = sanitizeIfcData(dirty);

    // Byte was outside a string -> REMOVED, so result is one byte shorter
    expect(result.length).toBe(base.length - 1);
    // No remaining bytes > 127
    let remaining = 0;
    for (let i = 0; i < result.length; i++) if (result[i] > 127) remaining++;
    expect(remaining).toBe(0);
  });

  it('handles a large file with a single dirty byte outside a string', () => {
    const large = new Uint8Array(512 * 1024).fill(65); // all 'A'
    large[large.length - 1] = 200; // dirty byte at end, outside string
    const result = sanitizeIfcData(large);
    expect(result).not.toBe(large);
    expect(result.length).toBe(large.length - 1); // removed
    let remaining = 0;
    for (let i = 0; i < result.length; i++) if (result[i] > 127) remaining++;
    expect(remaining).toBe(0);
  });

  it('handles a large file with a single dirty byte inside a string', () => {
    // Build: 'AAAA...A[0x9B]AAAA...A' (all inside a string)
    const inner = new Uint8Array(512 * 1024 - 2).fill(65); // all 'A'
    inner[250] = 200; // dirty byte inside
    const data = concatBytes(fromAscii("'"), inner, fromAscii("'"));
    const result = sanitizeIfcData(data);
    expect(result.length).toBe(data.length); // replaced not removed -> same length
    expect(result[251]).toBe(32); // space
    let remaining = 0;
    for (let i = 0; i < result.length; i++) if (result[i] > 127) remaining++;
    expect(remaining).toBe(0);
  });
});

// ── colorKey ─────────────────────────────────────────────────────────────────

describe('colorKey', () => {
  it('returns a string', () => {
    expect(typeof colorKey(1, 0, 0, 1)).toBe('string');
  });

  it('produces the same key for colours within the 2-decimal rounding band', () => {
    expect(colorKey(0.5, 0.5, 0.5, 1.0)).toBe(colorKey(0.501, 0.5, 0.5, 1.0));
  });

  it('produces different keys for colours outside the 2-decimal rounding band', () => {
    expect(colorKey(0.50, 0.5, 0.5, 1.0)).not.toBe(colorKey(0.51, 0.5, 0.5, 1.0));
  });

  it('includes the alpha component in the key', () => {
    expect(colorKey(1, 1, 1, 1.0)).not.toBe(colorKey(1, 1, 1, 0.5));
  });

  it('produces deterministic output for known inputs', () => {
    expect(colorKey(1, 0, 0, 1)).toBe('1.00,0.00,0.00,1.00');
    expect(colorKey(0.5, 0.25, 0.75, 0.5)).toBe('0.50,0.25,0.75,0.50');
  });

  it('handles zero alpha (transparent geometry)', () => {
    expect(colorKey(0, 0, 0, 0)).toBe('0.00,0.00,0.00,0.00');
  });
});

// ── formatBytes ───────────────────────────────────────────────────────────────

describe('formatBytes', () => {
  it('formats bytes below 1 KB as "<n> B"', () => {
    expect(formatBytes(0)).toBe('0 B');
    expect(formatBytes(512)).toBe('512 B');
    expect(formatBytes(1023)).toBe('1023 B');
  });

  it('formats sizes between 1 KB and 1 MB as "x KB"', () => {
    expect(formatBytes(1024)).toBe('1 KB');
    expect(formatBytes(1024 * 10)).toBe('10 KB');
    expect(formatBytes(1024 * 1024 - 1)).toMatch(/KB$/);
  });

  it('formats sizes >= 1 MB as "x MB"', () => {
    expect(formatBytes(1024 * 1024)).toBe('1 MB');
    expect(formatBytes(1024 * 1024 * 300)).toBe('300 MB');
  });

  it('returns empty string for non-finite values', () => {
    expect(formatBytes(NaN)).toBe('');
    expect(formatBytes(Infinity)).toBe('');
    expect(formatBytes(-Infinity)).toBe('');
  });
});

// ── escapeHtml ────────────────────────────────────────────────────────────────

describe('escapeHtml', () => {
  it('escapes & < > " \'', () => {
    expect(escapeHtml('&')).toBe('&amp;');
    expect(escapeHtml('<')).toBe('&lt;');
    expect(escapeHtml('>')).toBe('&gt;');
    expect(escapeHtml('"')).toBe('&quot;');
    expect(escapeHtml("'")).toBe('&#039;');
  });

  it('leaves plain text unchanged', () => {
    expect(escapeHtml('Hello World')).toBe('Hello World');
  });

  it('escapes all dangerous characters in a combined XSS payload', () => {
    expect(escapeHtml('<script>alert("xss")</script>')).toBe(
      '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;'
    );
  });

  it('handles empty string', () => {
    expect(escapeHtml('')).toBe('');
  });

  it('escapes a typical file name with multiple special characters', () => {
    const name = 'Wall & Window "Test" <Model>';
    const escaped = escapeHtml(name);
    expect(escaped).toBe('Wall &amp; Window &quot;Test&quot; &lt;Model&gt;');
    expect(escaped).not.toContain('<');
    expect(escaped).not.toContain('>');
    expect(escaped).not.toContain('"');
  });

  it('escapes a single quote (apostrophe) in a file name', () => {
    expect(escapeHtml("O'Brien's Building.ifc")).toBe("O&#039;Brien&#039;s Building.ifc");
  });
});
