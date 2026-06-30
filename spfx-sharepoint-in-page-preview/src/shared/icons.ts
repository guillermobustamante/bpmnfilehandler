const ICON_PATHS: Record<string, string> = {
  check: '<path d="M5 12l4 4L19 6" />',
  download: '<path d="M12 3v12" /><path d="M7 10l5 5 5-5" /><path d="M5 21h14" />',
  external: '<path d="M14 3h7v7" /><path d="M21 3l-9 9" /><path d="M19 14v5H5V5h5" />',
  maximize: '<path d="M4 9V4h5" /><path d="M20 9V4h-5" /><path d="M4 15v5h5" /><path d="M20 15v5h-5" />',
  refresh: '<path d="M20 6v5h-5" /><path d="M4 18v-5h5" /><path d="M18 9a6 6 0 0 0-10-3L4 10" /><path d="M6 15a6 6 0 0 0 10 3l4-4" />',
  restore: '<path d="M9 3H4v5" /><path d="M4 3l7 7" /><path d="M15 21h5v-5" /><path d="M20 21l-7-7" />',
  save: '<path d="M5 3h12l2 2v16H5z" /><path d="M8 3v6h8" /><path d="M8 21v-7h8v7" />',
  zoomIn: '<circle cx="10.5" cy="10.5" r="6.5" /><path d="M10.5 7.5v6" /><path d="M7.5 10.5h6" /><path d="M15.5 15.5L21 21" />',
  zoomOut: '<circle cx="10.5" cy="10.5" r="6.5" /><path d="M7.5 10.5h6" /><path d="M15.5 15.5L21 21" />'
};

export function renderIcon(name: string): string {
  return `<svg aria-hidden="true" focusable="false" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${ICON_PATHS[name] || ''}</svg>`;
}
