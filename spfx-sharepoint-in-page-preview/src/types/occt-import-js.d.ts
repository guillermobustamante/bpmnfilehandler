// Local type declaration for occt-import-js (no @types package available).
// Full types are defined inline in StepViewerDialog.ts via OcctInitFn / OcctResult.
declare module 'occt-import-js' {
  const init: (options?: { locateFile?: (path: string) => string }) => Promise<unknown>;
  export default init;
}
