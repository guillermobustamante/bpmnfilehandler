const DIALOG_BASE_STYLE_ID = 'bpf-dialog-base-styles';

export function ensureDialogBaseStyles(container: HTMLElement): void {
  if (container.querySelector('#' + DIALOG_BASE_STYLE_ID)) {
    return;
  }
  const style = document.createElement('style');
  style.id = DIALOG_BASE_STYLE_ID;
  style.textContent = [
    'dialog.bpf-viewer-dialog{all:unset;box-sizing:border-box;display:flex;flex-direction:column;',
    'height:100dvh;inset:0;max-height:100dvh;max-width:100vw;overflow:hidden;position:fixed;width:100vw;}',
    'dialog.bpf-viewer-dialog::backdrop{background:rgba(0,0,0,.6);}'
  ].join('');
  container.appendChild(style);
}
