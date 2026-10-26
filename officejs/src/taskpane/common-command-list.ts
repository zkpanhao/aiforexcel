interface Handlers {
  select(command: string): void;
  update(command: string, replacement: string): Promise<void>;
  remove(command: string): Promise<void>;
}

/** DOM buttons give each context menu an exact target, including in WebView2. */
export class CommonCommandList {
  private readonly list: HTMLElement;
  private readonly menu: HTMLElement;
  private readonly editor: HTMLElement;
  private readonly input: HTMLTextAreaElement;
  private readonly error: HTMLElement;
  private readonly edit: HTMLButtonElement;
  private readonly remove: HTMLButtonElement;
  private readonly save: HTMLButtonElement;
  private readonly cancel: HTMLButtonElement;
  private rows: HTMLButtonElement[] = [];
  private target?: { command: string; row: HTMLButtonElement };
  private editing?: string;
  private disabled = false;
  private pending = false;

  constructor(private readonly doc: Document, private readonly host: Window, private readonly handlers: Handlers) {
    const get = <T extends HTMLElement>(id: string) => doc.getElementById(id) as T;
    this.list = get('commonCommandList'); this.menu = get('commandMenu'); this.editor = get('commandEditor');
    this.input = get<HTMLTextAreaElement>('commandEditText'); this.error = get('commandError');
    this.edit = get<HTMLButtonElement>('commandEdit'); this.remove = get<HTMLButtonElement>('commandDelete');
    this.save = get<HTMLButtonElement>('commandEditSave'); this.cancel = get<HTMLButtonElement>('commandEditCancel');
    this.list.hidden = this.menu.hidden = this.editor.hidden = this.error.hidden = true;
    this.edit.onclick = () => {
      if (this.blocked() || !this.target) return;
      this.editing = this.target.command; this.input.value = this.editing;
      this.closeMenu(); this.clearError(); this.editor.hidden = false; this.input.focus();
    };
    this.remove.onclick = () => {
      if (this.blocked() || !this.target) return;
      const command = this.target.command;
      const closeEditor = this.editing === command;
      this.closeMenu(); void this.commit(() => this.handlers.remove(command), closeEditor);
    };
    this.save.onclick = () => {
      if (this.blocked() || this.editing === undefined) return;
      const value = this.input.value.trim();
      if (!value || value.length > 2000) { this.showError(!value ? '命令内容不能为空' : '单条常用命令最多 2000 字符'); return; }
      const original = this.editing;
      void this.commit(() => this.handlers.update(original, value), true);
    };
    this.cancel.onclick = () => { if (!this.blocked()) this.closeEditor(); };
    this.input.addEventListener('keydown', event => {
      if (event.key === 'Escape' && !this.blocked()) { event.preventDefault(); this.closeEditor(); }
      else if (event.ctrlKey && event.key === 'Enter') { event.preventDefault(); this.save.click(); }
    });
    this.menu.addEventListener('keydown', event => {
      if (event.key === 'Escape') { event.preventDefault(); this.closeMenu(true); }
      else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
        event.preventDefault(); (doc.activeElement === this.edit ? this.remove : this.edit).focus();
      } else if (event.key === 'Tab') this.closeMenu();
    });
    doc.addEventListener('pointerdown', event => { if (!this.menu.contains(event.target as Node)) this.closeMenu(); });
    doc.addEventListener('scroll', () => this.closeMenu(), true);
    host.addEventListener('resize', () => this.closeMenu());
    host.addEventListener('blur', () => this.closeMenu());
  }

  get hidden(): boolean { return this.list.hidden; }
  open(): void { this.list.hidden = false; this.clearError(); this.rows[0]?.focus(); }
  close(): void { this.list.hidden = true; this.closeMenu(); this.closeEditor(); }
  render(commands: string[]): void {
    this.closeMenu(); this.rows = []; this.list.replaceChildren();
    if (!commands.length) {
      const empty = this.doc.createElement('small'); empty.textContent = '暂无已保存命令'; this.list.append(empty);
    }
    for (const command of commands) {
      const row = this.doc.createElement('button'); row.type = 'button'; row.className = 'common-command-item';
      row.textContent = command.replace(/\s+/g, ' ').slice(0, 80); row.title = command;
      row.setAttribute('aria-haspopup', 'menu'); row.disabled = this.blocked();
      row.onclick = () => { if (!this.blocked()) { this.handlers.select(command); this.close(); } };
      row.addEventListener('contextmenu', event => {
        event.preventDefault(); if (!this.blocked()) this.openMenu(command, row, event.clientX, event.clientY);
      });
      row.addEventListener('keydown', event => {
        if (event.key === 'ContextMenu' || (event.shiftKey && event.key === 'F10')) {
          event.preventDefault(); const rect = row.getBoundingClientRect();
          if (!this.blocked()) this.openMenu(command, row, rect.left, rect.bottom);
        }
      });
      this.rows.push(row); this.list.append(row);
    }
  }
  setDisabled(disabled: boolean): void {
    this.disabled = disabled;
    for (const row of this.rows) row.disabled = this.blocked();
    this.edit.disabled = this.remove.disabled = this.save.disabled = this.cancel.disabled = this.input.disabled = this.blocked();
    if (disabled) this.closeMenu();
  }
  private blocked(): boolean { return this.disabled || this.pending; }
  private openMenu(command: string, row: HTMLButtonElement, x: number, y: number): void {
    this.target = { command, row }; this.clearError(); this.menu.hidden = false;
    const rect = this.menu.getBoundingClientRect();
    this.menu.style.left = `${Math.max(4, Math.min(x, this.host.innerWidth - rect.width - 4))}px`;
    this.menu.style.top = `${Math.max(4, Math.min(y, this.host.innerHeight - rect.height - 4))}px`;
    this.edit.focus();
  }
  private closeMenu(restoreFocus = false): void {
    if (restoreFocus) this.target?.row.focus();
    this.menu.hidden = true; this.target = undefined;
  }
  private closeEditor(): void { this.editor.hidden = true; this.editing = undefined; this.input.value = ''; this.clearError(); }
  private clearError(): void { this.error.hidden = true; this.error.textContent = ''; }
  private showError(message: string): void { this.error.textContent = message; this.error.hidden = false; }
  private async commit(action: () => Promise<void>, editing: boolean): Promise<void> {
    this.pending = true; this.clearError(); this.setDisabled(this.disabled);
    try { await action(); if (editing) this.closeEditor(); }
    catch (error) { this.showError(error instanceof Error ? error.message : '常用命令修改失败，请重试。'); }
    finally { this.pending = false; this.setDisabled(this.disabled); }
  }
}
