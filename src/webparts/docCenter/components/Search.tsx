import * as React from 'react';
import {
  Stack, TextField, DefaultButton, PrimaryButton, IconButton,
  MessageBar, MessageBarType, Spinner, SpinnerSize,
  Dialog, DialogType, DialogFooter
} from '@fluentui/react';
import { IHashtag, IDocument, SearchMode, groupHashtagsByCategory } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import { CatChip, ExtBadge, GroupLabel, Segmented } from './ui';
import styles from './DocCenter.module.scss';

interface IProps {
  libraryTitle: string;
  hashtagsListTitle: string;
  hashtags: IHashtag[];
  siteUrl: string;
  isAdmin: boolean;
}

interface IState {
  selectedTagIds: number[];
  mode: SearchMode;
  nameQuery: string;
  searching: boolean;
  results: IDocument[];
  searched: boolean;
  showSuggestions: boolean;
  highlightedSuggestion: number;
  currentPage: number;
  editingDoc?: IDocument;
  editingTagIds: number[];
  editFilter: string;
  savingEdit: boolean;
  editError?: string;
  deletingDoc?: IDocument;
  deleting: boolean;
  deleteError?: string;
  renamingDoc?: IDocument;
  renameValue: string;
  renaming: boolean;
  renameError?: string;
}

const MAX_SUGGESTIONS = 8;
const SEARCH_DEBOUNCE_MS = 2000;
const PAGE_SIZE = 10;

const EDIT_WINDOW_MS = 24 * 60 * 60 * 1000;

const MODE_OPTIONS = [
  { key: 'all', label: 'Tất cả (VÀ)' },
  { key: 'any', label: 'Bất kỳ (HOẶC)' }
];

export class Search extends React.Component<IProps, IState> {

  public state: IState = {
    selectedTagIds: [],
    mode: 'all',
    nameQuery: '',
    searching: false,
    results: [],
    searched: false,
    showSuggestions: false,
    highlightedSuggestion: -1,
    currentPage: 1,
    editingTagIds: [],
    editFilter: '',
    savingEdit: false,
    deleting: false,
    renameValue: '',
    renaming: false
  };

  private suggestionBlurTimer?: number;
  private searchDebounceTimer?: number;
  private scrollRaf?: number;
  private resultsTopRef = React.createRef<HTMLDivElement>();

  public componentWillUnmount(): void {
    if (this.suggestionBlurTimer !== undefined) {
      window.clearTimeout(this.suggestionBlurTimer);
    }
    if (this.searchDebounceTimer !== undefined) {
      window.clearTimeout(this.searchDebounceTimer);
    }
    if (this.scrollRaf !== undefined) {
      window.cancelAnimationFrame(this.scrollRaf);
    }
  }

  private scheduleSearch = (): void => {
    if (this.searchDebounceTimer !== undefined) {
      window.clearTimeout(this.searchDebounceTimer);
    }
    this.searchDebounceTimer = window.setTimeout(() => {
      this.searchDebounceTimer = undefined;
      void this.runSearch();
    }, SEARCH_DEBOUNCE_MS);
  };

  private getSuggestions = (): IDocument[] => {
    const q = this.state.nameQuery.trim().toLowerCase();
    if (!q) return [];
    return this.state.results
      .filter(d => d.Name.toLowerCase().indexOf(q) !== -1)
      .slice(0, MAX_SUGGESTIONS);
  };

  // Mở tài liệu ở tab mới; chặn hành vi mặc định để SharePoint
  // không chiếm quyền điều hướng và redirect ngay trong tab hiện tại.
  private openDocument = (e: React.MouseEvent, d: IDocument): void => {
    e.preventDefault();
    e.stopPropagation();
    window.open(this.buildAbsoluteUrl(d.ServerRelativeUrl), '_blank', 'noopener,noreferrer');
  };

  private openSuggestion = (d: IDocument): void => {
    window.open(this.buildAbsoluteUrl(d.ServerRelativeUrl), '_blank', 'noopener,noreferrer');
    this.setState({ showSuggestions: false, highlightedSuggestion: -1 });
  };

  private onSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (!this.state.showSuggestions) return;
    const suggestions = this.getSuggestions();
    if (suggestions.length === 0) return;
    const { highlightedSuggestion } = this.state;
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      this.setState({
        highlightedSuggestion: Math.min(highlightedSuggestion + 1, suggestions.length - 1)
      });
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      this.setState({
        highlightedSuggestion: Math.max(highlightedSuggestion - 1, 0)
      });
    } else if (e.key === 'Enter') {
      if (highlightedSuggestion >= 0 && highlightedSuggestion < suggestions.length) {
        e.preventDefault();
        this.openSuggestion(suggestions[highlightedSuggestion]);
      }
    } else if (e.key === 'Escape') {
      this.setState({ showSuggestions: false, highlightedSuggestion: -1 });
    }
  };

  private onSearchFocus = (): void => {
    if (this.suggestionBlurTimer !== undefined) {
      window.clearTimeout(this.suggestionBlurTimer);
      this.suggestionBlurTimer = undefined;
    }
    this.setState({ showSuggestions: true });
  };

  private onSearchBlur = (): void => {
    // Delay hiding so a mousedown on a suggestion can register first.
    this.suggestionBlurTimer = window.setTimeout(() => {
      this.setState({ showSuggestions: false, highlightedSuggestion: -1 });
    }, 150);
  };

  private isWithinEditWindow = (doc: IDocument): boolean => {
    const created = new Date(doc.Created).getTime();
    return Number.isFinite(created) && (Date.now() - created) < EDIT_WINDOW_MS;
  };

  private openRename = (doc: IDocument): void => {
    this.setState({
      renamingDoc: doc,
      renameValue: doc.Name,
      renaming: false,
      renameError: undefined
    });
  };

  private cancelRename = (): void => {
    if (this.state.renaming) return;
    this.setState({ renamingDoc: undefined, renameValue: '', renameError: undefined });
  };

  private confirmRename = async (): Promise<void> => {
    const { renamingDoc, renameValue } = this.state;
    if (!renamingDoc) return;
    const next = renameValue.trim();
    if (!next || next === renamingDoc.Name) {
      this.setState({ renamingDoc: undefined, renameValue: '', renameError: undefined });
      return;
    }
    this.setState({ renaming: true, renameError: undefined });
    try {
      await SharePointService.renameDocument(this.props.libraryTitle, renamingDoc.Id, next);
      this.setState({ renamingDoc: undefined, renameValue: '', renaming: false });
      await this.runSearch();
    } catch (e) {
      this.setState({
        renaming: false,
        renameError: e instanceof Error ? e.message : String(e)
      });
    }
  };

  private openDelete = (doc: IDocument): void => {
    this.setState({ deletingDoc: doc, deleteError: undefined, deleting: false });
  };

  private cancelDelete = (): void => {
    if (this.state.deleting) return;
    this.setState({ deletingDoc: undefined, deleteError: undefined });
  };

  private confirmDelete = async (): Promise<void> => {
    const { deletingDoc } = this.state;
    if (!deletingDoc) return;
    this.setState({ deleting: true, deleteError: undefined });
    try {
      await SharePointService.deleteDocument(this.props.libraryTitle, deletingDoc.Id);
      this.setState({ deletingDoc: undefined, deleting: false });
      await this.runSearch();
    } catch (e) {
      this.setState({
        deleting: false,
        deleteError: e instanceof Error ? e.message : String(e)
      });
    }
  };

  private toggleTag = (id: number): void => {
    const set = new Set(this.state.selectedTagIds);
    if (set.has(id)) set.delete(id); else set.add(id);
    this.setState({ selectedTagIds: Array.from(set) }, () => void this.runSearch());
  };

  private clearFilters = (): void =>
    this.setState({ selectedTagIds: [], nameQuery: '' }, () => void this.runSearch());

  private runSearch = async (): Promise<void> => {
    this.setState({ searching: true });
    try {
      const results = await SharePointService.searchDocuments(
        this.props.libraryTitle,
        this.state.selectedTagIds,
        this.state.mode,
        this.state.nameQuery,
        !this.props.isAdmin
      );
      this.setState({ results, searching: false, searched: true, currentPage: 1 });
    } catch (e) {
      this.setState({ searching: false, searched: true, results: [], currentPage: 1 });
      // eslint-disable-next-line no-console
      console.error(e);
    }
  };

  private goToPage = (page: number): void => {
    this.setState({ currentPage: page }, () => {
      window.requestAnimationFrame(() => this.scrollPageToTop());
    });
  };

  private findScrollContainer = (): HTMLElement | null => {
    let node: HTMLElement | null = this.resultsTopRef.current;
    while (node) {
      const overflowY = window.getComputedStyle(node).overflowY;
      if ((overflowY === 'auto' || overflowY === 'scroll') && node.scrollHeight > node.clientHeight) {
        return node;
      }
      node = node.parentElement;
    }
    return null;
  };

  private getScrollTop = (target: HTMLElement | null): number => {
    if (target) return target.scrollTop;
    return window.scrollY || window.pageYOffset || document.documentElement.scrollTop || 0;
  };

  private setScrollTop = (target: HTMLElement | null, y: number): void => {
    if (target) {
      target.scrollTop = y;
    } else {
      window.scrollTo(0, y);
    }
  };

  private scrollPageToTop = (): void => {
    const target = this.findScrollContainer();
    const start = this.getScrollTop(target);
    if (start <= 0) return;
    if (this.scrollRaf !== undefined) {
      window.cancelAnimationFrame(this.scrollRaf);
    }
    const duration = 350;
    const startTime = performance.now();
    const easeInOutQuad = (t: number): number =>
      t < 0.5 ? 2 * t * t : 1 - Math.pow(-2 * t + 2, 2) / 2;

    const step = (now: number): void => {
      const elapsed = now - startTime;
      const t = Math.min(1, elapsed / duration);
      const y = start * (1 - easeInOutQuad(t));
      this.setScrollTop(target, y);
      if (t < 1) {
        this.scrollRaf = window.requestAnimationFrame(step);
      } else {
        this.scrollRaf = undefined;
      }
    };
    this.scrollRaf = window.requestAnimationFrame(step);
  };

  public componentDidMount(): void {
    void this.runSearch();
  }

  private buildAbsoluteUrl(serverRelativeUrl: string): string {
    const origin = this.props.siteUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
    const absolute = serverRelativeUrl.startsWith('/') ? `${origin}${serverRelativeUrl}` : serverRelativeUrl;
    return absolute + (absolute.indexOf('?') === -1 ? '?web=1' : '&web=1');
  }

  private buildDownloadUrl(serverRelativeUrl: string): string {
    return `${this.props.siteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
  }

  private downloadDocument = (d: IDocument): void => {
    window.open(this.buildDownloadUrl(d.ServerRelativeUrl), '_blank', 'noopener,noreferrer');
  };

  private openEdit = (doc: IDocument): void => {
    this.setState({
      editingDoc: doc,
      editingTagIds: doc.Hashtags.map(t => t.Id),
      editFilter: '',
      savingEdit: false,
      editError: undefined
    });
  };

  private cancelEdit = (): void => {
    if (this.state.savingEdit) return;
    this.setState({ editingDoc: undefined, editingTagIds: [], editError: undefined });
  };

  private toggleEditTag = (id: number): void => {
    const set = new Set(this.state.editingTagIds);
    if (set.has(id)) set.delete(id); else set.add(id);
    this.setState({ editingTagIds: Array.from(set) });
  };

  private saveEdit = async (): Promise<void> => {
    const { editingDoc, editingTagIds } = this.state;
    if (!editingDoc) return;
    this.setState({ savingEdit: true, editError: undefined });
    try {
      await SharePointService.updateDocumentHashtags(this.props.libraryTitle, editingDoc.Id, editingTagIds);
      this.setState({ editingDoc: undefined, editingTagIds: [], savingEdit: false });
      await this.runSearch();
    } catch (e) {
      this.setState({
        savingEdit: false,
        editError: e instanceof Error ? e.message : String(e)
      });
    }
  };

  private getPageNumbers(current: number, total: number): Array<number | '...'> {
    if (total <= 7) {
      const all: number[] = [];
      for (let i = 1; i <= total; i++) all.push(i);
      return all;
    }
    const pages: Array<number | '...'> = [1];
    const start = Math.max(2, current - 1);
    const end = Math.min(total - 1, current + 1);
    if (start > 2) pages.push('...');
    for (let i = start; i <= end; i++) pages.push(i);
    if (end < total - 1) pages.push('...');
    pages.push(total);
    return pages;
  }

  private renderPageButtons(current: number, total: number): React.ReactNode {
    return this.getPageNumbers(current, total).map((p, idx) => {
      if (p === '...') {
        return <span key={`gap-${idx}`} className={styles.pageEllipsis}>…</span>;
      }
      const isActive = p === current;
      return (
        <button
          key={p}
          type="button"
          className={`${styles.pageBtn} ${isActive ? styles.pageBtnActive : ''}`}
          onClick={() => this.goToPage(p)}
          disabled={isActive}
          aria-current={isActive ? 'page' : undefined}
        >
          {p}
        </button>
      );
    });
  }

  private getFilterSummary(): string {
    const { selectedTagIds, mode } = this.state;
    if (selectedTagIds.length === 0) {
      return `tất cả tài liệu trong ${this.props.libraryTitle}`;
    }
    const names = selectedTagIds
      .map(id => this.props.hashtags.filter(t => t.Id === id)[0])
      .filter(t => !!t)
      .map(t => `#${t.Title}`);
    return `khớp ${names.join(mode === 'all' ? ' VÀ ' : ' HOẶC ')}`;
  }

  private renderRowActions(d: IDocument): React.ReactNode {
    const canRename = this.props.isAdmin || this.isWithinEditWindow(d);
    const canEditTags = !this.props.isAdmin && this.isWithinEditWindow(d);
    return (
      <span className={styles.rowActions}>
        <button
          type="button"
          className={styles.sqBtn}
          title="Mở"
          onClick={e => this.openDocument(e, d)}
        >
          ↗
        </button>
        <button
          type="button"
          className={styles.sqBtn}
          title="Tải về"
          onClick={() => this.downloadDocument(d)}
        >
          ↓
        </button>
        {canRename && (
          <button
            type="button"
            className={styles.sqBtn}
            title="Đổi tên tài liệu"
            onClick={() => this.openRename(d)}
          >
            ✎
          </button>
        )}
        {canEditTags && (
          <button
            type="button"
            className={styles.sqBtn}
            title="Sửa hashtag"
            onClick={() => this.openEdit(d)}
          >
            #
          </button>
        )}
        {d.CanDelete && (
          <button
            type="button"
            className={`${styles.sqBtn} ${styles.sqBtnDanger}`}
            title="Xoá tài liệu"
            onClick={() => this.openDelete(d)}
          >
            ✕
          </button>
        )}
      </span>
    );
  }

  public render(): React.ReactElement {
    const {
      selectedTagIds, mode, nameQuery, searching, results, searched,
      showSuggestions, highlightedSuggestion, currentPage,
      editingDoc, editingTagIds, editFilter, savingEdit, editError,
      deletingDoc, deleting, deleteError,
      renamingDoc, renameValue, renaming, renameError
    } = this.state;
    const suggestions = this.getSuggestions();
    const totalPages = Math.max(1, Math.ceil(results.length / PAGE_SIZE));
    const safePage = Math.min(currentPage, totalPages);
    const pageStart = (safePage - 1) * PAGE_SIZE;
    const pageEnd = pageStart + PAGE_SIZE;
    const pagedResults = results.slice(pageStart, pageEnd);
    const hasFilters = selectedTagIds.length > 0 || !!nameQuery.trim();
    const matchTag = (t: IHashtag, q: string): boolean => {
      if (!q) return true;
      const needle = q.toLowerCase();
      return t.Title.toLowerCase().includes(needle)
        || (t.Description || '').toLowerCase().includes(needle);
    };
    const filteredEditTags = this.props.hashtags.filter(t => matchTag(t, editFilter));

    return (
      <div className={styles.section}>
        <div className={styles.card} style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div style={{ display: 'flex', gap: 12 }}>
            <div className={styles.searchAutocomplete}>
              <div className={styles.searchBox}>
                <span className={styles.searchIcon}>⌕</span>
                <input
                  value={nameQuery}
                  autoComplete="off"
                  placeholder="Nhập một phần tên file hoặc hashtag…"
                  onChange={e => this.setState(
                    { nameQuery: e.target.value, showSuggestions: true, highlightedSuggestion: -1 },
                    this.scheduleSearch
                  )}
                  onFocus={this.onSearchFocus}
                  onBlur={this.onSearchBlur}
                  onKeyDown={this.onSearchKeyDown}
                />
              </div>
              {showSuggestions && suggestions.length > 0 && (
                <div className={styles.suggestBox} role="listbox">
                  {suggestions.map((d, idx) => (
                    <div
                      key={d.Id}
                      role="option"
                      aria-selected={idx === highlightedSuggestion}
                      className={`${styles.suggestItem} ${idx === highlightedSuggestion ? styles.suggestItemActive : ''}`}
                      onMouseDown={e => { e.preventDefault(); this.openSuggestion(d); }}
                      onMouseEnter={() => this.setState({ highlightedSuggestion: idx })}
                    >
                      <ExtBadge fileName={d.Name} size={24} />
                      <span className={styles.suggestName}>{d.Name}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <Segmented
              options={MODE_OPTIONS}
              value={mode}
              onChange={key => this.setState({ mode: key as SearchMode }, () => void this.runSearch())}
            />
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            {this.props.hashtags.length === 0 && (
              <span className={styles.muted}>Chưa có hashtag nào.</span>
            )}
            {groupHashtagsByCategory(this.props.hashtags).map(g => (
              <div key={g.name} className={styles.groupRow}>
                <GroupLabel name={g.name} />
                <div className={styles.tagWrap}>
                  {g.items.map(t => {
                    const selected = selectedTagIds.indexOf(t.Id) !== -1;
                    return (
                      <CatChip
                        key={t.Id}
                        tag={t}
                        selected={selected}
                        showRemove={selected}
                        onClick={() => this.toggleTag(t.Id)}
                      />
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className={styles.resultBar}>
          <span className={styles.countPill}>{results.length} tài liệu</span>
          <span>{this.getFilterSummary()}</span>
          {hasFilters && (
            <span className={styles.clearLink} onClick={this.clearFilters}>
              Xoá bộ lọc
            </span>
          )}
          <span style={{ marginLeft: 'auto' }}>
            <button
              type="button"
              className={styles.sqBtn}
              title="Làm mới"
              onClick={() => void this.runSearch()}
              disabled={searching}
            >
              ⟳
            </button>
          </span>
        </div>

        {searching && <Spinner size={SpinnerSize.medium} label="Đang tìm..." />}

        {!searching && searched && results.length === 0 && (
          <div className={styles.emptyBox}>
            Không có tài liệu nào khớp bộ lọc — thử bỏ bớt hashtag hoặc đổi sang <b>Bất kỳ (HOẶC)</b>.
          </div>
        )}

        {!searching && results.length > 0 && (
          <div className={styles.docList} ref={this.resultsTopRef}>
            {pagedResults.map(d => (
              <div key={d.Id} className={styles.docRow}>
                <ExtBadge fileName={d.Name} />
                <span style={{ flex: 1, minWidth: 0 }}>
                  <span className={styles.docTitle}>
                    <a
                      href={this.buildAbsoluteUrl(d.ServerRelativeUrl)}
                      target="_blank"
                      rel="noopener noreferrer"
                      data-interception="off"
                      onClick={e => this.openDocument(e, d)}
                    >
                      {d.Name}
                    </a>
                  </span>
                  <span className={styles.docMeta}>
                    {d.SizeKB ? `${d.SizeKB} KB · ` : ''}
                    {new Date(d.Modified).toLocaleString()}{d.CreatedBy ? ` · ${d.CreatedBy}` : ''}
                  </span>
                </span>
                <span className={styles.docChips}>
                  {d.Hashtags.map(t => (
                    <CatChip
                      key={t.Id}
                      tag={t}
                      small
                      selected={selectedTagIds.indexOf(t.Id) !== -1}
                      onClick={() => this.toggleTag(t.Id)}
                    />
                  ))}
                </span>
                {this.renderRowActions(d)}
              </div>
            ))}
          </div>
        )}

        {!searching && results.length > PAGE_SIZE && (
          <Stack horizontalAlign="center" tokens={{ childrenGap: 6 }} className={styles.pagination}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
              <IconButton
                iconProps={{ iconName: 'ChevronLeft' }}
                title="Trang trước"
                ariaLabel="Trang trước"
                disabled={safePage <= 1}
                onClick={() => this.goToPage(safePage - 1)}
              />
              {this.renderPageButtons(safePage, totalPages)}
              <IconButton
                iconProps={{ iconName: 'ChevronRight' }}
                title="Trang sau"
                ariaLabel="Trang sau"
                disabled={safePage >= totalPages}
                onClick={() => this.goToPage(safePage + 1)}
              />
            </Stack>
            <span className={styles.muted}>
              Hiển thị {pageStart + 1}–{Math.min(pageEnd, results.length)} / {results.length}
            </span>
          </Stack>
        )}

        <Dialog
          hidden={!editingDoc}
          onDismiss={this.cancelEdit}
          minWidth={520}
          maxWidth={720}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Sửa hashtag',
            subText: editingDoc ? `Cho: ${editingDoc.Name}` : ''
          }}
        >
          <Stack tokens={{ childrenGap: 10 }}>
            <TextField
              placeholder="Lọc hashtag..."
              iconProps={{ iconName: 'Filter' }}
              value={editFilter}
              onChange={(_, v) => this.setState({ editFilter: v || '' })}
              disabled={savingEdit}
            />
            <div style={{ maxHeight: 320, overflowY: 'auto' }}>
              {filteredEditTags.length === 0 && <span className={styles.muted}>Không tìm thấy hashtag.</span>}
              {groupHashtagsByCategory(filteredEditTags).map(g => (
                <div key={g.name} className={styles.groupRow} style={{ marginTop: 10 }}>
                  <GroupLabel name={g.name} />
                  <div className={styles.tagWrap}>
                    {g.items.map(t => (
                      <CatChip
                        key={t.Id}
                        tag={t}
                        small
                        selected={editingTagIds.indexOf(t.Id) !== -1}
                        onClick={() => !savingEdit && this.toggleEditTag(t.Id)}
                      />
                    ))}
                  </div>
                </div>
              ))}
            </div>
            <span className={styles.muted}>Đã chọn {editingTagIds.length}</span>
            {editError && <MessageBar messageBarType={MessageBarType.error}>{editError}</MessageBar>}
          </Stack>
          <DialogFooter>
            <PrimaryButton text={savingEdit ? 'Đang lưu...' : 'Lưu'} onClick={this.saveEdit} disabled={savingEdit} />
            <DefaultButton text="Huỷ" onClick={this.cancelEdit} disabled={savingEdit} />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={!deletingDoc}
          onDismiss={this.cancelDelete}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Xoá tài liệu?',
            subText: deletingDoc
              ? `"${deletingDoc.Name}" sẽ được chuyển vào thùng rác của site.`
              : ''
          }}
        >
          {deleteError && <MessageBar messageBarType={MessageBarType.error}>{deleteError}</MessageBar>}
          <DialogFooter>
            <PrimaryButton text={deleting ? 'Đang xoá...' : 'Xoá'} onClick={this.confirmDelete} disabled={deleting} />
            <DefaultButton text="Huỷ" onClick={this.cancelDelete} disabled={deleting} />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={!renamingDoc}
          onDismiss={this.cancelRename}
          minWidth={480}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Đổi tên tài liệu',
            subText: renamingDoc ? `Tên hiện tại: ${renamingDoc.Name}` : ''
          }}
        >
          <Stack tokens={{ childrenGap: 10 }}>
            <TextField
              label="Tên file mới"
              value={renameValue}
              onChange={(_, v) => this.setState({ renameValue: v || '' })}
              onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.confirmRename(); } }}
              disabled={renaming}
            />
            <span className={styles.muted}>Bao gồm phần mở rộng. Ký tự không hợp lệ: \ / : * ? &quot; &lt; &gt; |</span>
            {renameError && <MessageBar messageBarType={MessageBarType.error}>{renameError}</MessageBar>}
          </Stack>
          <DialogFooter>
            <PrimaryButton
              text={renaming ? 'Đang đổi...' : 'Đổi tên'}
              onClick={this.confirmRename}
              disabled={renaming || !renameValue.trim() || renameValue.trim() === renamingDoc?.Name}
            />
            <DefaultButton text="Huỷ" onClick={this.cancelRename} disabled={renaming} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
