import * as React from 'react';
import {
  Stack, TextField, ChoiceGroup, IChoiceGroupOption,
  DefaultButton, PrimaryButton, IconButton, MessageBar, MessageBarType,
  Spinner, SpinnerSize, Link, Icon, Dialog, DialogType, DialogFooter,
  TooltipHost, DirectionalHint, ITooltipHostStyles
} from '@fluentui/react';
import {
  initializeFileTypeIcons, getFileTypeIconProps
} from '@fluentui/react-file-type-icons';
import { IHashtag, IDocument, SearchMode, groupHashtagsByCategory } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import styles from './DocCenter.module.scss';

initializeFileTypeIcons();

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
  filter: string;
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

const MODE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'all', text: 'Tất cả (VÀ)' },
  { key: 'any', text: 'Bất kỳ (HOẶC)' }
];

const TOOLTIP_HOST_STYLES: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

export class Search extends React.Component<IProps, IState> {

  public state: IState = {
    selectedTagIds: [],
    mode: 'all',
    nameQuery: '',
    filter: '',
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

  private onSearchKeyDown = (e: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
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

  private clearTags = (): void => this.setState({ selectedTagIds: [] }, () => void this.runSearch());

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

  private getExtension(name: string): string {
    const dot = name.lastIndexOf('.');
    return dot >= 0 ? name.substring(dot + 1).toLowerCase() : '';
  }

  private buildAbsoluteUrl(serverRelativeUrl: string): string {
    const origin = this.props.siteUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
    const absolute = serverRelativeUrl.startsWith('/') ? `${origin}${serverRelativeUrl}` : serverRelativeUrl;
    return absolute + (absolute.indexOf('?') === -1 ? '?web=1' : '&web=1');
  }

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

  public render(): React.ReactElement {
    const {
      selectedTagIds, mode, nameQuery, filter, searching, results, searched,
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
    const matchTag = (t: IHashtag, q: string): boolean => {
      if (!q) return true;
      const needle = q.toLowerCase();
      return t.Title.toLowerCase().includes(needle)
        || (t.Description || '').toLowerCase().includes(needle);
    };
    const filteredTags = this.props.hashtags.filter(t => matchTag(t, filter));
    const filteredEditTags = this.props.hashtags.filter(t => matchTag(t, editFilter));

    return (
      <Stack tokens={{ childrenGap: 16 }} className={styles.section}>
        <div className={styles.card}>
          <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="start">
            <Stack.Item grow>
              <div className={styles.searchAutocomplete}>
                <TextField
                  label="Tìm theo tên file hoặc hashtag"
                  placeholder="Nhập một phần tên file hoặc hashtag..."
                  iconProps={{ iconName: 'Search' }}
                  value={nameQuery}
                  autoComplete="off"
                  onChange={(_, v) => this.setState(
                    { nameQuery: v || '', showSuggestions: true, highlightedSuggestion: -1 },
                    this.scheduleSearch
                  )}
                  onFocus={this.onSearchFocus}
                  onBlur={this.onSearchBlur}
                  onKeyDown={this.onSearchKeyDown}
                />
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
                        <Icon {...getFileTypeIconProps({ extension: this.getExtension(d.Name), size: 16 })} />
                        <span className={styles.suggestName}>{d.Name}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </Stack.Item>
            <Stack.Item>
              <ChoiceGroup
                label="Kiểu khớp"
                selectedKey={mode}
                options={MODE_OPTIONS}
                onChange={(_, opt) => opt && this.setState({ mode: opt.key as SearchMode }, () => void this.runSearch())}
              />
            </Stack.Item>
          </Stack>

          <div style={{ marginTop: 16 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
              <span className={styles.sectionLabel}>Hashtag</span>
              {selectedTagIds.length > 0 && (
                <span className={styles.clearLink} onClick={this.clearTags}>
                  Bỏ chọn {selectedTagIds.length}
                </span>
              )}
            </div>
            <TextField
              placeholder="Lọc hashtag..."
              iconProps={{ iconName: 'Filter' }}
              value={filter}
              onChange={(_, v) => this.setState({ filter: v || '' })}
            />
            {filteredTags.length === 0 && <div className={styles.tagWrap}><span className={styles.muted}>Không tìm thấy hashtag.</span></div>}
            {groupHashtagsByCategory(filteredTags).map(g => (
              <div key={g.name} style={{ marginTop: 6 }}>
                <div className={styles.muted} style={{ fontWeight: 600, marginBottom: 4 }}>{g.name}</div>
                <div className={styles.tagWrap} style={{ marginTop: 0 }}>
                  {g.items.map(t => {
                    const selected = selectedTagIds.indexOf(t.Id) !== -1;
                    const pill = (
                      <span
                        className={`${styles.tagPill} ${styles.tagPillBig} ${selected ? styles.tagPillSelected : ''}`}
                        onClick={() => this.toggleTag(t.Id)}
                      >
                        #{t.Title}
                      </span>
                    );
                    return t.Description
                      ? (
                        <TooltipHost
                          key={t.Id}
                          content={t.Description}
                          directionalHint={DirectionalHint.topCenter}
                          styles={TOOLTIP_HOST_STYLES}
                        >
                          {pill}
                        </TooltipHost>
                      )
                      : <React.Fragment key={t.Id}>{pill}</React.Fragment>;
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>

        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <DefaultButton text="Làm mới" iconProps={{ iconName: 'Refresh' }} onClick={this.runSearch} disabled={searching} />
          {!searching && searched && <span className={styles.countPill}>{results.length} tài liệu</span>}
        </Stack>

        {searching && <Spinner size={SpinnerSize.medium} label="Đang tìm..." />}

        {!searching && searched && results.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>Không có tài liệu nào phù hợp.</MessageBar>
        )}

        {!searching && results.length > 0 && (
          <div className={styles.docList} ref={this.resultsTopRef}>
            {pagedResults.map(d => (
              <div key={d.Id} className={styles.docRow}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div className={styles.docTitle}>
                    <Icon {...getFileTypeIconProps({ extension: this.getExtension(d.Name), size: 20 })} />
                    <Link
                      href={this.buildAbsoluteUrl(d.ServerRelativeUrl)}
                      target="_blank"
                      rel="noopener noreferrer"
                      data-interception="off"
                      onClick={e => this.openDocument(e, d)}
                    >
                      {d.Name}
                    </Link>
                  </div>
                  <div className={styles.docMeta}>
                    {d.SizeKB ? `${d.SizeKB} KB · ` : ''}
                    {new Date(d.Modified).toLocaleString()}{d.CreatedBy ? ` · ${d.CreatedBy}` : ''}
                  </div>
                  <div className={styles.tagWrap}>
                    {d.Hashtags.length === 0 && <span className={styles.muted}>Không có hashtag</span>}
                    {d.Hashtags.map(t => {
                      const active = selectedTagIds.indexOf(t.Id) !== -1;
                      return (
                        <span
                          key={t.Id}
                          className={`${styles.tagPill} ${active ? styles.tagPillSelected : ''}`}
                          title={active ? 'Nhấn để bỏ khỏi bộ lọc' : 'Nhấn để thêm vào bộ lọc'}
                          onClick={() => this.toggleTag(t.Id)}
                        >
                          #{t.Title}
                        </span>
                      );
                    })}
                  </div>
                </div>
                {this.props.isAdmin && (
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <IconButton
                      iconProps={{ iconName: 'Rename' }}
                      title="Đổi tên tài liệu"
                      ariaLabel="Đổi tên tài liệu"
                      onClick={() => this.openRename(d)}
                    />
                    {d.CanDelete && (
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Xoá tài liệu"
                        ariaLabel="Xoá tài liệu"
                        onClick={() => this.openDelete(d)}
                      />
                    )}
                  </Stack>
                )}
                {!this.props.isAdmin && this.isWithinEditWindow(d) && (
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <IconButton
                      iconProps={{ iconName: 'Rename' }}
                      title="Đổi tên tài liệu"
                      ariaLabel="Đổi tên tài liệu"
                      onClick={() => this.openRename(d)}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Edit' }}
                      title="Sửa hashtag"
                      ariaLabel="Sửa hashtag"
                      onClick={() => this.openEdit(d)}
                    />
                    {d.CanDelete && (
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Xoá tài liệu"
                        ariaLabel="Xoá tài liệu"
                        onClick={() => this.openDelete(d)}
                      />
                    )}
                  </Stack>
                )}
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
                <div key={g.name} style={{ marginTop: 6 }}>
                  <div className={styles.muted} style={{ fontWeight: 600, marginBottom: 4 }}>{g.name}</div>
                  <div className={styles.tagWrap} style={{ marginTop: 0 }}>
                    {g.items.map(t => {
                      const selected = editingTagIds.indexOf(t.Id) !== -1;
                      const pill = (
                        <span
                          className={`${styles.tagPill} ${selected ? styles.tagPillSelected : ''}`}
                          onClick={() => !savingEdit && this.toggleEditTag(t.Id)}
                        >
                          #{t.Title}
                        </span>
                      );
                      return t.Description
                        ? (
                          <TooltipHost
                            key={t.Id}
                            content={t.Description}
                            directionalHint={DirectionalHint.topCenter}
                            styles={TOOLTIP_HOST_STYLES}
                          >
                            {pill}
                          </TooltipHost>
                        )
                        : <React.Fragment key={t.Id}>{pill}</React.Fragment>;
                    })}
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
      </Stack>
    );
  }
}
