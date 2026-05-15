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

const EDIT_WINDOW_MS = 24 * 60 * 60 * 1000;

const MODE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'any', text: 'Bất kỳ (HOẶC)' },
  { key: 'all', text: 'Tất cả (VÀ)' }
];

const TOOLTIP_HOST_STYLES: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

export class Search extends React.Component<IProps, IState> {

  public state: IState = {
    selectedTagIds: [],
    mode: 'any',
    nameQuery: '',
    filter: '',
    searching: false,
    results: [],
    searched: false,
    editingTagIds: [],
    editFilter: '',
    savingEdit: false,
    deleting: false,
    renameValue: '',
    renaming: false
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
      this.setState({ results, searching: false, searched: true });
    } catch (e) {
      this.setState({ searching: false, searched: true, results: [] });
      // eslint-disable-next-line no-console
      console.error(e);
    }
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

  public render(): React.ReactElement {
    const {
      selectedTagIds, mode, nameQuery, filter, searching, results, searched,
      editingDoc, editingTagIds, editFilter, savingEdit, editError,
      deletingDoc, deleting, deleteError,
      renamingDoc, renameValue, renaming, renameError
    } = this.state;
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
              <TextField
                label="Tìm theo tên file hoặc hashtag"
                placeholder="Nhập một phần tên file hoặc hashtag..."
                iconProps={{ iconName: 'Search' }}
                value={nameQuery}
                onChange={(_, v) => this.setState({ nameQuery: v || '' }, () => void this.runSearch())}
              />
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
          <div className={styles.docList}>
            {results.map(d => (
              <div key={d.Id} className={styles.docRow}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div className={styles.docTitle}>
                    <Icon {...getFileTypeIconProps({ extension: this.getExtension(d.Name), size: 20 })} />
                    <Link href={this.buildAbsoluteUrl(d.ServerRelativeUrl)} target="_blank">
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
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Xoá tài liệu"
                      ariaLabel="Xoá tài liệu"
                      onClick={() => this.openDelete(d)}
                    />
                  </Stack>
                )}
              </div>
            ))}
          </div>
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
