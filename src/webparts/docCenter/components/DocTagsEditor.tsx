import * as React from 'react';
import { Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import { IHashtag, IDocument } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import { CatChip, ExtBadge } from './ui';
import styles from './DocCenter.module.scss';

const MAX_QUICK_ADD = 14;

interface IProps {
  libraryTitle: string;
  hashtags: IHashtag[];
  siteUrl: string;
}

interface IState {
  loading: boolean;
  docs: IDocument[];
  nameQuery: string;
  editingId?: number;
  editingTagIds: number[];
  editFilter: string;
  savingEdit: boolean;
  editError?: string;
  loadError?: string;
}

export class DocTagsEditor extends React.Component<IProps, IState> {

  public state: IState = {
    loading: true,
    docs: [],
    nameQuery: '',
    editingTagIds: [],
    editFilter: '',
    savingEdit: false
  };

  public async componentDidMount(): Promise<void> {
    await this.reload();
  }

  private reload = async (): Promise<void> => {
    this.setState({ loading: true, loadError: undefined });
    try {
      const docs = await SharePointService.searchDocuments(this.props.libraryTitle, [], 'any', '');
      this.setState({ docs, loading: false });
    } catch (e) {
      this.setState({ loading: false, loadError: e instanceof Error ? e.message : String(e) });
    }
  };

  private buildAbsoluteUrl(serverRelativeUrl: string): string {
    const origin = this.props.siteUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
    const absolute = serverRelativeUrl.startsWith('/') ? `${origin}${serverRelativeUrl}` : serverRelativeUrl;
    return absolute + (absolute.indexOf('?') === -1 ? '?web=1' : '&web=1');
  }

  private openEdit = (doc: IDocument): void => {
    this.setState({
      editingId: doc.Id,
      editingTagIds: doc.Hashtags.map(t => t.Id),
      editFilter: '',
      savingEdit: false,
      editError: undefined
    });
  };

  private cancelEdit = (): void => {
    if (this.state.savingEdit) return;
    this.setState({ editingId: undefined, editingTagIds: [], editError: undefined });
  };

  private removeEditTag = (id: number): void => {
    this.setState({ editingTagIds: this.state.editingTagIds.filter(x => x !== id) });
  };

  private addEditTag = (id: number): void => {
    if (this.state.editingTagIds.indexOf(id) !== -1) return;
    this.setState({ editingTagIds: [...this.state.editingTagIds, id] });
  };

  private saveEdit = async (): Promise<void> => {
    const { editingId, editingTagIds } = this.state;
    if (editingId == null) return;
    this.setState({ savingEdit: true, editError: undefined });
    try {
      await SharePointService.updateDocumentHashtags(this.props.libraryTitle, editingId, editingTagIds);
      this.setState({ editingId: undefined, editingTagIds: [], savingEdit: false });
      await this.reload();
    } catch (e) {
      this.setState({
        savingEdit: false,
        editError: e instanceof Error ? e.message : String(e)
      });
    }
  };

  private renderEditPanel(): React.ReactElement {
    const { editingTagIds, editFilter, savingEdit, editError } = this.state;
    const currentTags = this.props.hashtags.filter(t => editingTagIds.indexOf(t.Id) !== -1);
    const q = editFilter.trim().toLowerCase();
    const addable = this.props.hashtags
      .filter(t => editingTagIds.indexOf(t.Id) === -1)
      .filter(t => !q
        || t.Title.toLowerCase().includes(q)
        || (t.Description || '').toLowerCase().includes(q))
      .slice(0, MAX_QUICK_ADD);

    return (
      <div className={styles.editPanel}>
        <div className={styles.editPanelRow}>
          <span className={styles.editPanelLabel}>Tag hiện tại</span>
          {currentTags.length === 0 && <span className={styles.muted}>Chưa gắn hashtag nào</span>}
          {currentTags.map(t => (
            <CatChip
              key={t.Id}
              tag={t}
              small
              selected
              showRemove
              onClick={() => !savingEdit && this.removeEditTag(t.Id)}
            />
          ))}
        </div>
        <div className={styles.editPanelRow} style={{ alignItems: 'flex-start' }}>
          <span className={styles.editPanelLabel} style={{ marginTop: 5 }}>Thêm nhanh</span>
          <div className={styles.tagWrap}>
            {addable.map(t => (
              <CatChip
                key={t.Id}
                tag={t}
                small
                prefix="+ "
                onClick={() => !savingEdit && this.addEditTag(t.Id)}
              />
            ))}
            <input
              className={styles.textInput}
              style={{ width: 170, height: 28, background: '#fff', fontSize: 11.5 }}
              placeholder="lọc hashtag…"
              value={editFilter}
              disabled={savingEdit}
              onChange={e => this.setState({ editFilter: e.target.value })}
            />
          </div>
        </div>
        {editError && <MessageBar messageBarType={MessageBarType.error}>{editError}</MessageBar>}
        <div style={{ display: 'flex', gap: 8, marginTop: 2 }}>
          <button type="button" className={styles.solidBtn} onClick={() => void this.saveEdit()} disabled={savingEdit}>
            {savingEdit ? 'Đang lưu…' : 'Lưu thay đổi'}
          </button>
          <button type="button" className={styles.ghostBtn} onClick={this.cancelEdit} disabled={savingEdit}>
            Huỷ
          </button>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement {
    const { loading, docs, nameQuery, editingId, loadError } = this.state;

    const filteredDocs = nameQuery
      ? docs.filter(d => d.Name.toLowerCase().includes(nameQuery.toLowerCase()))
      : docs;

    if (loading) {
      return <Spinner size={SpinnerSize.large} label="Đang tải tài liệu..." styles={{ root: { marginTop: 24 } }} />;
    }

    return (
      <div className={styles.section}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div className={`${styles.searchBox} ${styles.searchBoxSmall}`} style={{ flex: 1, maxWidth: 420 }}>
            <span className={styles.searchIcon}>⌕</span>
            <input
              placeholder="Nhập một phần tên file…"
              value={nameQuery}
              onChange={e => this.setState({ nameQuery: e.target.value })}
            />
          </div>
          <span className={styles.muted}>{filteredDocs.length} tài liệu</span>
          <button type="button" className={styles.sqBtn} title="Làm mới" onClick={() => void this.reload()}>
            ⟳
          </button>
        </div>

        {loadError && <MessageBar messageBarType={MessageBarType.error}>{loadError}</MessageBar>}

        {filteredDocs.length === 0 && (
          <div className={styles.emptyBox}>Không tìm thấy tài liệu.</div>
        )}

        {filteredDocs.length > 0 && (
          <div className={styles.groupCard}>
            {filteredDocs.map(d => {
              const isEditing = editingId === d.Id;
              return (
                <div key={d.Id} className={`${styles.editRow} ${isEditing ? styles.editRowActive : ''}`}>
                  <div className={styles.editRowMain}>
                    <ExtBadge fileName={d.Name} size={34} />
                    <span style={{ flex: 1, minWidth: 0 }}>
                      <span className={styles.docTitle} style={{ fontSize: 13 }}>
                        <a
                          href={this.buildAbsoluteUrl(d.ServerRelativeUrl)}
                          target="_blank"
                          rel="noopener noreferrer"
                          data-interception="off"
                        >
                          {d.Name}
                        </a>
                      </span>
                      <span className={styles.docMeta} style={{ fontSize: 11 }}>
                        {d.SizeKB ? `${d.SizeKB} KB · ` : ''}
                        {new Date(d.Modified).toLocaleString()}{d.CreatedBy ? ` · ${d.CreatedBy}` : ''}
                      </span>
                    </span>
                    {!isEditing && (
                      <>
                        <span className={styles.docChips} style={{ maxWidth: 320 }}>
                          {d.Hashtags.map(t => (
                            <CatChip key={t.Id} tag={t} small readonly />
                          ))}
                        </span>
                        <button type="button" className={styles.editTagBtn} onClick={() => this.openEdit(d)}>
                          ✎ Sửa tag
                        </button>
                      </>
                    )}
                    {isEditing && <span className={styles.editBadge}>ĐANG SỬA</span>}
                  </div>
                  {isEditing && this.renderEditPanel()}
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }
}
