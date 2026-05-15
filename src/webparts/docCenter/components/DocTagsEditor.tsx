import * as React from 'react';
import {
  Stack, TextField, PrimaryButton, DefaultButton,
  MessageBar, MessageBarType, Spinner, SpinnerSize, Icon, Link,
  Dialog, DialogType, DialogFooter
} from '@fluentui/react';
import {
  initializeFileTypeIcons, getFileTypeIconProps
} from '@fluentui/react-file-type-icons';
import { IHashtag, IDocument, groupHashtagsByCategory } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import styles from './DocCenter.module.scss';

initializeFileTypeIcons();

interface IProps {
  libraryTitle: string;
  hashtags: IHashtag[];
  siteUrl: string;
}

interface IState {
  loading: boolean;
  docs: IDocument[];
  nameQuery: string;
  editingDoc?: IDocument;
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
      await this.reload();
    } catch (e) {
      this.setState({
        savingEdit: false,
        editError: e instanceof Error ? e.message : String(e)
      });
    }
  };

  public render(): React.ReactElement {
    const {
      loading, docs, nameQuery, editingDoc, editingTagIds, editFilter, savingEdit, editError, loadError
    } = this.state;

    const filteredDocs = nameQuery
      ? docs.filter(d => d.Name.toLowerCase().includes(nameQuery.toLowerCase()))
      : docs;

    const filteredEditTags = this.props.hashtags.filter(t => {
      if (!editFilter) return true;
      const q = editFilter.toLowerCase();
      return t.Title.toLowerCase().includes(q)
        || (t.Description || '').toLowerCase().includes(q);
    });

    if (loading) {
      return <Spinner size={SpinnerSize.large} label="Đang tải tài liệu..." styles={{ root: { marginTop: 24 } }} />;
    }

    return (
      <Stack tokens={{ childrenGap: 16 }} className={styles.section}>
        <MessageBar messageBarType={MessageBarType.info}>
          Sửa hashtag cho bất kỳ tài liệu đã tải lên. Thay đổi được lưu tức thì.
        </MessageBar>

        <div className={styles.card}>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
            <Stack.Item grow>
              <TextField
                label="Tìm tài liệu"
                placeholder="Nhập một phần tên file..."
                iconProps={{ iconName: 'Search' }}
                value={nameQuery}
                onChange={(_, v) => this.setState({ nameQuery: v || '' })}
              />
            </Stack.Item>
            <DefaultButton text="Làm mới" iconProps={{ iconName: 'Refresh' }} onClick={this.reload} />
          </Stack>
        </div>

        {loadError && <MessageBar messageBarType={MessageBarType.error}>{loadError}</MessageBar>}

        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <span className={styles.countPill}>{filteredDocs.length} tài liệu</span>
        </Stack>

        {filteredDocs.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>Không tìm thấy tài liệu.</MessageBar>
        )}

        <div className={styles.docList}>
          {filteredDocs.map(d => (
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
                  {d.Hashtags.map(t => (
                    <span key={t.Id} className={`${styles.tagPill} ${styles.tagPillReadonly}`}>#{t.Title}</span>
                  ))}
                </div>
              </div>
              <PrimaryButton
                text="Sửa tag"
                iconProps={{ iconName: 'Edit' }}
                onClick={() => this.openEdit(d)}
              />
            </div>
          ))}
        </div>

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
                      return (
                        <span
                          key={t.Id}
                          className={`${styles.tagPill} ${selected ? styles.tagPillSelected : ''}`}
                          title={t.Description || ''}
                          onClick={() => !savingEdit && this.toggleEditTag(t.Id)}
                        >
                          #{t.Title}
                        </span>
                      );
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
      </Stack>
    );
  }
}
