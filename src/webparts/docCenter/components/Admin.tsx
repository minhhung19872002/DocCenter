import * as React from 'react';
import {
  Stack, TextField, PrimaryButton, DefaultButton, IconButton,
  MessageBar, MessageBarType, Dialog, DialogType, DialogFooter,
  Pivot, PivotItem, ComboBox, IComboBoxOption
} from '@fluentui/react';
import { IHashtag } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import { DocTagsEditor } from './DocTagsEditor';
import styles from './DocCenter.module.scss';

interface IProps {
  hashtagsListTitle: string;
  libraryTitle: string;
  hashtags: IHashtag[];
  siteUrl: string;
  onChange: () => Promise<void>;
}

interface IState {
  newTag: string;
  newTagDesc: string;
  newTagCat: string;
  filter: string;
  editingId?: number;
  editingValue: string;
  editingDesc: string;
  editingCat: string;
  deletingId?: number;
  status?: { type: MessageBarType; text: string };
  busy: boolean;
}

export class Admin extends React.Component<IProps, IState> {

  public state: IState = {
    newTag: '',
    newTagDesc: '',
    newTagCat: '',
    filter: '',
    editingValue: '',
    editingDesc: '',
    editingCat: '',
    busy: false
  };

  private withBusy = async (fn: () => Promise<void>, okMessage: string): Promise<void> => {
    this.setState({ busy: true, status: undefined });
    try {
      await fn();
      await this.props.onChange();
      this.setState({ busy: false, status: { type: MessageBarType.success, text: okMessage } });
    } catch (e) {
      this.setState({
        busy: false,
        status: { type: MessageBarType.error, text: e instanceof Error ? e.message : String(e) }
      });
    }
  };

  private addTag = (): Promise<void> => this.withBusy(async () => {
    await SharePointService.addHashtag(
      this.props.hashtagsListTitle,
      this.state.newTag,
      this.state.newTagDesc,
      this.state.newTagCat
    );
    this.setState({ newTag: '', newTagDesc: '', newTagCat: '' });
  }, 'Đã thêm hashtag.');

  private startEdit = (h: IHashtag): void =>
    this.setState({
      editingId: h.Id,
      editingValue: h.Title,
      editingDesc: h.Description || '',
      editingCat: h.Category || ''
    });

  private cancelEdit = (): void =>
    this.setState({ editingId: undefined, editingValue: '', editingDesc: '', editingCat: '' });

  private saveEdit = (): Promise<void> => this.withBusy(async () => {
    if (this.state.editingId == null) return;
    await SharePointService.updateHashtag(
      this.props.hashtagsListTitle,
      this.state.editingId,
      this.state.editingValue,
      this.state.editingDesc,
      this.state.editingCat
    );
    this.setState({ editingId: undefined, editingValue: '', editingDesc: '', editingCat: '' });
  }, 'Đã cập nhật hashtag. Tài liệu sẽ hiển thị tên mới.');

  private getExistingCategories(): string[] {
    const set = new Set<string>();
    for (const h of this.props.hashtags) {
      const c = (h.Category || '').trim();
      if (c) set.add(c);
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }

  private confirmDelete = (): Promise<void> => this.withBusy(async () => {
    if (this.state.deletingId == null) return;
    await SharePointService.deleteHashtag(this.props.hashtagsListTitle, this.state.deletingId);
    this.setState({ deletingId: undefined });
  }, 'Đã xoá hashtag.');

  private renderHashtagManager(): React.ReactElement {
    const { newTag, newTagDesc, newTagCat, filter, editingId, editingValue, editingDesc, editingCat, deletingId, status, busy } = this.state;
    const sorted = [...this.props.hashtags].sort((a, b) => a.Title.localeCompare(b.Title));
    const filtered = sorted.filter(t => {
      if (!filter) return true;
      const q = filter.toLowerCase();
      return t.Title.toLowerCase().includes(q)
        || (t.Description || '').toLowerCase().includes(q)
        || (t.Category || '').toLowerCase().includes(q);
    });
    const deletingTag = this.props.hashtags.find(t => t.Id === deletingId);
    const existingCats = this.getExistingCategories();
    const catOptions: IComboBoxOption[] = existingCats.map(c => ({ key: c, text: c }));
    const catPlaceholder = existingCats.length > 0
      ? 'Chọn từ danh sách hoặc nhập mới'
      : 'vd. Năm, Tháng, Công ty, Công việc';

    return (
      <Stack tokens={{ childrenGap: 16 }} className={styles.section}>
        <MessageBar messageBarType={MessageBarType.info}>
          Đổi tên hashtag sẽ tự cập nhật trên mọi tài liệu đang dùng. Xoá sẽ gỡ hashtag khỏi tất cả tài liệu.
        </MessageBar>

        <div className={styles.card}>
          <Stack tokens={{ childrenGap: 8 }}>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
              <Stack.Item grow>
                <TextField
                  label="Thêm hashtag"
                  placeholder="vd. hoa-don"
                  value={newTag}
                  onChange={(_, v) => this.setState({ newTag: v || '' })}
                  onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.addTag(); } }}
                  disabled={busy}
                />
              </Stack.Item>
              <PrimaryButton text="Thêm" iconProps={{ iconName: 'Add' }} onClick={this.addTag} disabled={busy || !newTag.trim()} />
            </Stack>
            <TextField
              label="Mô tả (tuỳ chọn)"
              placeholder="Ghi chú ngắn về hashtag — hiển thị khi người dùng tìm kiếm"
              multiline
              rows={2}
              value={newTagDesc}
              onChange={(_, v) => this.setState({ newTagDesc: v || '' })}
              disabled={busy}
            />
            <ComboBox
              label="Nhóm (tuỳ chọn)"
              placeholder={catPlaceholder}
              allowFreeform
              autoComplete="on"
              options={catOptions}
              text={newTagCat}
              onChange={(_, option, __, value) => {
                this.setState({ newTagCat: option ? option.text : (value ?? '') });
              }}
              disabled={busy}
            />
          </Stack>
        </div>

        {status && (
          <MessageBar messageBarType={status.type} onDismiss={() => this.setState({ status: undefined })}>
            {status.text}
          </MessageBar>
        )}

        <div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
            <span className={styles.sectionLabel}>Tất cả hashtag</span>
            <span className={styles.countPill}>{sorted.length}</span>
          </div>
          <TextField
            placeholder="Lọc..."
            iconProps={{ iconName: 'Filter' }}
            value={filter}
            onChange={(_, v) => this.setState({ filter: v || '' })}
          />
        </div>

        <Stack tokens={{ childrenGap: 8 }}>
          {filtered.length === 0 && <span className={styles.muted}>Chưa có hashtag.</span>}
          {filtered.map(h => (
            <div key={h.Id} className={styles.adminRow}>
              {editingId === h.Id ? (
                <Stack tokens={{ childrenGap: 8 }} grow>
                  <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                    <Stack.Item grow>
                      <TextField
                        label="Tên"
                        value={editingValue}
                        onChange={(_, v) => this.setState({ editingValue: v || '' })}
                        onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.saveEdit(); } }}
                        disabled={busy}
                      />
                    </Stack.Item>
                    <PrimaryButton text="Lưu" iconProps={{ iconName: 'Save' }} onClick={this.saveEdit} disabled={busy || !editingValue.trim()} />
                    <DefaultButton text="Huỷ" onClick={this.cancelEdit} disabled={busy} />
                  </Stack>
                  <TextField
                    label="Mô tả (tuỳ chọn)"
                    placeholder="Ghi chú ngắn về hashtag — hiển thị khi người dùng tìm kiếm"
                    multiline
                    rows={2}
                    value={editingDesc}
                    onChange={(_, v) => this.setState({ editingDesc: v || '' })}
                    disabled={busy}
                  />
                  <ComboBox
                    label="Nhóm (tuỳ chọn)"
                    placeholder={catPlaceholder}
                    allowFreeform
                    autoComplete="on"
                    options={catOptions}
                    text={editingCat}
                    onChange={(_, option, __, value) => {
                      this.setState({ editingCat: option ? option.text : (value ?? '') });
                    }}
                    disabled={busy}
                  />
                </Stack>
              ) : (
                <>
                  <Stack tokens={{ childrenGap: 2 }} grow>
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" style={{ alignSelf: 'flex-start' }}>
                      <span className={`${styles.tagPill} ${styles.tagPillReadonly} ${styles.tagPillBig}`}>#{h.Title}</span>
                      {h.Category && <span className={styles.muted}>· {h.Category}</span>}
                    </Stack>
                    {h.Description && <span className={styles.muted}>{h.Description}</span>}
                  </Stack>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Sửa" onClick={() => this.startEdit(h)} />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Xoá" onClick={() => this.setState({ deletingId: h.Id })} />
                  </Stack>
                </>
              )}
            </div>
          ))}
        </Stack>

        <Dialog
          hidden={deletingId == null}
          onDismiss={() => this.setState({ deletingId: undefined })}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Xoá hashtag?',
            subText: deletingTag
              ? `"#${deletingTag.Title}" sẽ bị gỡ khỏi tất cả tài liệu đang dùng. Không thể hoàn tác.`
              : ''
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.confirmDelete} text="Xoá" disabled={busy} />
            <DefaultButton onClick={() => this.setState({ deletingId: undefined })} text="Huỷ" disabled={busy} />
          </DialogFooter>
        </Dialog>
      </Stack>
    );
  }

  public render(): React.ReactElement {
    return (
      <Pivot>
        <PivotItem headerText="Quản lý hashtag" itemIcon="Tag">
          {this.renderHashtagManager()}
        </PivotItem>
        <PivotItem headerText="Sửa tag tài liệu" itemIcon="EditNote">
          <DocTagsEditor
            libraryTitle={this.props.libraryTitle}
            hashtags={this.props.hashtags}
            siteUrl={this.props.siteUrl}
          />
        </PivotItem>
      </Pivot>
    );
  }
}
