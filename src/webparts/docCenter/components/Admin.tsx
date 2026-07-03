import * as React from 'react';
import {
  PrimaryButton, DefaultButton, MessageBar, MessageBarType,
  Dialog, DialogType, DialogFooter
} from '@fluentui/react';
import { IHashtag, groupHashtagsByCategory } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import { DocTagsEditor } from './DocTagsEditor';
import { Segmented, chipStyle, getCategoryPalette } from './ui';
import styles from './DocCenter.module.scss';

interface IProps {
  hashtagsListTitle: string;
  libraryTitle: string;
  hashtags: IHashtag[];
  siteUrl: string;
  onChange: () => Promise<void>;
}

interface IState {
  activeTab: 'tags' | 'edit';
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

const ADMIN_TABS = [
  { key: 'tags', label: '⌂  Quản lý hashtag' },
  { key: 'edit', label: '✎  Sửa tag tài liệu' }
];

export class Admin extends React.Component<IProps, IState> {

  public state: IState = {
    activeTab: 'tags',
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

  private renderCategoryPicker(value: string, onPick: (cat: string) => void, disabled: boolean): React.ReactElement {
    const existingCats = this.getExistingCategories();
    return (
      <div className={styles.groupPickRow}>
        <span className={styles.groupPickLabel}>Nhóm:</span>
        {existingCats.map(c => (
          <span
            key={c}
            style={chipStyle(getCategoryPalette(c), value === c)}
            onClick={() => !disabled && onPick(value === c ? '' : c)}
          >
            {c}
          </span>
        ))}
        <input
          className={styles.textInput}
          style={{ width: 180, height: 32, background: '#fff' }}
          placeholder="hoặc nhóm mới…"
          value={value}
          disabled={disabled}
          onChange={e => onPick(e.target.value)}
        />
      </div>
    );
  }

  private renderHashtagManager(): React.ReactElement {
    const { newTag, newTagDesc, newTagCat, filter, editingId, editingValue, editingDesc, editingCat, deletingId, status, busy } = this.state;
    const filtered = this.props.hashtags.filter(t => {
      if (!filter) return true;
      const q = filter.toLowerCase();
      return t.Title.toLowerCase().includes(q)
        || (t.Description || '').toLowerCase().includes(q)
        || (t.Category || '').toLowerCase().includes(q);
    });
    const deletingTag = this.props.hashtags.find(t => t.Id === deletingId);

    return (
      <div className={styles.section}>
        <div className={styles.warnBanner}>
          ⓘ Đổi tên hashtag sẽ tự cập nhật trên mọi tài liệu đang dùng. Xoá sẽ gỡ hashtag khỏi tất cả tài liệu.
        </div>

        <div className={styles.card} style={{ padding: '18px 20px' }}>
          <div className={styles.cardTitle}>Thêm hashtag mới</div>
          <div className={styles.addTagGrid}>
            <input
              className={styles.textInput}
              placeholder="# vd. hoa-don"
              value={newTag}
              disabled={busy}
              onChange={e => this.setState({ newTag: e.target.value })}
              onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.addTag(); } }}
            />
            <input
              className={styles.textInput}
              placeholder="Mô tả ngắn — hiển thị khi người dùng tìm kiếm (tuỳ chọn)"
              value={newTagDesc}
              disabled={busy}
              onChange={e => this.setState({ newTagDesc: e.target.value })}
            />
            <button
              type="button"
              className={styles.primaryBtn}
              onClick={() => void this.addTag()}
              disabled={busy || !newTag.trim()}
            >
              + Thêm
            </button>
          </div>
          {this.renderCategoryPicker(newTagCat, cat => this.setState({ newTagCat: cat }), busy)}
        </div>

        {status && (
          <MessageBar messageBarType={status.type} onDismiss={() => this.setState({ status: undefined })}>
            {status.text}
          </MessageBar>
        )}

        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div className={`${styles.searchBox} ${styles.searchBoxSmall}`} style={{ flex: 1, maxWidth: 420 }}>
            <span className={styles.searchIcon}>⌕</span>
            <input
              placeholder="Lọc hashtag…"
              value={filter}
              onChange={e => this.setState({ filter: e.target.value })}
            />
          </div>
          <span className={styles.muted}>{this.props.hashtags.length} hashtag</span>
        </div>

        {filtered.length === 0 && <span className={styles.muted}>Chưa có hashtag.</span>}

        {groupHashtagsByCategory(filtered).map(g => {
          const p = getCategoryPalette(g.name);
          return (
            <div key={g.name} className={styles.groupCard}>
              <div className={styles.groupCardHead} style={{ background: p.bg }}>
                <span className={styles.groupCardDot} style={{ background: p.solid }} />
                <span className={styles.groupCardTitle} style={{ color: p.text }}>{g.name}</span>
                <span className={styles.groupCardCount}>{g.items.length} hashtag</span>
              </div>
              {g.items.map(h => (
                editingId === h.Id ? (
                  <div key={h.Id} className={styles.tagRowEdit}>
                    <div className={styles.addTagGrid} style={{ marginBottom: 0 }}>
                      <input
                        className={styles.textInput}
                        placeholder="Tên hashtag"
                        value={editingValue}
                        disabled={busy}
                        onChange={e => this.setState({ editingValue: e.target.value })}
                        onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.saveEdit(); } }}
                      />
                      <input
                        className={styles.textInput}
                        placeholder="Mô tả ngắn (tuỳ chọn)"
                        value={editingDesc}
                        disabled={busy}
                        onChange={e => this.setState({ editingDesc: e.target.value })}
                      />
                      <span />
                    </div>
                    {this.renderCategoryPicker(editingCat, cat => this.setState({ editingCat: cat }), busy)}
                    <div style={{ display: 'flex', gap: 8 }}>
                      <button
                        type="button"
                        className={styles.solidBtn}
                        onClick={() => void this.saveEdit()}
                        disabled={busy || !editingValue.trim()}
                      >
                        Lưu thay đổi
                      </button>
                      <button type="button" className={styles.ghostBtn} onClick={this.cancelEdit} disabled={busy}>
                        Huỷ
                      </button>
                    </div>
                  </div>
                ) : (
                  <div key={h.Id} className={styles.tagRow}>
                    <span style={{ ...chipStyle(p, false), cursor: 'default', width: 'fit-content' }}>
                      #{h.Title}
                    </span>
                    <span className={h.Description ? styles.tagRowDesc : styles.tagRowDescEmpty}>
                      {h.Description || 'Chưa có mô tả — thêm để người dùng dễ tìm'}
                    </span>
                    <span className={styles.tagRowActions}>
                      <button
                        type="button"
                        className={`${styles.sqBtn} ${styles.sqBtnSm}`}
                        title="Sửa"
                        onClick={() => this.startEdit(h)}
                        disabled={busy}
                      >
                        ✎
                      </button>
                      <button
                        type="button"
                        className={`${styles.sqBtn} ${styles.sqBtnSm} ${styles.sqBtnDanger}`}
                        title="Xoá"
                        onClick={() => this.setState({ deletingId: h.Id })}
                        disabled={busy}
                      >
                        ✕
                      </button>
                    </span>
                  </div>
                )
              ))}
            </div>
          );
        })}

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
            <PrimaryButton onClick={this.confirmDelete} text="Xoá" disabled={this.state.busy} />
            <DefaultButton onClick={() => this.setState({ deletingId: undefined })} text="Huỷ" disabled={this.state.busy} />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  public render(): React.ReactElement {
    const { activeTab } = this.state;
    return (
      <div className={styles.section}>
        <Segmented
          options={ADMIN_TABS}
          value={activeTab}
          variant="admin"
          onChange={key => this.setState({ activeTab: key as 'tags' | 'edit' })}
        />
        {activeTab === 'tags' && this.renderHashtagManager()}
        {activeTab === 'edit' && (
          <DocTagsEditor
            libraryTitle={this.props.libraryTitle}
            hashtags={this.props.hashtags}
            siteUrl={this.props.siteUrl}
          />
        )}
      </div>
    );
  }
}
