import * as React from 'react';
import { ProgressIndicator } from '@fluentui/react';
import { IHashtag, IUploadResult, groupHashtagsByCategory } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import { CatChip, ExtBadge, GroupLabel } from './ui';
import styles from './DocCenter.module.scss';

interface IProps {
  libraryTitle: string;
  hashtagsListTitle: string;
  hashtags: IHashtag[];
  isAdmin: boolean;
  onHashtagsChanged: () => Promise<void>;
  onGoToSearch?: () => void;
}

interface IState {
  files: File[];
  selectedTagIds: number[];
  dragActive: boolean;
  uploading: boolean;
  results: IUploadResult[];
  quickTag: string;
}

export class Upload extends React.Component<IProps, IState> {

  private fileInputRef = React.createRef<HTMLInputElement>();

  public state: IState = {
    files: [],
    selectedTagIds: [],
    dragActive: false,
    uploading: false,
    results: [],
    quickTag: ''
  };

  private onPickClick = (): void => this.fileInputRef.current?.click();

  private onFileInput = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const list = e.target.files;
    if (!list) return;
    const next: File[] = [];
    for (let i = 0; i < list.length; i++) next.push(list[i]);
    this.setState({ files: [...this.state.files, ...next], results: [] });
    e.target.value = '';
  };

  private onDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    this.setState({ dragActive: false });
    if (!e.dataTransfer.files) return;
    const next: File[] = [];
    for (let i = 0; i < e.dataTransfer.files.length; i++) next.push(e.dataTransfer.files[i]);
    this.setState({ files: [...this.state.files, ...next], results: [] });
  };

  private toggleTag = (id: number): void => {
    const set = new Set(this.state.selectedTagIds);
    if (set.has(id)) set.delete(id); else set.add(id);
    this.setState({ selectedTagIds: Array.from(set) });
  };

  private removeFile = (idx: number): void => {
    this.setState({ files: this.state.files.filter((_, i) => i !== idx) });
  };

  private addQuickTag = async (): Promise<void> => {
    const name = this.state.quickTag.trim();
    if (!name) return;
    const created = await SharePointService.addHashtag(this.props.hashtagsListTitle, name, '');
    await this.props.onHashtagsChanged();
    this.setState(s => ({
      quickTag: '',
      selectedTagIds: s.selectedTagIds.indexOf(created.Id) === -1 ? [...s.selectedTagIds, created.Id] : s.selectedTagIds
    }));
  };

  private upload = async (): Promise<void> => {
    if (this.state.files.length === 0) return;
    this.setState({ uploading: true, results: [] });
    const out: IUploadResult[] = [];
    for (const f of this.state.files) {
      const r = await SharePointService.uploadDocument(this.props.libraryTitle, f, this.state.selectedTagIds);
      out.push(r);
    }
    this.setState({ uploading: false, results: out, files: [], selectedTagIds: [] });
  };

  private formatSize(bytes: number): string {
    if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
    return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  }

  public render(): React.ReactElement {
    const { files, selectedTagIds, dragActive, uploading, results, quickTag } = this.state;
    const canUpload = !uploading && files.length > 0;
    const okCount = results.filter(r => r.success).length;
    const failed = results.filter(r => !r.success);
    const selectedTags = this.props.hashtags.filter(t => selectedTagIds.indexOf(t.Id) !== -1);
    const fileCountLabel = files.length > 0 ? `${files.length} file` : 'Chưa có file';

    return (
      <div>
        {okCount > 0 && (
          <div className={styles.successBanner}>
            ✓ Đã tải lên {okCount} file.
            {this.props.onGoToSearch && (
              <span className={styles.successBannerLink} onClick={this.props.onGoToSearch}>
                Xem trong Tìm kiếm →
              </span>
            )}
          </div>
        )}
        {failed.map((r, i) => (
          <div key={i} className={styles.errorBanner}>
            ✕ {r.fileName}: {r.error}
          </div>
        ))}

        <div className={styles.uploadGrid}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div className={styles.card}>
              <div className={styles.stepHeader}>
                <span className={styles.stepNumber}>1</span>
                <span className={styles.stepLabel}>Chọn file</span>
                <span className={styles.stepMetaAccent}>{fileCountLabel}</span>
              </div>
              <div
                className={`${styles.dropZone} ${dragActive ? styles.dropZoneActive : ''}`}
                onClick={this.onPickClick}
                onDragOver={e => { e.preventDefault(); this.setState({ dragActive: true }); }}
                onDragLeave={() => this.setState({ dragActive: false })}
                onDrop={this.onDrop}
              >
                <div className={styles.dropZoneTitle}>
                  Kéo thả file vào đây, hoặc <span>chọn từ máy tính</span>
                </div>
                <div className={styles.dropZoneHint}>PDF, Word, Excel, PowerPoint · nhiều file cùng lúc</div>
                <input
                  type="file"
                  multiple
                  ref={this.fileInputRef}
                  style={{ display: 'none' }}
                  onClick={e => e.stopPropagation()}
                  onChange={this.onFileInput}
                />
              </div>

              {files.length > 0 && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 12 }}>
                  {files.map((f, i) => (
                    <div key={i} className={styles.fileRow}>
                      <ExtBadge fileName={f.name} size={32} />
                      <span className={styles.fileRowName}>{f.name}</span>
                      <span className={styles.fileRowSize}>{this.formatSize(f.size)}</span>
                      <span className={styles.fileRowRemove} title="Bỏ file" onClick={() => this.removeFile(i)}>✕</span>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div className={styles.card}>
              <div className={styles.stepHeader}>
                <span className={styles.stepNumber}>2</span>
                <span className={styles.stepLabel}>Gắn hashtag</span>
                <span className={styles.stepMeta}>Đã chọn <b>{selectedTagIds.length}</b></span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                {this.props.hashtags.length === 0 && (
                  <span className={styles.muted}>Chưa có hashtag nào. Thêm mới ở dưới.</span>
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
                <div style={{ display: 'flex', gap: 10 }}>
                  <input
                    className={styles.textInput}
                    style={{ flex: 1 }}
                    placeholder="Thêm hashtag mới — vd. hoa-don"
                    value={quickTag}
                    onChange={e => this.setState({ quickTag: e.target.value })}
                    onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.addQuickTag(); } }}
                  />
                  <button
                    type="button"
                    className={styles.primaryBtn}
                    style={{ width: 110 }}
                    onClick={() => void this.addQuickTag()}
                    disabled={!quickTag.trim()}
                  >
                    + Thêm
                  </button>
                </div>
              </div>
            </div>
          </div>

          <div className={styles.uploadSide}>
            <div className={styles.card}>
              <div className={styles.stepHeader}>
                <span className={styles.stepNumber}>3</span>
                <span className={styles.stepLabel}>Xác nhận &amp; tải lên</span>
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
                <div className={styles.summaryRow}>
                  <span>File</span>
                  <b>{fileCountLabel}</b>
                </div>
                <div className={styles.summaryRow}>
                  <span>Thư viện đích</span>
                  <b>{this.props.libraryTitle}</b>
                </div>
                <div className={styles.summaryRow}>
                  <span>Hashtag</span>
                  <span className={styles.summaryChips}>
                    {selectedTags.length === 0 && <b>—</b>}
                    {selectedTags.map(t => (
                      <CatChip key={t.Id} tag={t} small readonly />
                    ))}
                  </span>
                </div>
              </div>
              <div className={styles.divider} />
              <button
                type="button"
                className={styles.uploadBtn}
                onClick={() => void this.upload()}
                disabled={!canUpload}
              >
                {uploading ? 'Đang tải lên…' : '↑ Tải lên'}
              </button>
              {uploading && (
                <div style={{ marginTop: 10 }}>
                  <ProgressIndicator label="Đang tải file..." />
                </div>
              )}
              <div className={styles.uploadNote}>Cần ít nhất 1 file · hashtag không bắt buộc</div>
            </div>

            <div className={styles.tipBox}>
              <b>Mẹo:</b> chọn ít nhất 1 hashtag <b>Loại tài liệu</b> và 1 hashtag <b>Phòng ban</b> để tài liệu dễ tìm về sau.
            </div>
          </div>
        </div>
      </div>
    );
  }
}
