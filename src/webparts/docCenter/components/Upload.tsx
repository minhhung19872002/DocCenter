import * as React from 'react';
import {
  PrimaryButton, DefaultButton, Stack, MessageBar, MessageBarType,
  ProgressIndicator, TextField, Icon
} from '@fluentui/react';
import {
  initializeFileTypeIcons, getFileTypeIconProps
} from '@fluentui/react-file-type-icons';
import { IHashtag, IUploadResult } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import styles from './DocCenter.module.scss';

initializeFileTypeIcons();

interface IProps {
  libraryTitle: string;
  hashtagsListTitle: string;
  hashtags: IHashtag[];
  isAdmin: boolean;
  onHashtagsChanged: () => Promise<void>;
}

interface IState {
  files: File[];
  selectedTagIds: number[];
  dragActive: boolean;
  uploading: boolean;
  results: IUploadResult[];
  quickTag: string;
  filter: string;
}

export class Upload extends React.Component<IProps, IState> {

  private fileInputRef = React.createRef<HTMLInputElement>();

  public state: IState = {
    files: [],
    selectedTagIds: [],
    dragActive: false,
    uploading: false,
    results: [],
    quickTag: '',
    filter: ''
  };

  private onPickClick = (): void => this.fileInputRef.current?.click();

  private onFileInput = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const list = e.target.files;
    if (!list) return;
    const next: File[] = [];
    for (let i = 0; i < list.length; i++) next.push(list[i]);
    this.setState({ files: [...this.state.files, ...next] });
    e.target.value = '';
  };

  private onDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    this.setState({ dragActive: false });
    if (!e.dataTransfer.files) return;
    const next: File[] = [];
    for (let i = 0; i < e.dataTransfer.files.length; i++) next.push(e.dataTransfer.files[i]);
    this.setState({ files: [...this.state.files, ...next] });
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
    const created = await SharePointService.addHashtag(this.props.hashtagsListTitle, name);
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

  private renderStep(num: number, label: string): React.ReactElement {
    return (
      <div className={styles.stepHeader}>
        <span className={styles.stepNumber}>{num}</span>
        <span className={styles.stepLabel}>{label}</span>
      </div>
    );
  }

  public render(): React.ReactElement {
    const { files, selectedTagIds, dragActive, uploading, results, quickTag, filter } = this.state;
    const filteredTags = this.props.hashtags.filter(
      t => !filter || t.Title.toLowerCase().includes(filter.toLowerCase())
    );

    return (
      <Stack tokens={{ childrenGap: 18 }} className={styles.section}>
        <div>
          {this.renderStep(1, 'Choose files')}
          <div
            className={`${styles.dropZone} ${dragActive ? styles.dropZoneActive : ''}`}
            onDragOver={e => { e.preventDefault(); this.setState({ dragActive: true }); }}
            onDragLeave={() => this.setState({ dragActive: false })}
            onDrop={this.onDrop}
          >
            <Icon iconName="CloudUpload" className={styles.dropZoneIcon} style={{ fontSize: 32 }} />
            <div className={styles.dropZoneHint}>Drag files here, or browse from your computer</div>
            <PrimaryButton text="Browse files" iconProps={{ iconName: 'OpenFile' }} onClick={this.onPickClick} />
            <input
              type="file"
              multiple
              ref={this.fileInputRef}
              style={{ display: 'none' }}
              onChange={this.onFileInput}
            />
          </div>

          {files.length > 0 && (
            <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 8 }}>
              <span className={styles.muted}>{files.length} file(s) ready</span>
              {files.map((f, i) => {
                const dot = f.name.lastIndexOf('.');
                const ext = dot >= 0 ? f.name.substring(dot + 1).toLowerCase() : '';
                return (
                  <div key={i} className={styles.fileChip}>
                    <span className={styles.fileChipName}>
                      <Icon {...getFileTypeIconProps({ extension: ext, size: 20 })} />
                      <span>{f.name} <span className={styles.muted}>· {Math.round(f.size / 1024)} KB</span></span>
                    </span>
                    <DefaultButton iconProps={{ iconName: 'Cancel' }} text="Remove" onClick={() => this.removeFile(i)} />
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div>
          {this.renderStep(2, 'Pick hashtags')}
          <div className={styles.card}>
            <TextField
              placeholder="Filter hashtags..."
              iconProps={{ iconName: 'Filter' }}
              value={filter}
              onChange={(_, v) => this.setState({ filter: v || '' })}
            />
            <div className={styles.tagWrap}>
              {filteredTags.length === 0 && <span className={styles.muted}>No hashtags yet. Add one below.</span>}
              {filteredTags.map(t => {
                const selected = selectedTagIds.indexOf(t.Id) !== -1;
                return (
                  <span
                    key={t.Id}
                    className={`${styles.tagPill} ${selected ? styles.tagPillSelected : ''}`}
                    onClick={() => this.toggleTag(t.Id)}
                  >
                    #{t.Title}
                  </span>
                );
              })}
            </div>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end" style={{ marginTop: 12 }}>
              <Stack.Item grow>
                <TextField
                  label="Add a new hashtag"
                  placeholder="e.g. quarterly-report"
                  value={quickTag}
                  onChange={(_, v) => this.setState({ quickTag: v || '' })}
                  onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.addQuickTag(); } }}
                />
              </Stack.Item>
              <DefaultButton text="Add" iconProps={{ iconName: 'Add' }} onClick={this.addQuickTag} disabled={!quickTag.trim()} />
            </Stack>
          </div>
        </div>

        <div>
          {this.renderStep(3, 'Upload')}
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text={uploading ? 'Uploading...' : files.length > 0 ? `Upload ${files.length} file(s)` : 'Upload'}
              iconProps={{ iconName: 'Upload' }}
              onClick={this.upload}
              disabled={uploading || files.length === 0}
            />
            {files.length > 0 && !uploading && (
              <DefaultButton text="Clear" onClick={() => this.setState({ files: [], selectedTagIds: [] })} />
            )}
          </Stack>
          {uploading && <div style={{ marginTop: 12 }}><ProgressIndicator label="Uploading files..." /></div>}

          {results.length > 0 && (
            <Stack tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
              {results.map((r, i) => (
                <MessageBar
                  key={i}
                  messageBarType={r.success ? MessageBarType.success : MessageBarType.error}
                >
                  {r.fileName}: {r.success ? 'uploaded' : r.error}
                </MessageBar>
              ))}
            </Stack>
          )}
        </div>
      </Stack>
    );
  }
}
