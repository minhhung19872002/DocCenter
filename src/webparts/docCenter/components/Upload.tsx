import * as React from 'react';
import {
  PrimaryButton, DefaultButton, Stack, Label, MessageBar, MessageBarType,
  ProgressIndicator, TextField, Icon
} from '@fluentui/react';
import {
  initializeFileTypeIcons, getFileTypeIconProps
} from '@fluentui/react-file-type-icons';

initializeFileTypeIcons();
import { IHashtag, IUploadResult } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import styles from './DocCenter.module.scss';

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

  public render(): React.ReactElement {
    const { files, selectedTagIds, dragActive, uploading, results, quickTag, filter } = this.state;
    const filteredTags = this.props.hashtags.filter(
      t => !filter || t.Title.toLowerCase().includes(filter.toLowerCase())
    );

    return (
      <Stack tokens={{ childrenGap: 12 }} className={styles.section}>
        <Label>1. Choose files</Label>
        <div
          className={`${styles.dropZone} ${dragActive ? styles.dropZoneActive : ''}`}
          onDragOver={e => { e.preventDefault(); this.setState({ dragActive: true }); }}
          onDragLeave={() => this.setState({ dragActive: false })}
          onDrop={this.onDrop}
        >
          <Icon iconName="CloudUpload" style={{ fontSize: 28, color: '#0078d4' }} />
          <div style={{ marginTop: 8 }}>Drag files here, or</div>
          <div style={{ marginTop: 8 }}>
            <PrimaryButton text="Browse files" iconProps={{ iconName: 'OpenFile' }} onClick={this.onPickClick} />
            <input
              type="file"
              multiple
              ref={this.fileInputRef}
              style={{ display: 'none' }}
              onChange={this.onFileInput}
            />
          </div>
        </div>

        {files.length > 0 && (
          <div>
            <Label>{files.length} file(s) ready</Label>
            {files.map((f, i) => {
              const dot = f.name.lastIndexOf('.');
              const ext = dot >= 0 ? f.name.substring(dot + 1).toLowerCase() : '';
              return (
                <div key={i} className={styles.fileChip}>
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
                    <Icon {...getFileTypeIconProps({ extension: ext, size: 20 })} />
                    {f.name} <span className={styles.muted}>({Math.round(f.size / 1024)} KB)</span>
                  </span>
                  <DefaultButton iconProps={{ iconName: 'Cancel' }} text="Remove" onClick={() => this.removeFile(i)} />
                </div>
              );
            })}
          </div>
        )}

        <Label>2. Pick hashtags</Label>
        <TextField
          placeholder="Filter hashtags..."
          iconProps={{ iconName: 'Filter' }}
          value={filter}
          onChange={(_, v) => this.setState({ filter: v || '' })}
        />
        <div>
          {filteredTags.length === 0 && <div className={styles.muted}>No hashtags yet. Add one below.</div>}
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

        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
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

        <Label>3. Upload</Label>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={uploading ? 'Uploading...' : `Upload ${files.length || ''}`}
            iconProps={{ iconName: 'Upload' }}
            onClick={this.upload}
            disabled={uploading || files.length === 0}
          />
          {files.length > 0 && !uploading && (
            <DefaultButton text="Clear" onClick={() => this.setState({ files: [], selectedTagIds: [] })} />
          )}
        </Stack>
        {uploading && <ProgressIndicator label="Uploading files..." />}

        {results.length > 0 && (
          <Stack tokens={{ childrenGap: 4 }}>
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
      </Stack>
    );
  }
}
