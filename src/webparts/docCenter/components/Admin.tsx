import * as React from 'react';
import {
  Stack, TextField, PrimaryButton, DefaultButton, IconButton,
  MessageBar, MessageBarType, Dialog, DialogType, DialogFooter,
  Pivot, PivotItem
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
  filter: string;
  editingId?: number;
  editingValue: string;
  deletingId?: number;
  status?: { type: MessageBarType; text: string };
  busy: boolean;
}

export class Admin extends React.Component<IProps, IState> {

  public state: IState = { newTag: '', filter: '', editingValue: '', busy: false };

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
    await SharePointService.addHashtag(this.props.hashtagsListTitle, this.state.newTag);
    this.setState({ newTag: '' });
  }, 'Hashtag added.');

  private startEdit = (h: IHashtag): void =>
    this.setState({ editingId: h.Id, editingValue: h.Title });

  private cancelEdit = (): void =>
    this.setState({ editingId: undefined, editingValue: '' });

  private saveEdit = (): Promise<void> => this.withBusy(async () => {
    if (this.state.editingId == null) return;
    await SharePointService.updateHashtag(this.props.hashtagsListTitle, this.state.editingId, this.state.editingValue);
    this.setState({ editingId: undefined, editingValue: '' });
  }, 'Hashtag updated. Documents now reflect the new name.');

  private confirmDelete = (): Promise<void> => this.withBusy(async () => {
    if (this.state.deletingId == null) return;
    await SharePointService.deleteHashtag(this.props.hashtagsListTitle, this.state.deletingId);
    this.setState({ deletingId: undefined });
  }, 'Hashtag deleted.');

  private renderHashtagManager(): React.ReactElement {
    const { newTag, filter, editingId, editingValue, deletingId, status, busy } = this.state;
    const sorted = [...this.props.hashtags].sort((a, b) => a.Title.localeCompare(b.Title));
    const filtered = sorted.filter(t => !filter || t.Title.toLowerCase().includes(filter.toLowerCase()));
    const deletingTag = this.props.hashtags.find(t => t.Id === deletingId);

    return (
      <Stack tokens={{ childrenGap: 16 }} className={styles.section}>
        <MessageBar messageBarType={MessageBarType.info}>
          Renaming a hashtag updates it everywhere automatically. Deleting removes the tag from all documents that use it.
        </MessageBar>

        <div className={styles.card}>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
            <Stack.Item grow>
              <TextField
                label="Add hashtag"
                placeholder="e.g. invoice"
                value={newTag}
                onChange={(_, v) => this.setState({ newTag: v || '' })}
                onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.addTag(); } }}
                disabled={busy}
              />
            </Stack.Item>
            <PrimaryButton text="Add" iconProps={{ iconName: 'Add' }} onClick={this.addTag} disabled={busy || !newTag.trim()} />
          </Stack>
        </div>

        {status && (
          <MessageBar messageBarType={status.type} onDismiss={() => this.setState({ status: undefined })}>
            {status.text}
          </MessageBar>
        )}

        <div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
            <span className={styles.sectionLabel}>All hashtags</span>
            <span className={styles.countPill}>{sorted.length}</span>
          </div>
          <TextField
            placeholder="Filter..."
            iconProps={{ iconName: 'Filter' }}
            value={filter}
            onChange={(_, v) => this.setState({ filter: v || '' })}
          />
        </div>

        <Stack tokens={{ childrenGap: 8 }}>
          {filtered.length === 0 && <span className={styles.muted}>No hashtags.</span>}
          {filtered.map(h => (
            <div key={h.Id} className={styles.adminRow}>
              {editingId === h.Id ? (
                <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" grow>
                  <Stack.Item grow>
                    <TextField
                      value={editingValue}
                      onChange={(_, v) => this.setState({ editingValue: v || '' })}
                      onKeyDown={e => { if (e.key === 'Enter') { e.preventDefault(); void this.saveEdit(); } }}
                      disabled={busy}
                    />
                  </Stack.Item>
                  <PrimaryButton text="Save" iconProps={{ iconName: 'Save' }} onClick={this.saveEdit} disabled={busy || !editingValue.trim()} />
                  <DefaultButton text="Cancel" onClick={this.cancelEdit} disabled={busy} />
                </Stack>
              ) : (
                <>
                  <span className={`${styles.tagPill} ${styles.tagPillReadonly} ${styles.tagPillBig}`}>#{h.Title}</span>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <IconButton iconProps={{ iconName: 'Edit' }} title="Rename" onClick={() => this.startEdit(h)} />
                    <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => this.setState({ deletingId: h.Id })} />
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
            title: 'Delete hashtag?',
            subText: deletingTag
              ? `"#${deletingTag.Title}" will be removed from all documents that use it. This cannot be undone.`
              : ''
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this.confirmDelete} text="Delete" disabled={busy} />
            <DefaultButton onClick={() => this.setState({ deletingId: undefined })} text="Cancel" disabled={busy} />
          </DialogFooter>
        </Dialog>
      </Stack>
    );
  }

  public render(): React.ReactElement {
    return (
      <Pivot>
        <PivotItem headerText="Manage hashtags" itemIcon="Tag">
          {this.renderHashtagManager()}
        </PivotItem>
        <PivotItem headerText="Edit document tags" itemIcon="EditNote">
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
