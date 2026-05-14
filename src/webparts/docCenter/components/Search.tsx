import * as React from 'react';
import {
  Stack, Label, TextField, ChoiceGroup, IChoiceGroupOption,
  DefaultButton, PrimaryButton, IconButton, MessageBar, MessageBarType,
  Spinner, SpinnerSize, Link, Icon, Dialog, DialogType, DialogFooter
} from '@fluentui/react';
import { IHashtag, IDocument, SearchMode } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import styles from './DocCenter.module.scss';

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
}

const MODE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'any', text: 'Match ANY selected (OR)' },
  { key: 'all', text: 'Match ALL selected (AND)' }
];

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
    savingEdit: false
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
        this.state.nameQuery
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

  private buildAbsoluteUrl(serverRelativeUrl: string): string {
    const origin = this.props.siteUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
    return serverRelativeUrl.startsWith('/') ? `${origin}${serverRelativeUrl}` : serverRelativeUrl;
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
      editingDoc, editingTagIds, editFilter, savingEdit, editError
    } = this.state;
    const filteredTags = this.props.hashtags.filter(
      t => !filter || t.Title.toLowerCase().includes(filter.toLowerCase())
    );
    const filteredEditTags = this.props.hashtags.filter(
      t => !editFilter || t.Title.toLowerCase().includes(editFilter.toLowerCase())
    );

    return (
      <Stack tokens={{ childrenGap: 12 }} className={styles.section}>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <Stack.Item grow>
            <TextField
              label="Search by file name"
              placeholder="Type to filter by file name..."
              iconProps={{ iconName: 'Search' }}
              value={nameQuery}
              onChange={(_, v) => this.setState({ nameQuery: v || '' }, () => void this.runSearch())}
            />
          </Stack.Item>
          <Stack.Item>
            <ChoiceGroup
              label="Match mode"
              selectedKey={mode}
              options={MODE_OPTIONS}
              onChange={(_, opt) => opt && this.setState({ mode: opt.key as SearchMode }, () => void this.runSearch())}
            />
          </Stack.Item>
        </Stack>

        <Label>Hashtags {selectedTagIds.length > 0 && <Link onClick={this.clearTags}>(clear {selectedTagIds.length})</Link>}</Label>
        <TextField
          placeholder="Filter hashtags..."
          iconProps={{ iconName: 'Filter' }}
          value={filter}
          onChange={(_, v) => this.setState({ filter: v || '' })}
        />
        <div>
          {filteredTags.length === 0 && <div className={styles.muted}>No hashtags found.</div>}
          {filteredTags.map(t => {
            const selected = selectedTagIds.indexOf(t.Id) !== -1;
            return (
              <span
                key={t.Id}
                className={`${styles.tagPill} ${styles.tagPillBig} ${selected ? styles.tagPillSelected : ''}`}
                onClick={() => this.toggleTag(t.Id)}
              >
                #{t.Title}
              </span>
            );
          })}
        </div>

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Refresh" iconProps={{ iconName: 'Refresh' }} onClick={this.runSearch} disabled={searching} />
        </Stack>

        {searching && <Spinner size={SpinnerSize.medium} label="Searching..." />}

        {!searching && searched && results.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            No documents matched your criteria.
          </MessageBar>
        )}

        {!searching && results.length > 0 && (
          <div>
            <div className={styles.muted}>{results.length} document(s)</div>
            {results.map(d => (
              <div key={d.Id} className={styles.docRow}>
                <div style={{ flex: 1 }}>
                  <div>
                    <Icon iconName="Page" />{' '}
                    <Link href={this.buildAbsoluteUrl(d.ServerRelativeUrl)} target="_blank">
                      {d.Name}
                    </Link>
                  </div>
                  <div className={styles.docMeta}>
                    {d.SizeKB ? `${d.SizeKB} KB · ` : ''}
                    Modified {new Date(d.Modified).toLocaleString()}{d.ModifiedBy ? ` by ${d.ModifiedBy}` : ''}
                  </div>
                  <div style={{ marginTop: 4 }}>
                    {d.Hashtags.length === 0 && <span className={styles.muted}>No hashtags</span>}
                    {d.Hashtags.map(t => {
                      const active = selectedTagIds.indexOf(t.Id) !== -1;
                      return (
                        <span
                          key={t.Id}
                          className={`${styles.tagPill} ${active ? styles.tagPillSelected : ''}`}
                          title={active ? 'Click to remove from filter' : 'Click to add to filter'}
                          onClick={() => this.toggleTag(t.Id)}
                        >
                          #{t.Title}
                        </span>
                      );
                    })}
                  </div>
                </div>
                {this.props.isAdmin && (
                  <IconButton
                    iconProps={{ iconName: 'Edit' }}
                    title="Edit hashtags"
                    ariaLabel="Edit hashtags"
                    onClick={() => this.openEdit(d)}
                  />
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
            title: 'Edit hashtags',
            subText: editingDoc ? `For: ${editingDoc.Name}` : ''
          }}
        >
          <Stack tokens={{ childrenGap: 8 }}>
            <TextField
              placeholder="Filter hashtags..."
              iconProps={{ iconName: 'Filter' }}
              value={editFilter}
              onChange={(_, v) => this.setState({ editFilter: v || '' })}
              disabled={savingEdit}
            />
            <div style={{ maxHeight: 320, overflowY: 'auto' }}>
              {filteredEditTags.length === 0 && <div className={styles.muted}>No hashtags found.</div>}
              {filteredEditTags.map(t => {
                const selected = editingTagIds.indexOf(t.Id) !== -1;
                return (
                  <span
                    key={t.Id}
                    className={`${styles.tagPill} ${selected ? styles.tagPillSelected : ''}`}
                    onClick={() => !savingEdit && this.toggleEditTag(t.Id)}
                  >
                    #{t.Title}
                  </span>
                );
              })}
            </div>
            <div className={styles.muted}>{editingTagIds.length} selected</div>
            {editError && (
              <MessageBar messageBarType={MessageBarType.error}>{editError}</MessageBar>
            )}
          </Stack>
          <DialogFooter>
            <PrimaryButton text={savingEdit ? 'Saving...' : 'Save'} onClick={this.saveEdit} disabled={savingEdit} />
            <DefaultButton text="Cancel" onClick={this.cancelEdit} disabled={savingEdit} />
          </DialogFooter>
        </Dialog>
      </Stack>
    );
  }
}
