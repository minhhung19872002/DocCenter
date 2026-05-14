import * as React from 'react';
import {
  Stack, Label, TextField, ChoiceGroup, IChoiceGroupOption,
  DefaultButton, MessageBar, MessageBarType, Spinner, SpinnerSize, Link, Icon
} from '@fluentui/react';
import { IHashtag, IDocument, SearchMode } from '../services/types';
import { SharePointService } from '../services/SharePointService';
import styles from './DocCenter.module.scss';

interface IProps {
  libraryTitle: string;
  hashtagsListTitle: string;
  hashtags: IHashtag[];
  siteUrl: string;
}

interface IState {
  selectedTagIds: number[];
  mode: SearchMode;
  nameQuery: string;
  filter: string;
  searching: boolean;
  results: IDocument[];
  searched: boolean;
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
    searched: false
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

  public render(): React.ReactElement {
    const { selectedTagIds, mode, nameQuery, filter, searching, results, searched } = this.state;
    const filteredTags = this.props.hashtags.filter(
      t => !filter || t.Title.toLowerCase().includes(filter.toLowerCase())
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
                <div>
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
              </div>
            ))}
          </div>
        )}
      </Stack>
    );
  }
}
