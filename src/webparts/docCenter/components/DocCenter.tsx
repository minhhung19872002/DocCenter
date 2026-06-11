import * as React from 'react';
import { Pivot, PivotItem, Spinner, SpinnerSize, MessageBar, MessageBarType, ThemeProvider } from '@fluentui/react';
import { IDocCenterProps } from './IDocCenterProps';
import { SharePointService } from '../services/SharePointService';
import { IHashtag } from '../services/types';
import { Upload } from './Upload';
import { Search } from './Search';
import { Admin } from './Admin';
import { claudeTheme } from './theme';
import styles from './DocCenter.module.scss';

interface IState {
  loading: boolean;
  error?: string;
  isAdmin: boolean;
  hashtags: IHashtag[];
}

export class DocCenter extends React.Component<IDocCenterProps, IState> {

  public state: IState = { loading: true, isAdmin: false, hashtags: [] };

  public async componentDidMount(): Promise<void> {
    try {
      await SharePointService.ensureProvisioned(this.props.documentsLibraryTitle, this.props.hashtagsListTitle);
      const [hashtags, isAdmin] = await Promise.all([
        SharePointService.getHashtags(this.props.hashtagsListTitle),
        SharePointService.isCurrentUserAdmin(this.props.documentsLibraryTitle)
      ]);
      this.setState({ loading: false, hashtags, isAdmin });
    } catch (e) {
      this.setState({ loading: false, error: e instanceof Error ? e.message : String(e) });
    }
  }

  private reloadHashtags = async (): Promise<void> => {
    const hashtags = await SharePointService.getHashtags(this.props.hashtagsListTitle);
    this.setState({ hashtags });
  };

  public render(): React.ReactElement {
    const { loading, error, isAdmin, hashtags } = this.state;

    if (loading) {
      return (
        <ThemeProvider theme={claudeTheme}>
          <div className={styles.docCenter}>
            <Spinner size={SpinnerSize.large} label="Đang tải Trung tâm Tài liệu..." />
          </div>
        </ThemeProvider>
      );
    }

    if (error) {
      return (
        <ThemeProvider theme={claudeTheme}>
          <div className={styles.docCenter}>
            <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
          </div>
        </ThemeProvider>
      );
    }

    return (
      <ThemeProvider theme={claudeTheme}>
        <div className={styles.docCenter}>
          <div className={styles.header}>
            <div className={styles.headerLeft}>
              <div className={styles.title}>Trung tâm Tài liệu</div>
              <div className={styles.subtitle}>
                Tải lên tài liệu, gắn hashtag và tìm kiếm nhanh chóng.
              </div>
            </div>
            <div className={styles.headerRight}>
              {isAdmin && <span className={styles.adminBadge}>● Quản trị</span>}
              <span className={styles.metaChip}>
                Thư viện: <b style={{ marginLeft: 4 }}>{this.props.documentsLibraryTitle}</b>
              </span>
              <span className={styles.metaChip}>
                Hashtag: <b style={{ marginLeft: 4 }}>{this.props.hashtagsListTitle}</b>
              </span>
            </div>
          </div>

          <Pivot styles={{ root: { marginTop: 16, marginBottom: 8 } }}>
            <PivotItem headerText="Tìm kiếm" itemIcon="Search">
              <Search
                libraryTitle={this.props.documentsLibraryTitle}
                hashtagsListTitle={this.props.hashtagsListTitle}
                hashtags={hashtags}
                siteUrl={this.props.siteUrl}
                isAdmin={isAdmin}
              />
            </PivotItem>

            <PivotItem headerText="Tải lên" itemIcon="Upload">
              <Upload
                libraryTitle={this.props.documentsLibraryTitle}
                hashtags={hashtags}
                onHashtagsChanged={this.reloadHashtags}
                hashtagsListTitle={this.props.hashtagsListTitle}
                isAdmin={isAdmin}
              />
            </PivotItem>

            {isAdmin && (
              <PivotItem headerText="Quản trị" itemIcon="Settings">
                <Admin
                  hashtagsListTitle={this.props.hashtagsListTitle}
                  libraryTitle={this.props.documentsLibraryTitle}
                  hashtags={hashtags}
                  siteUrl={this.props.siteUrl}
                  onChange={this.reloadHashtags}
                />
              </PivotItem>
            )}
          </Pivot>
        </div>
      </ThemeProvider>
    );
  }
}
