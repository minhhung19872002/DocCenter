import * as React from 'react';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, ThemeProvider } from '@fluentui/react';
import { IDocCenterProps } from './IDocCenterProps';
import { SharePointService } from '../services/SharePointService';
import { IHashtag } from '../services/types';
import { Upload } from './Upload';
import { Search } from './Search';
import { Admin } from './Admin';
import { docCenterTheme } from './theme';
import styles from './DocCenter.module.scss';

type TabKey = 'search' | 'upload' | 'admin';

interface IState {
  loading: boolean;
  error?: string;
  isAdmin: boolean;
  hashtags: IHashtag[];
  activeTab: TabKey;
}

export class DocCenter extends React.Component<IDocCenterProps, IState> {

  public state: IState = { loading: true, isAdmin: false, hashtags: [], activeTab: 'search' };

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

  private goToSearch = (): void => this.setState({ activeTab: 'search' });

  private getUserInitials(): string {
    const name = (this.props.currentUserName || this.props.currentUserLogin || '').trim();
    const source = name.indexOf('@') !== -1 ? name.split('@')[0].replace(/[._-]+/g, ' ') : name;
    const words = source.split(/\s+/).filter(w => w.length > 0);
    if (words.length === 0) return '?';
    if (words.length === 1) return words[0].substring(0, 2).toUpperCase();
    return (words[words.length - 2][0] + words[words.length - 1][0]).toUpperCase();
  }

  private renderHeader(): React.ReactElement {
    const { isAdmin, activeTab } = this.state;
    const tabs: Array<{ key: TabKey; label: string }> = [
      { key: 'search', label: '⌕  Tìm kiếm' },
      { key: 'upload', label: '↑  Tải lên' }
    ];
    if (isAdmin) tabs.push({ key: 'admin', label: '⚙︎  Quản trị' });

    return (
      <div className={styles.headerBand}>
        <div className={styles.headerInner}>
          <div className={styles.brandRow}>
            <div className={styles.logoBox}>T</div>
            <div>
              <div className={styles.brandTitle}>Trung tâm Tài liệu</div>
              <div className={styles.brandSub}>Tải lên tài liệu, gắn hashtag và tìm kiếm nhanh chóng</div>
            </div>
            <div className={styles.brandRight}>
              {isAdmin && <span className={styles.adminBadge}>Quản trị</span>}
              <span className={styles.libChip}>
                Thư viện <b>{this.props.documentsLibraryTitle}</b>
              </span>
              <div className={styles.avatar}>{this.getUserInitials()}</div>
            </div>
          </div>
          <div className={styles.tabStrip}>
            {tabs.map(t => (
              <div
                key={t.key}
                className={`${styles.tab} ${activeTab === t.key ? styles.tabActive : ''}`}
                onClick={() => this.setState({ activeTab: t.key })}
              >
                {t.label}
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement {
    const { loading, error, isAdmin, hashtags, activeTab } = this.state;

    if (loading) {
      return (
        <ThemeProvider theme={docCenterTheme}>
          <div className={styles.docCenter}>
            <div className={styles.content}>
              <Spinner size={SpinnerSize.large} label="Đang tải Trung tâm Tài liệu..." styles={{ root: { marginTop: 64 } }} />
            </div>
          </div>
        </ThemeProvider>
      );
    }

    if (error) {
      return (
        <ThemeProvider theme={docCenterTheme}>
          <div className={styles.docCenter}>
            <div className={styles.content}>
              <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
            </div>
          </div>
        </ThemeProvider>
      );
    }

    return (
      <ThemeProvider theme={docCenterTheme}>
        <div className={styles.docCenter}>
          {this.renderHeader()}

          <div className={styles.content}>
            {activeTab === 'search' && (
              <Search
                libraryTitle={this.props.documentsLibraryTitle}
                hashtagsListTitle={this.props.hashtagsListTitle}
                hashtags={hashtags}
                siteUrl={this.props.siteUrl}
                isAdmin={isAdmin}
              />
            )}

            {activeTab === 'upload' && (
              <Upload
                libraryTitle={this.props.documentsLibraryTitle}
                hashtags={hashtags}
                onHashtagsChanged={this.reloadHashtags}
                hashtagsListTitle={this.props.hashtagsListTitle}
                isAdmin={isAdmin}
                onGoToSearch={this.goToSearch}
              />
            )}

            {activeTab === 'admin' && isAdmin && (
              <Admin
                hashtagsListTitle={this.props.hashtagsListTitle}
                libraryTitle={this.props.documentsLibraryTitle}
                hashtags={hashtags}
                siteUrl={this.props.siteUrl}
                onChange={this.reloadHashtags}
              />
            )}
          </div>
        </div>
      </ThemeProvider>
    );
  }
}
