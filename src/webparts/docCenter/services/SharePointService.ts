import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/views';
import '@pnp/sp/views/list';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/security';
import '@pnp/sp/security/list';
import { PermissionKind } from '@pnp/sp/security';
import { IHashtag, IDocument, IUploadResult, SearchMode } from './types';

const HASHTAG_FIELD_INTERNAL = 'DocHashtags';
const HASHTAG_FIELD_DISPLAY = 'Hashtags';

const HASHTAG_DESC_FIELD_INTERNAL = 'HashtagDescription';
const HASHTAG_DESC_FIELD_DISPLAY = 'Description';

const HASHTAG_CAT_FIELD_INTERNAL = 'HashtagCategory';
const HASHTAG_CAT_FIELD_DISPLAY = 'Category';

// AddFieldOptions flags — see SharePoint AddFieldOptions enum.
const ADD_FIELD_INTERNAL_NAME_HINT = 8;
const ADD_FIELD_TO_DEFAULT_VIEW = 16;

let _sp: SPFI;

// Cache of "actual" hashtag-field internal name, keyed by library title.
// SharePoint sometimes mangles the internal name when a field is created via XML;
// callers must use the resolved name for OData property paths like `${name}Id`.
const _fieldNameByLibrary: Map<string, string> = new Map();

let _currentUserId: number | undefined;

export class SharePointService {

  public static configure(context: WebPartContext): void {
    _sp = spfi().using(SPFx(context));
  }

  // ---------------------------------------------------------------------------
  // Provisioning — runs once on first load, idempotent.
  // ---------------------------------------------------------------------------
  public static async ensureProvisioned(libraryTitle: string, hashtagsListTitle: string): Promise<void> {
    _fieldNameByLibrary.delete(libraryTitle);
    const hashtagsListId = await this.ensureHashtagsList(hashtagsListTitle);
    await this.ensureHashtagDescriptionField(hashtagsListTitle);
    await this.ensureHashtagCategoryField(hashtagsListTitle);
    await this.ensureDocumentLibrary(libraryTitle);
    await this.ensureHashtagsLookupField(libraryTitle, hashtagsListId);
    // Resolve and cache the internal name now so first upload doesn't pay the cost.
    await this.resolveHashtagFieldName(libraryTitle);
  }

  private static async ensureHashtagCategoryField(hashtagsListTitle: string): Promise<void> {
    const list = _sp.web.lists.getByTitle(hashtagsListTitle);
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_CAT_FIELD_INTERNAL)();
      return;
    } catch { /* not found, fall through */ }
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_CAT_FIELD_DISPLAY)();
      return;
    } catch { /* not found, fall through */ }
    const schemaXml = `<Field Type="Text" `
      + `DisplayName="${HASHTAG_CAT_FIELD_INTERNAL}" `
      + `StaticName="${HASHTAG_CAT_FIELD_INTERNAL}" `
      + `Name="${HASHTAG_CAT_FIELD_INTERNAL}" `
      + `MaxLength="80" />`;
    await list.fields.createFieldAsXml({
      SchemaXml: schemaXml,
      Options: ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW
    });
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_CAT_FIELD_INTERNAL)
        .update({ Title: HASHTAG_CAT_FIELD_DISPLAY });
    } catch {
      // Non-fatal — column still usable.
    }
  }

  private static async ensureHashtagDescriptionField(hashtagsListTitle: string): Promise<void> {
    const list = _sp.web.lists.getByTitle(hashtagsListTitle);
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_DESC_FIELD_INTERNAL)();
      return;
    } catch { /* not found, fall through */ }
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_DESC_FIELD_DISPLAY)();
      return;
    } catch { /* not found, fall through */ }
    const schemaXml = `<Field Type="Note" `
      + `DisplayName="${HASHTAG_DESC_FIELD_INTERNAL}" `
      + `StaticName="${HASHTAG_DESC_FIELD_INTERNAL}" `
      + `Name="${HASHTAG_DESC_FIELD_INTERNAL}" `
      + `NumLines="3" RichText="FALSE" AppendOnly="FALSE" />`;
    await list.fields.createFieldAsXml({
      SchemaXml: schemaXml,
      Options: ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW
    });
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_DESC_FIELD_INTERNAL)
        .update({ Title: HASHTAG_DESC_FIELD_DISPLAY });
    } catch {
      // Non-fatal — column still usable.
    }
  }

  private static async ensureHashtagsList(title: string): Promise<string> {
    try {
      const list = await _sp.web.lists.getByTitle(title)();
      return list.Id;
    } catch {
      const created = await _sp.web.lists.add(title, 'Hashtags managed by Document Center', 100, false);
      return created.Id;
    }
  }

  private static async ensureDocumentLibrary(title: string): Promise<void> {
    try {
      await _sp.web.lists.getByTitle(title)();
    } catch {
      // 101 = Document Library template
      await _sp.web.lists.add(title, 'Documents managed by Document Center', 101, false);
    }
  }

  private static async ensureHashtagsLookupField(libraryTitle: string, hashtagsListId: string): Promise<void> {
    const list = _sp.web.lists.getByTitle(libraryTitle);
    // First, check if a field with the intended internal name OR display name already exists.
    if (await this.hashtagFieldExists(libraryTitle)) {
      return;
    }
    const cleanId = hashtagsListId.replace(/[{}]/g, '');
    const schemaXml = `<Field Type="LookupMulti" Mult="TRUE" `
      + `DisplayName="${HASHTAG_FIELD_INTERNAL}" `
      + `StaticName="${HASHTAG_FIELD_INTERNAL}" `
      + `Name="${HASHTAG_FIELD_INTERNAL}" `
      + `List="{${cleanId}}" ShowField="Title" />`;
    await list.fields.createFieldAsXml({
      SchemaXml: schemaXml,
      Options: ADD_FIELD_INTERNAL_NAME_HINT | ADD_FIELD_TO_DEFAULT_VIEW
    });
    // After creating with the internal name forced, set the user-friendly display name.
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_FIELD_INTERNAL)
        .update({ Title: HASHTAG_FIELD_DISPLAY });
    } catch {
      // Non-fatal — column still usable.
    }
  }

  private static async hashtagFieldExists(libraryTitle: string): Promise<boolean> {
    const list = _sp.web.lists.getByTitle(libraryTitle);
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_FIELD_INTERNAL)();
      return true;
    } catch { /* ignore */ }
    try {
      await list.fields.getByInternalNameOrTitle(HASHTAG_FIELD_DISPLAY)();
      return true;
    } catch { /* ignore */ }
    return false;
  }

  /**
   * Resolves the actual InternalName of the hashtags lookup field on the given library,
   * caches it for the lifetime of this page load.
   */
  private static async resolveHashtagFieldName(libraryTitle: string): Promise<string> {
    const cached = _fieldNameByLibrary.get(libraryTitle);
    if (cached) return cached;
    const list = _sp.web.lists.getByTitle(libraryTitle);
    let info: { InternalName?: string } | undefined;
    try {
      info = await list.fields.getByInternalNameOrTitle(HASHTAG_FIELD_INTERNAL).select('InternalName')();
    } catch { /* fall through */ }
    if (!info?.InternalName) {
      info = await list.fields.getByInternalNameOrTitle(HASHTAG_FIELD_DISPLAY).select('InternalName')();
    }
    const name = info?.InternalName || HASHTAG_FIELD_INTERNAL;
    _fieldNameByLibrary.set(libraryTitle, name);
    return name;
  }

  // ---------------------------------------------------------------------------
  // Admin check — owner group of the current site.
  // ---------------------------------------------------------------------------
  /**
   * Returns true if the current user has Owner-level rights on the document library
   * (i.e. they have ManagePermissions on the list — the privilege that distinguishes
   * Owners/Full-Control from Contribute/Edit users).
   */
  private static async getCurrentUserId(): Promise<number> {
    if (_currentUserId !== undefined) return _currentUserId;
    const user = await _sp.web.currentUser();
    _currentUserId = user.Id;
    return _currentUserId;
  }

  public static async isCurrentUserAdmin(libraryTitle: string): Promise<boolean> {
    try {
      const perms = await _sp.web.lists.getByTitle(libraryTitle).getCurrentUserEffectivePermissions();
      return this.hasPermission(perms, PermissionKind.ManagePermissions);
    } catch (e) {
      // eslint-disable-next-line no-console
      console.warn('[DocCenter] admin check failed:', e);
      return false;
    }
  }

  /**
   * Bitwise check against SharePoint BasePermissions { Low, High }.
   * Each PermissionKind value is a 1-based bit position into a 64-bit mask;
   * bits 1-32 live in `Low`, bits 33-64 in `High`.
   */
  private static hasPermission(perms: { Low: number; High: number }, kind: PermissionKind): boolean {
    if (kind <= 0) return false;
    const bit = (kind as number) - 1;
    if (bit < 32) {
      // eslint-disable-next-line no-bitwise
      return (perms.Low & (1 << bit)) !== 0;
    }
    // eslint-disable-next-line no-bitwise
    return (perms.High & (1 << (bit - 32))) !== 0;
  }

  // ---------------------------------------------------------------------------
  // Hashtag CRUD
  // ---------------------------------------------------------------------------
  public static async getHashtags(hashtagsListTitle: string): Promise<IHashtag[]> {
    const items = await _sp.web.lists.getByTitle(hashtagsListTitle).items
      .select('Id', 'Title', HASHTAG_DESC_FIELD_INTERNAL, HASHTAG_CAT_FIELD_INTERNAL)
      .orderBy('Title', true)
      .top(5000)();
    return items.map((i: Record<string, unknown>) => ({
      Id: i.Id as number,
      Title: i.Title as string,
      Description: (i[HASHTAG_DESC_FIELD_INTERNAL] as string) || '',
      Category: (i[HASHTAG_CAT_FIELD_INTERNAL] as string) || ''
    }));
  }

  public static async addHashtag(
    hashtagsListTitle: string,
    title: string,
    description?: string,
    category?: string
  ): Promise<IHashtag> {
    const clean = this.normalizeHashtag(title);
    const desc = (description || '').trim();
    const cat = (category || '').trim();
    const payload: Record<string, unknown> = { Title: clean };
    if (desc) payload[HASHTAG_DESC_FIELD_INTERNAL] = desc;
    if (cat) payload[HASHTAG_CAT_FIELD_INTERNAL] = cat;
    const result = await _sp.web.lists.getByTitle(hashtagsListTitle).items.add(payload);
    const newId: number = result?.Id ?? result?.data?.Id;
    return { Id: newId, Title: clean, Description: desc, Category: cat };
  }

  public static async updateHashtag(
    hashtagsListTitle: string,
    id: number,
    newTitle: string,
    description?: string,
    category?: string
  ): Promise<void> {
    const clean = this.normalizeHashtag(newTitle);
    const payload: Record<string, unknown> = { Title: clean };
    if (description !== undefined) {
      payload[HASHTAG_DESC_FIELD_INTERNAL] = description.trim();
    }
    if (category !== undefined) {
      payload[HASHTAG_CAT_FIELD_INTERNAL] = category.trim();
    }
    await _sp.web.lists.getByTitle(hashtagsListTitle).items.getById(id).update(payload);
  }

  public static async deleteHashtag(hashtagsListTitle: string, id: number): Promise<void> {
    await _sp.web.lists.getByTitle(hashtagsListTitle).items.getById(id).delete();
  }

  private static normalizeHashtag(raw: string): string {
    const trimmed = (raw || '').trim().replace(/^#+/, '');
    return trimmed.replace(/\s+/g, '-');
  }

  // ---------------------------------------------------------------------------
  // Document upload
  // ---------------------------------------------------------------------------
  public static async uploadDocument(
    libraryTitle: string,
    file: File,
    hashtagIds: number[]
  ): Promise<IUploadResult> {
    const uniqueName = this.appendTimestampToFileName(file.name);
    try {
      const yearFolderUrl = await this.ensureYearFolder(libraryTitle);
      const fileInfo = await _sp.web.getFolderByServerRelativePath(yearFolderUrl)
        .files.addUsingPath(uniqueName, file, { Overwrite: true });

      const item = await _sp.web.getFileByServerRelativePath(fileInfo.ServerRelativeUrl).getItem();
      if (hashtagIds.length > 0) {
        const fieldName = await this.resolveHashtagFieldName(libraryTitle);
        // PnP v4 / modern REST: multi-lookup expects a bare array, not { results: [...] }.
        await item.update({ [`${fieldName}Id`]: hashtagIds });
      }
      return { success: true, fileName: uniqueName };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { success: false, fileName: uniqueName, error: msg };
    }
  }

  private static async ensureYearFolder(libraryTitle: string): Promise<string> {
    const year = new Date().getFullYear().toString();
    const list = _sp.web.lists.getByTitle(libraryTitle);
    const rootFolder = await list.rootFolder();
    const yearFolderUrl = `${rootFolder.ServerRelativeUrl}/${year}`;
    try {
      await _sp.web.getFolderByServerRelativePath(yearFolderUrl)();
      return yearFolderUrl;
    } catch {
      await _sp.web.folders.addUsingPath(yearFolderUrl);
      return yearFolderUrl;
    }
  }

  private static appendTimestampToFileName(originalName: string): string {
    const d = new Date();
    const pad = (n: number): string => n.toString().padStart(2, '0');
    const stamp = `${pad(d.getDate())}${pad(d.getMonth() + 1)}${d.getFullYear()}`
      + `_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
    const dot = originalName.lastIndexOf('.');
    if (dot <= 0) return `${originalName}_${stamp}`;
    return `${originalName.substring(0, dot)}_${stamp}${originalName.substring(dot)}`;
  }

  // ---------------------------------------------------------------------------
  // Search — by selected hashtag ids (and optional name fragment).
  // ---------------------------------------------------------------------------
  public static async searchDocuments(
    libraryTitle: string,
    selectedHashtagIds: number[],
    mode: SearchMode,
    nameQuery: string,
    restrictToCurrentUser: boolean = false
  ): Promise<IDocument[]> {
    const list = _sp.web.lists.getByTitle(libraryTitle);
    const fieldName = await this.resolveHashtagFieldName(libraryTitle);

    let filter = 'FSObjType eq 0';
    if (restrictToCurrentUser) {
      const uid = await this.getCurrentUserId();
      filter += ` and AuthorId eq ${uid}`;
    }

    const items = await list.items
      .select(
        'Id',
        'FileLeafRef',
        'FileRef',
        'Created',
        'Modified',
        'Author/Title',
        'File/Length',
        `${fieldName}/Id`,
        `${fieldName}/Title`
      )
      .expand('Author', 'File', fieldName)
      .filter(filter)
      .top(5000)();

    const docs: IDocument[] = items.map((i: Record<string, unknown>) => {
      const raw = i[fieldName];
      const tags: IHashtag[] = Array.isArray(raw)
        ? (raw as Array<{ Id: number; Title: string }>).map(t => ({ Id: t.Id, Title: t.Title }))
        : [];
      const author = i.Author as { Title?: string } | undefined;
      const filePart = i.File as { Length?: number } | undefined;
      return {
        Id: i.Id as number,
        Name: i.FileLeafRef as string,
        ServerRelativeUrl: i.FileRef as string,
        Created: i.Created as string,
        Modified: i.Modified as string,
        CreatedBy: author?.Title ?? '',
        SizeKB: filePart?.Length ? Math.round(filePart.Length / 1024) : 0,
        Hashtags: tags
      };
    });

    return docs.filter(d => {
      if (nameQuery) {
        const q = nameQuery.toLowerCase();
        const inName = d.Name.toLowerCase().includes(q);
        const inTags = d.Hashtags.some(t => t.Title.toLowerCase().includes(q));
        if (!inName && !inTags) {
          return false;
        }
      }
      if (selectedHashtagIds.length === 0) {
        return true;
      }
      const docTagIds = d.Hashtags.map(t => t.Id);
      return mode === 'all'
        ? selectedHashtagIds.every(id => docTagIds.indexOf(id) !== -1)
        : selectedHashtagIds.some(id => docTagIds.indexOf(id) !== -1);
    }).sort((a, b) => b.Modified.localeCompare(a.Modified));
  }

  public static async updateDocumentHashtags(
    libraryTitle: string,
    itemId: number,
    hashtagIds: number[]
  ): Promise<void> {
    const fieldName = await this.resolveHashtagFieldName(libraryTitle);
    await _sp.web.lists.getByTitle(libraryTitle).items.getById(itemId)
      .update({ [`${fieldName}Id`]: hashtagIds });
  }

  public static async deleteDocument(libraryTitle: string, itemId: number): Promise<void> {
    await _sp.web.lists.getByTitle(libraryTitle).items.getById(itemId).recycle();
  }

  public static async renameDocument(libraryTitle: string, itemId: number, newName: string): Promise<void> {
    const clean = (newName || '').trim();
    if (!clean) throw new Error('New name cannot be empty.');
    if (/[\\/:*?"<>|]/.test(clean)) {
      throw new Error('Name contains invalid characters (\\ / : * ? " < > |).');
    }
    const list = _sp.web.lists.getByTitle(libraryTitle);
    const item = await list.items.getById(itemId).select('FileRef')() as { FileRef: string };
    const currentPath = item.FileRef;
    const slash = currentPath.lastIndexOf('/');
    const parent = currentPath.substring(0, slash);
    const newPath = `${parent}/${clean}`;
    if (newPath === currentPath) return;
    await _sp.web.getFileByServerRelativePath(currentPath).moveByPath(newPath, true);
  }
}
