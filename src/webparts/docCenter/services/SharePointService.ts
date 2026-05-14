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
import { IHashtag, IDocument, IUploadResult, SearchMode } from './types';

const HASHTAG_FIELD_INTERNAL = 'DocHashtags';
const HASHTAG_FIELD_DISPLAY = 'Hashtags';

// AddFieldOptions flags — see SharePoint AddFieldOptions enum.
const ADD_FIELD_INTERNAL_NAME_HINT = 8;
const ADD_FIELD_TO_DEFAULT_VIEW = 16;

let _sp: SPFI;
let _context: WebPartContext;

// Cache of "actual" hashtag-field internal name, keyed by library title.
// SharePoint sometimes mangles the internal name when a field is created via XML;
// callers must use the resolved name for OData property paths like `${name}Id`.
const _fieldNameByLibrary: Map<string, string> = new Map();

export class SharePointService {

  public static configure(context: WebPartContext): void {
    _context = context;
    _sp = spfi().using(SPFx(context));
  }

  // ---------------------------------------------------------------------------
  // Provisioning — runs once on first load, idempotent.
  // ---------------------------------------------------------------------------
  public static async ensureProvisioned(libraryTitle: string, hashtagsListTitle: string): Promise<void> {
    _fieldNameByLibrary.delete(libraryTitle);
    const hashtagsListId = await this.ensureHashtagsList(hashtagsListTitle);
    await this.ensureDocumentLibrary(libraryTitle);
    await this.ensureHashtagsLookupField(libraryTitle, hashtagsListId);
    // Resolve and cache the internal name now so first upload doesn't pay the cost.
    await this.resolveHashtagFieldName(libraryTitle);
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
  public static async isCurrentUserAdmin(): Promise<boolean> {
    try {
      const ownerGroup = await _sp.web.associatedOwnerGroup();
      const users = await _sp.web.siteGroups.getById(ownerGroup.Id).users();
      const me = _context.pageContext.user.loginName.toLowerCase();
      return users.some(u => (u.LoginName || '').toLowerCase() === me);
    } catch {
      return false;
    }
  }

  // ---------------------------------------------------------------------------
  // Hashtag CRUD
  // ---------------------------------------------------------------------------
  public static async getHashtags(hashtagsListTitle: string): Promise<IHashtag[]> {
    const items = await _sp.web.lists.getByTitle(hashtagsListTitle).items
      .select('Id', 'Title')
      .orderBy('Title', true)
      .top(5000)();
    return items.map(i => ({ Id: i.Id, Title: i.Title }));
  }

  public static async addHashtag(hashtagsListTitle: string, title: string): Promise<IHashtag> {
    const clean = this.normalizeHashtag(title);
    const result = await _sp.web.lists.getByTitle(hashtagsListTitle).items.add({ Title: clean });
    const newId: number = result?.Id ?? result?.data?.Id;
    return { Id: newId, Title: clean };
  }

  public static async updateHashtag(hashtagsListTitle: string, id: number, newTitle: string): Promise<void> {
    const clean = this.normalizeHashtag(newTitle);
    await _sp.web.lists.getByTitle(hashtagsListTitle).items.getById(id).update({ Title: clean });
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
    try {
      const list = _sp.web.lists.getByTitle(libraryTitle);
      const rootFolder = await list.rootFolder();
      const fileInfo = await _sp.web.getFolderByServerRelativePath(rootFolder.ServerRelativeUrl)
        .files.addUsingPath(file.name, file, { Overwrite: true });

      const item = await _sp.web.getFileByServerRelativePath(fileInfo.ServerRelativeUrl).getItem();
      if (hashtagIds.length > 0) {
        const fieldName = await this.resolveHashtagFieldName(libraryTitle);
        await item.update({ [`${fieldName}Id`]: { results: hashtagIds } });
      }
      return { success: true, fileName: file.name };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { success: false, fileName: file.name, error: msg };
    }
  }

  // ---------------------------------------------------------------------------
  // Search — by selected hashtag ids (and optional name fragment).
  // ---------------------------------------------------------------------------
  public static async searchDocuments(
    libraryTitle: string,
    selectedHashtagIds: number[],
    mode: SearchMode,
    nameQuery: string
  ): Promise<IDocument[]> {
    const list = _sp.web.lists.getByTitle(libraryTitle);
    const fieldName = await this.resolveHashtagFieldName(libraryTitle);

    const items = await list.items
      .select(
        'Id',
        'FileLeafRef',
        'FileRef',
        'Modified',
        'Editor/Title',
        'File/Length',
        `${fieldName}/Id`,
        `${fieldName}/Title`
      )
      .expand('Editor', 'File', fieldName)
      .filter("FSObjType eq 0")
      .top(5000)();

    const docs: IDocument[] = items.map((i: Record<string, unknown>) => {
      const raw = i[fieldName];
      const tags: IHashtag[] = Array.isArray(raw)
        ? (raw as Array<{ Id: number; Title: string }>).map(t => ({ Id: t.Id, Title: t.Title }))
        : [];
      const editor = i.Editor as { Title?: string } | undefined;
      const filePart = i.File as { Length?: number } | undefined;
      return {
        Id: i.Id as number,
        Name: i.FileLeafRef as string,
        ServerRelativeUrl: i.FileRef as string,
        Modified: i.Modified as string,
        ModifiedBy: editor?.Title ?? '',
        SizeKB: filePart?.Length ? Math.round(filePart.Length / 1024) : 0,
        Hashtags: tags
      };
    });

    return docs.filter(d => {
      if (nameQuery && !d.Name.toLowerCase().includes(nameQuery.toLowerCase())) {
        return false;
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
      .update({ [`${fieldName}Id`]: { results: hashtagIds } });
  }
}
