function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ファイルナビ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * DriveフォルダのURL or ID からファイル一覧を返す
 * 設定（既定ソート/件数制限）を適用
 */
function listFiles(folderInput) {
  try {
    Logger.log('[listFiles] entry folderInput="%s" len=%s', String(folderInput || ''), String(folderInput || '').length);
    const folderId = extractFolderId_(folderInput);
    Logger.log('[listFiles] extracted folderId="%s" ok=%s', folderId, !!folderId);
    if (!folderId) {
      throw new Error('フォルダIDを取得できませんでした。URLまたはIDを確認してください。入力値: ' + String(folderInput || '').substring(0, 100));
    }

    // 設定を読み込み
    const settings = getSettings();
    const limit = settings.limit > 0 ? settings.limit : null;
    const defaultSort = settings.defaultSort || 'modifiedTimeDesc';
    
    // Drive APIのorderByを決定
    let orderBy = null;
    if (defaultSort === 'modifiedTimeDesc') {
      orderBy = 'modifiedTime desc';
    } else if (defaultSort === 'modifiedTimeAsc') {
      orderBy = 'modifiedTime asc';
    } else if (defaultSort === 'nameAsc') {
      orderBy = 'name asc';
    } else if (defaultSort === 'nameDesc') {
      orderBy = 'name desc';
    }
    
    Logger.log('[listFiles] settings: sort=%s limit=%s orderBy=%s', defaultSort, limit || 'none', orderBy || 'none');

    const files = [];
    let pageToken = null;
    const FOLDER_MIME = 'application/vnd.google-apps.folder';
    let pageCount = 0;
    const MAX_PAGES = 50; // 無限ループ防止
    let totalFetched = 0;

    do {
      pageCount++;
      if (pageCount > MAX_PAGES) {
        throw new Error('ページネーションが上限（' + MAX_PAGES + 'ページ）を超えました。フォルダが大きすぎる可能性があります。');
      }

      // 件数制限がある場合、このページで取得する件数を計算
      let pageSize = 200;
      if (limit && limit > 0) {
        const remaining = limit - totalFetched;
        if (remaining <= 0) break; // 既に制限に達している
        pageSize = Math.min(200, remaining);
      }

      let res;
      try {
        const queryOptions = {
          q: `'${folderId}' in parents and trashed = false`,
          maxResults: pageSize,
          pageSize: pageSize,
          pageToken,
          supportsAllDrives: true,
          includeItemsFromAllDrives: true,
          fields: 'files(id,name,mimeType,modifiedTime,size,webViewLink,webContentLink),nextPageToken'
        };
        
        // orderByが指定されている場合のみ追加
        if (orderBy) {
          queryOptions.orderBy = orderBy;
        }
        
        res = Drive.Files.list(queryOptions);
      } catch (apiError) {
        // Drive API エラーを詳細化
        const errorMsg = apiError && apiError.message ? apiError.message : String(apiError);
        Logger.log('[listFiles] Drive.Files.list error: %s', errorMsg);
        if (errorMsg.includes('permission') || errorMsg.includes('権限')) {
          throw new Error('フォルダへのアクセス権限がありません。フォルダID: ' + folderId + '\nエラー詳細: ' + errorMsg);
        } else if (errorMsg.includes('not found') || errorMsg.includes('見つかりません')) {
          throw new Error('フォルダが見つかりません。フォルダID: ' + folderId + '\nエラー詳細: ' + errorMsg);
        } else {
          throw new Error('Drive API エラー: ' + errorMsg + '\nフォルダID: ' + folderId);
        }
      }

      if (!res) {
        throw new Error('Drive API のレスポンスが空です。フォルダID: ' + folderId);
      }

      const items = res.files || res.items || [];
      Logger.log('[listFiles] page=%s token=%s items=%s (files=%s items=%s)',
        String(pageCount),
        String(pageToken || ''),
        String(items.length),
        String((res.files || []).length),
        String((res.items || []).length)
      );

      items.forEach(f => {
        const mimeType = f.mimeType || '';
        const isFolder = mimeType === FOLDER_MIME;
        const name = f.name || f.title || '';
        const modified = f.modifiedTime || f.modifiedDate || '';
        const size = Number(f.size || f.fileSize || 0);
        const viewLinkRaw = f.webViewLink || f.alternateLink || '';
        const downloadLinkRaw = f.webContentLink || '';
        files.push({
          id: f.id,
          name,
          mimeType,
          isFolder,
          updated: modified ? String(modified).replace('T', ' ').slice(0, 16) : '',
          size,
          viewLink: isFolder ? `https://drive.google.com/drive/folders/${f.id}` : viewLinkRaw,
          downloadLink: isFolder ? '' : downloadLinkRaw
        });
        totalFetched++;
      });

      // 件数制限に達したら終了
      if (limit && limit > 0 && totalFetched >= limit) {
        Logger.log('[listFiles] reached limit: %s', limit);
        break;
      }

      pageToken = res.nextPageToken;
    } while (pageToken);

    // orderByが指定されていない場合、またはクライアント側で再ソートが必要な場合
    // （Drive APIのorderByが効かない場合のフォールバック）
    if (!orderBy || defaultSort === 'modifiedTimeDesc') {
      // 更新日降順（デフォルト）
      files.sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
    } else if (defaultSort === 'modifiedTimeAsc') {
      files.sort((a, b) => (a.updated || '').localeCompare(b.updated || ''));
    } else if (defaultSort === 'nameAsc') {
      files.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    } else if (defaultSort === 'nameDesc') {
      files.sort((a, b) => (b.name || '').localeCompare(a.name || ''));
    }

    Logger.log('[listFiles] returning: count=%s (limit=%s)', files.length, limit || 'none');
    return { folderId, count: files.length, files };
  } catch (e) {
    // エラーを再スロー（詳細メッセージ付き）
    const errorMsg = e && e.message ? e.message : String(e);
    Logger.log('[listFiles] throw: %s', errorMsg);
    throw new Error('listFiles エラー: ' + errorMsg);
  }
}


/**
 * Folder URL -> folderId
 * 例:
 * - https://drive.google.com/drive/folders/{ID}
 * - https://drive.google.com/drive/u/0/folders/{ID}
 * - {ID} (直接ID)
 */
function extractFolderId_(input) {
  if (!input) return '';
  const s = String(input).trim();
  if (!s) return '';

  // すでにIDっぽい（英数-_、最低15文字以上、Drive IDは通常28文字だが短いものもある）
  // ただし、URLっぽいもの（http:// や / を含む）は除外
  if (!s.includes('/') && !s.includes('http') && /^[a-zA-Z0-9_-]{15,}$/.test(s)) {
    return s;
  }

  // folders/{id} パターン（最も一般的）
  // https://drive.google.com/drive/folders/{ID}
  // https://drive.google.com/drive/u/0/folders/{ID}
  const m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  // id={id} パターン（クエリパラメータ）
  const q = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (q && q[1]) return q[1];

  // 最後の手段：URLの最後の部分がIDっぽい場合（例: .../folders/ABC123...）
  const lastPart = s.split('/').pop();
  if (lastPart && /^[a-zA-Z0-9_-]{15,}$/.test(lastPart.split('?')[0])) {
    return lastPart.split('?')[0];
  }

  return '';
}

function safeGetSize_(file) {
  try {
    const bytes = file.getSize();
    return bytes ? bytes : 0;
  } catch (e) {
    return 0;
  }
}

/**
 * フォルダツリー用：指定親配下のフォルダ一覧を返す（共有ドライブ対応）
 * @param {string} parentKey - 'root' | 'sharedDrives' | 'drive:<driveId>' | '<folderId>'
 * @param {string} pageToken - ページネーション用（省略可）
 * @returns {Object} { parentKey, nodes: [{id, name, type}], nextPageToken }
 */
function listChildFolders(parentKey, pageToken) {
  try {
    // parentKey の型と値を詳細にログ出力
    Logger.log('[listChildFolders] entry parentKey type=%s value="%s" pageToken="%s"', 
      typeof parentKey, 
      String(parentKey !== null && parentKey !== undefined ? parentKey : 'null/undefined'), 
      String(pageToken || ''));
    
    // parentKey の検証を最初に行う（より厳密に）
    // NOTE: 呼び出し側の不具合や手動実行で引数なしのケースがあり得るため、
    // ここでは throw せず「空の結果」を返して UI/ログを汚さないようにする。
    if (parentKey === null || parentKey === undefined) {
      Logger.log('[listChildFolders] WARN: parentKey is null/undefined (return empty result)');
      return { parentKey: null, nodes: [], nextPageToken: null };
    }
    
    if (typeof parentKey !== 'string') {
      Logger.log('[listChildFolders] ERROR: parentKey is not a string, type=%s', typeof parentKey);
      throw new Error('parentKeyが無効です。parentKey の型: ' + typeof parentKey + ', 値: ' + String(parentKey));
    }
    
    const trimmedKey = parentKey.trim();
    if (trimmedKey === '') {
      Logger.log('[listChildFolders] ERROR: parentKey is empty string');
      throw new Error('parentKeyが無効です。parentKey: 空文字列');
    }
    
    const FOLDER_MIME = 'application/vnd.google-apps.folder';
    const nodes = [];
    let nextPageToken = null;
    
    // 特殊キー処理
    if (parentKey === 'root') {
      // マイドライブ直下
      const rootFolder = DriveApp.getRootFolder();
      const folders = rootFolder.getFolders();
      let count = 0;
      const MAX_FOLDERS = 200;
      
      while (folders.hasNext() && count < MAX_FOLDERS) {
        const folder = folders.next();
        nodes.push({
          id: folder.getId(),
          name: folder.getName(),
          type: 'folder'
        });
        count++;
      }
      
      return { parentKey: 'root', nodes, nextPageToken: null };
    }
    
    if (parentKey === 'sharedDrives') {
      // 共有ドライブ一覧
      try {
        const drives = Drive.Drives.list({
          pageSize: 100,
          pageToken: pageToken || undefined,
          fields: 'drives(id,name),nextPageToken'
        });
        
        (drives.drives || []).forEach(d => {
          nodes.push({
            id: 'drive:' + d.id,
            name: d.name || '(無題の共有ドライブ)',
            type: 'sharedDrive'
          });
        });
        
        return { parentKey: 'sharedDrives', nodes, nextPageToken: drives.nextPageToken || null };
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        Logger.log('[listChildFolders] sharedDrives error: %s', msg);
        // 共有ドライブが取得できない場合（権限不足等）は空配列を返す（UIは落とさない）
        return { parentKey: 'sharedDrives', nodes: [], nextPageToken: null, error: msg };
      }
    }

    if (parentKey === 'sharedWithMe') {
      // 「共有アイテム（自分と共有）」内のフォルダ一覧（仮想ルート）
      try {
        const res = Drive.Files.list({
          q: "sharedWithMe and mimeType = '" + FOLDER_MIME + "' and trashed = false",
          pageSize: 200,
          pageToken: pageToken || undefined,
          supportsAllDrives: true,
          includeItemsFromAllDrives: true,
          fields: 'files(id,name),nextPageToken'
        });

        (res.files || []).forEach(f => {
          nodes.push({
            id: f.id,
            name: f.name || '(無題)',
            type: 'folder'
          });
        });

        return { parentKey: 'sharedWithMe', nodes, nextPageToken: res.nextPageToken || null };
      } catch (e) {
        const msg = e && e.message ? e.message : String(e);
        Logger.log('[listChildFolders] sharedWithMe error: %s', msg);
        return { parentKey: 'sharedWithMe', nodes: [], nextPageToken: null, error: msg };
      }
    }
    
    if (parentKey && parentKey.startsWith('drive:')) {
      // 共有ドライブ直下
      const driveId = parentKey.substring(6);
      try {
        const res = Drive.Files.list({
          q: "'" + driveId + "' in parents and mimeType = '" + FOLDER_MIME + "' and trashed = false",
          pageSize: 200,
          pageToken: pageToken || undefined,
          supportsAllDrives: true,
          includeItemsFromAllDrives: true,
          fields: 'files(id,name),nextPageToken'
        });
        
        (res.files || []).forEach(f => {
          nodes.push({
            id: f.id,
            name: f.name || '(無題)',
            type: 'folder'
          });
        });
        
        return { parentKey, nodes, nextPageToken: res.nextPageToken || null };
      } catch (e) {
        Logger.log('[listChildFolders] drive:%s error: %s', driveId, String(e));
        throw new Error('共有ドライブの取得に失敗しました: ' + (e && e.message ? e.message : String(e)));
      }
    }
    
    // 通常フォルダ直下（parentKey は既に検証済み）
    try {
      const res = Drive.Files.list({
        q: "'" + parentKey + "' in parents and mimeType = '" + FOLDER_MIME + "' and trashed = false",
        pageSize: 200,
        pageToken: pageToken || undefined,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
        fields: 'files(id,name),nextPageToken'
      });
      
      (res.files || []).forEach(f => {
        nodes.push({
          id: f.id,
          name: f.name || '(無題)',
          type: 'folder'
        });
      });
      
      Logger.log('[listChildFolders] folderId="%s" nodes=%s nextToken=%s', parentKey, String(nodes.length), String(res.nextPageToken || ''));
      return { parentKey, nodes, nextPageToken: res.nextPageToken || null };
    } catch (e) {
      const errorMsg = e && e.message ? e.message : String(e);
      Logger.log('[listChildFolders] folderId="%s" error: %s', parentKey, errorMsg);
      if (errorMsg.includes('permission') || errorMsg.includes('権限')) {
        throw new Error('フォルダへのアクセス権限がありません');
      } else if (errorMsg.includes('not found') || errorMsg.includes('見つかりません')) {
        throw new Error('フォルダが見つかりません');
      } else {
        throw new Error('フォルダ一覧の取得に失敗しました: ' + errorMsg);
      }
    }
  } catch (e) {
    const errorMsg = e && e.message ? e.message : String(e);
    Logger.log('[listChildFolders] throw: %s', errorMsg);
    throw new Error('listChildFolders エラー: ' + errorMsg);
  }
}

/**
 * マイドライブのルートフォルダIDを取得
 * @returns {string} ルートフォルダID
 */
function getRootFolderId() {
  try {
    const rootFolder = DriveApp.getRootFolder();
    const rootId = rootFolder.getId();
    Logger.log('[getRootFolderId] rootId="%s"', rootId);
    return rootId;
  } catch (e) {
    const errorMsg = e && e.message ? e.message : String(e);
    Logger.log('[getRootFolderId] error: %s', errorMsg);
    throw new Error('マイドライブのルートフォルダID取得に失敗しました: ' + errorMsg);
  }
}

/**
 * デフォルト設定を返す
 * @returns {Object} デフォルト設定オブジェクト
 */
function getDefaultSettings_() {
  return {
    favorites: [],
    defaultSort: 'modifiedTimeDesc',
    limit: 0, // 0 = 制限なし
    historyEnabled: false,
    history: [],
    theme: 'solarizedLight'
  };
}

/**
 * ユーザー設定を取得（UserPropertiesから読み込み）
 * @returns {Object} 設定オブジェクト
 */
function getSettings() {
  try {
    const props = PropertiesService.getUserProperties();
    const raw = props.getProperty('FILE_NAVI_SETTINGS');
    
    if (!raw) {
      Logger.log('[getSettings] no saved settings, returning defaults');
      return getDefaultSettings_();
    }
    
    try {
      const parsed = JSON.parse(raw);
      // 検証：必須フィールドが存在するか確認
      const validated = {
        favorites: Array.isArray(parsed.favorites) ? parsed.favorites : [],
        defaultSort: (parsed.defaultSort && ['modifiedTimeDesc', 'modifiedTimeAsc', 'nameAsc', 'nameDesc'].includes(parsed.defaultSort)) 
          ? parsed.defaultSort : 'modifiedTimeDesc',
        limit: (typeof parsed.limit === 'number' && parsed.limit >= 0) ? parsed.limit : 0,
        historyEnabled: typeof parsed.historyEnabled === 'boolean' ? parsed.historyEnabled : false,
        history: Array.isArray(parsed.history) ? parsed.history : [],
        theme: (parsed.theme && ['solarizedLight', 'light', 'dark'].includes(parsed.theme)) 
          ? parsed.theme : 'solarizedLight'
      };
      Logger.log('[getSettings] loaded settings: sort=%s limit=%s theme=%s', validated.defaultSort, validated.limit, validated.theme);
      return validated;
    } catch (parseError) {
      Logger.log('[getSettings] JSON parse error: %s, returning defaults', parseError);
      return getDefaultSettings_();
    }
  } catch (e) {
    const errorMsg = e && e.message ? e.message : String(e);
    Logger.log('[getSettings] error: %s', errorMsg);
    return getDefaultSettings_();
  }
}

/**
 * ユーザー設定を保存（UserPropertiesにJSONで保存）
 * @param {Object|string} settings - 設定オブジェクトまたはJSON文字列
 * @returns {Object} 保存された設定オブジェクト
 */
function saveSettings(settings) {
  try {
    let settingsObj;
    if (typeof settings === 'string') {
      try {
        settingsObj = JSON.parse(settings);
      } catch (e) {
        throw new Error('設定のJSON解析に失敗しました: ' + (e && e.message ? e.message : String(e)));
      }
    } else if (typeof settings === 'object' && settings !== null) {
      settingsObj = settings;
    } else {
      throw new Error('設定が無効です。オブジェクトまたはJSON文字列を指定してください。');
    }
    
    // 検証と正規化
    const validated = {
      favorites: Array.isArray(settingsObj.favorites) ? settingsObj.favorites : [],
      defaultSort: (settingsObj.defaultSort && ['modifiedTimeDesc', 'modifiedTimeAsc', 'nameAsc', 'nameDesc'].includes(settingsObj.defaultSort)) 
        ? settingsObj.defaultSort : 'modifiedTimeDesc',
      limit: (typeof settingsObj.limit === 'number' && settingsObj.limit >= 0) ? settingsObj.limit : 0,
      historyEnabled: typeof settingsObj.historyEnabled === 'boolean' ? settingsObj.historyEnabled : false,
      history: Array.isArray(settingsObj.history) ? settingsObj.history : [],
      theme: (settingsObj.theme && ['solarizedLight', 'light', 'dark'].includes(settingsObj.theme)) 
        ? settingsObj.theme : 'solarizedLight'
    };
    
    // 履歴の最大件数制限（100件まで）
    const MAX_HISTORY = 100;
    if (validated.history.length > MAX_HISTORY) {
      validated.history = validated.history.slice(0, MAX_HISTORY);
    }
    
    const props = PropertiesService.getUserProperties();
    props.setProperty('FILE_NAVI_SETTINGS', JSON.stringify(validated));
    Logger.log('[saveSettings] saved: sort=%s limit=%s theme=%s favorites=%s', 
      validated.defaultSort, validated.limit, validated.theme, validated.favorites.length);
    
    return validated;
  } catch (e) {
    const errorMsg = e && e.message ? e.message : String(e);
    Logger.log('[saveSettings] error: %s', errorMsg);
    throw new Error('設定の保存に失敗しました: ' + errorMsg);
  }
}

/**
 * アプリ情報を取得（ヘルプ用）
 * @returns {Object} { clientVersion, serverVersion }
 */
function getAppInfo() {
  try {
    return {
      clientVersion: '2025-12-23-1', // Index.htmlのwindow.__fileNavi.versionと同期
      serverVersion: '2025-12-23-1'
    };
  } catch (e) {
    Logger.log('[getAppInfo] error: %s', e);
    return { clientVersion: 'unknown', serverVersion: 'unknown' };
  }
}

/**
 * フォルダIDからフォルダ名を取得（お気に入り追加時用）
 * @param {string} folderId - フォルダID
 * @returns {string} フォルダ名
 */
function getFolderName(folderId) {
  try {
    if (!folderId || typeof folderId !== 'string' || folderId.trim() === '') {
      throw new Error('フォルダIDが無効です');
    }
    const folder = DriveApp.getFolderById(folderId.trim());
    const name = folder.getName();
    Logger.log('[getFolderName] folderId="%s" name="%s"', folderId, name);
    return name;
  } catch (e) {
    const errorMsg = e && e.message ? e.message : String(e);
    Logger.log('[getFolderName] error: %s', errorMsg);
    throw new Error('フォルダ名の取得に失敗しました: ' + errorMsg);
  }
}