// ── Config Azure ────────────────────────────────
const OD_CLIENT_ID = 'd3d97a68-884e-4ac6-948d-11fc509bf593';
const OD_TENANT_ID = '906bb4c2-f82c-4b7e-a6bb-41311b792bfe';
const OD_SCOPES    = ['Files.ReadWrite', 'User.Read'];

window.OD = { connected: false, account: null, token: null, msal: null };

// ── Initialisation MSAL ─────────────────────────
async function odInit() {
  if (!window.msal) {
    console.error('[OD] MSAL non disponible — librairie non chargée');
    return;
  }
  console.log('[OD] MSAL chargé ✓', typeof window.msal);
  try {
    const cfg = {
      auth: {
        clientId: OD_CLIENT_ID,
        authority: 'https://login.microsoftonline.com/' + OD_TENANT_ID,
        redirectUri: window.location.origin + window.location.pathname
      },
      cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
    };
    window.OD.msal = new msal.PublicClientApplication(cfg);
    await window.OD.msal.initialize();

    const accounts = window.OD.msal.getAllAccounts();
    if (accounts.length > 0) {
      window.OD.account = accounts[0];
      try {
        const resp = await window.OD.msal.acquireTokenSilent({ scopes: OD_SCOPES, account: accounts[0] });
        window.OD.token = resp.accessToken;
        window.OD.connected = true;
        console.log('[OD] Reconnecté silencieusement :', accounts[0].username);
      } catch(e) { window.OD.connected = false; }
    }
  } catch(e) { console.warn('[OD] Init error:', e); }
}

// ── Obtenir un token valide ─────────────────────
async function odGetToken() {
  if (!window.OD.msal || !window.OD.account) return null;
  try {
    const r = await window.OD.msal.acquireTokenSilent({ scopes: OD_SCOPES, account: window.OD.account });
    window.OD.token = r.accessToken;
    return r.accessToken;
  } catch(e) {
    try {
      const r = await window.OD.msal.acquireTokenPopup({ scopes: OD_SCOPES });
      window.OD.token = r.accessToken;
      return r.accessToken;
    } catch(e2) { return null; }
  }
}

// ── Appel Graph API ─────────────────────────────
async function odGraph(method, path, body, token, isUpload, uploadBlob, mimeType) {
  const t = token || await odGetToken();
  if (!t) throw new Error('Token indisponible');
  const headers = { 'Authorization': 'Bearer ' + t };
  let bodyData = null;
  if (isUpload && uploadBlob) {
    headers['Content-Type'] = mimeType || 'application/octet-stream';
    bodyData = uploadBlob;
  } else if (body) {
    headers['Content-Type'] = 'application/json';
    bodyData = JSON.stringify(body);
  }
  const r = await fetch('https://graph.microsoft.com/v1.0' + path, { method, headers, body: bodyData });
  if (!r.ok) {
    const txt = await r.text();
    throw new Error('Graph ' + r.status + ': ' + txt);
  }
  return r.status === 204 ? null : r.json();
}

// ── Connexion ───────────────────────────────────
window.odConnect = async function() {
  if (!window.OD.msal) {
    if (window.showNotif) showNotif('Initialisation...', 'success');
    await odInit();
    if (!window.OD.msal) { if (window.showNotif) showNotif('MSAL non disponible', 'error'); return; }
  }
  try {
    const resp = await window.OD.msal.loginPopup({ scopes: OD_SCOPES });
    window.OD.account = resp.account;
    window.OD.token = resp.accessToken;
    window.OD.connected = true;
    if (window.showNotif) showNotif('OneDrive EGT connecté ✓ — ' + resp.account.username, 'success');
    if (window.currentChantier) renderDetailTab('docs');
  } catch(e) {
    console.error('[OD] Login error:', e);
    if (window.showNotif) showNotif('Connexion annulée ou refusée', 'error');
  }
};

// ── Déconnexion ─────────────────────────────────
window.odDisconnect = async function() {
  try {
    if (window.OD.msal && window.OD.account) {
      await window.OD.msal.logoutPopup({ account: window.OD.account });
    }
  } catch(e) {}
  window.OD.connected = false; window.OD.account = null; window.OD.token = null;
  if (window.showNotif) showNotif('OneDrive déconnecté', 'success');
  if (window.currentChantier) renderDetailTab('docs');
};

// ── Créer la structure de dossiers d'un chantier ─
window.odCreateChantierFolders = async function(chantier) {
  if (!chantier) return;
  const token = await odGetToken();
  if (!token) { if (window.showNotif) showNotif('Connectez-vous à OneDrive d\'abord', 'error'); return; }

  if (window.showNotif) showNotif('⏳ Création des dossiers OneDrive...', 'success');

  const safe = s => (s || '').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-').trim().slice(0, 80);
  const folderName = safe(chantier.id + ' — ' + (chantier.nom || chantier.adresse || 'Chantier'));

  try {
    // 1. Dossier racine "Chantiers EGT"
    await odGraph('POST', '/me/drive/root/children', {
      name: 'Chantiers EGT', folder: {}, '@microsoft.graph.conflictBehavior': 'fail'
    }, token).catch(() => {});

    // 2. Dossier du chantier
    const folder = await odGraph('POST', '/me/drive/root:/Chantiers EGT:/children', {
      name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename'
    }, token);

    // 3. Sous-dossiers
    for (const sf of ['Photos', 'Documents', 'DFT', 'Securite']) {
      await odGraph('POST', '/me/drive/items/' + folder.id + '/children', {
        name: sf, folder: {}, '@microsoft.graph.conflictBehavior': 'fail'
      }, token).catch(() => {});
    }

    // 4. Mémoriser le chemin dans le chantier
    chantier.onedrivePath = 'Chantiers EGT/' + folderName;
    chantier.onedriveId   = folder.id;
    save();
    if (window.fbSave) window.fbSave('chantiers', String(chantier.id), chantier);

    if (window.showNotif) showNotif('📁 Dossier OneDrive créé ✓', 'success');
    if (window.currentChantier && window.currentChantier.id === chantier.id) renderDetailTab('docs');
    return folder;
  } catch(e) {
    console.error('[OD] Folder error:', e);
    if (window.showNotif) showNotif('Erreur dossier : ' + e.message, 'error');
  }
};

// ── Upload photo vers OneDrive ──────────────────
window.odUploadPhoto = async function(chantierId, fileName, base64Data, subfolder) {
  const chantier = DB.chantiers.find(c => c.id == chantierId);
  if (!chantier || !chantier.onedriveId) return;
  const token = await odGetToken();
  if (!token) return;
  subfolder = subfolder || 'Photos';
  try {
    const parts   = base64Data.split(',');
    const mime    = (parts[0].match(/:(.*?);/) || [])[1] || 'image/jpeg';
    const ext     = mime.split('/')[1] || 'jpg';
    const binary  = atob(parts[1] || parts[0]);
    const ab      = new ArrayBuffer(binary.length);
    const ia      = new Uint8Array(ab);
    for (let i = 0; i < binary.length; i++) ia[i] = binary.charCodeAt(i);
    const blob    = new Blob([ab], { type: mime });
    const safeName = fileName.replace(/[<>:"/\\|?*]/g, '-').slice(0, 60) + '.' + ext;
    const path = '/me/drive/items/' + chantier.onedriveId + ':/' + subfolder + '/' + encodeURIComponent(safeName) + ':/content';
    await odGraph('PUT', path, null, token, true, blob, mime);
    console.log('[OD] Photo uploadée :', safeName);
  } catch(e) { console.warn('[OD] Upload photo error:', e); }
};

// ── Upload document vers OneDrive ───────────────
window.odUploadDoc = async function(chantierId, fileName, base64Data) {
  return window.odUploadPhoto(chantierId, fileName, base64Data, 'Documents');
};

// ── Init automatique au chargement ─────────────
window.addEventListener('load', () => setTimeout(odInit, 1500));
