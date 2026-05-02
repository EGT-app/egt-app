import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js';
import { getFirestore, collection, doc, setDoc, getDoc, getDocs, onSnapshot, deleteDoc, query, orderBy, serverTimestamp } from 'https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js';

const FIREBASE_CONFIG = {
  apiKey: "AIzaSyDLweNIl7Vc1a_SUrKiWUriMKTc9MEkex0",
  authDomain: "egt-app-bb6c0.firebaseapp.com",
  projectId: "egt-app-bb6c0",
  storageBucket: "egt-app-bb6c0.firebasestorage.app",
  messagingSenderId: "211765326412",
  appId: "1:211765326412:web:ecef15d9ee892310d4fae9"
};

const fbApp = initializeApp(FIREBASE_CONFIG);
const db = getFirestore(fbApp);
window.EGT_DB = db;

// ── Fonctions globales exposées ──────────────────
window.fbSave = async (col, id, data) => {
  try {
    await setDoc(doc(db, col, String(id)), { ...data, _updatedAt: serverTimestamp() });
  } catch(e) { console.warn('[FB] Save error:', e); }
};

window.fbDelete = async (col, id) => {
  try {
    await deleteDoc(doc(db, col, String(id)));
  } catch(e) { console.warn('[FB] Delete error:', e); }
};

// ── Sync initiale depuis Firebase ────────────────
async function syncInit() {
  try {
    // Chantiers
    const snap = await getDocs(collection(db, 'chantiers'));
    if (!snap.empty) {
      const items = [];
      snap.forEach(d => { const v = d.data(); delete v._updatedAt; items.push(v); });
      DB.chantiers = items;
      console.log('[FB] Chantiers:', items.length);
    }
    // Techniciens
    const snapT = await getDocs(collection(db, 'techniciens'));
    if (!snapT.empty) {
      const items = [];
      snapT.forEach(d => { const v = d.data(); delete v._updatedAt; items.push(v); });
      DB.techniciens = items;
    }
    save();
    if (window.renderChantiers) renderChantiers();
    showFBBadge('● SYNC', true);
  } catch(e) {
    console.warn('[FB] Sync init error:', e);
    showFBBadge('○ LOCAL', false);
  }
}

// ── Listeners temps réel ─────────────────────────
function startListeners() {
  // Chantiers temps réel
  onSnapshot(collection(db, 'chantiers'), snapshot => {
    let changed = false;
    snapshot.docChanges().forEach(change => {
      const data = change.doc.data();
      delete data._updatedAt;
      if (change.type === 'added' || change.type === 'modified') {
        const idx = DB.chantiers.findIndex(c => c.id === data.id);
        if (idx >= 0) { DB.chantiers[idx] = data; }
        else { DB.chantiers.unshift(data); }
        changed = true;
      } else if (change.type === 'removed') {
        DB.chantiers = DB.chantiers.filter(c => c.id !== data.id);
        changed = true;
      }
    });
    if (changed) {
      save();
      if (window.renderChantiers) renderChantiers();
    }
  });

  // Techniciens temps réel
  onSnapshot(collection(db, 'techniciens'), snapshot => {
    snapshot.docChanges().forEach(change => {
      const data = change.doc.data();
      delete data._updatedAt;
      if (change.type === 'added' || change.type === 'modified') {
        const idx = DB.techniciens.findIndex(t => t.id === data.id);
        if (idx >= 0) DB.techniciens[idx] = data;
        else DB.techniciens.push(data);
      } else if (change.type === 'removed') {
        DB.techniciens = DB.techniciens.filter(t => t.id !== data.id);
      }
    });
    save();
  });
}

// ── Badge SYNC dans topbar ────────────────────────
function showFBBadge(text, ok) {
  let badge = document.getElementById('fb-sync-badge');
  if (!badge) {
    badge = document.createElement('span');
    badge.id = 'fb-sync-badge';
    badge.style.cssText = 'font-size:9px;padding:2px 7px;border-radius:10px;font-weight:700;font-family:monospace;cursor:default';
    const tbRight = document.querySelector('.tb-right');
    if (tbRight) tbRight.prepend(badge);
  }
  badge.textContent = text;
  badge.style.background = ok ? 'rgba(0,196,122,.2)' : 'rgba(150,90,150,.2)';
  badge.style.color = ok ? 'var(--green)' : 'var(--accent)';
}

// ── Patch saveChantier ────────────────────────────
window.addEventListener('load', () => {
  setTimeout(() => {
    // Patch saveChantier
    const _origSC = window.saveChantier;
    window.saveChantier = async function(editId) {
      _origSC.apply(this, arguments);
      const c = editId ? DB.chantiers.find(x => x.id === editId) : DB.chantiers[0];
      if (c) await window.fbSave('chantiers', c.id, c);
    };

    // Patch saveTechnicien
    const _origST = window.saveTechnicien;
    if (_origST) window.saveTechnicien = async function() {
      _origST.apply(this, arguments);
      const t = DB.techniciens[DB.techniciens.length - 1];
      if (t) await window.fbSave('techniciens', t.id, t);
    };

    // Patch saveEditTechnicien
    const _origSET = window.saveEditTechnicien;
    if (_origSET) window.saveEditTechnicien = async function(tId) {
      _origSET.apply(this, arguments);
      const t = DB.techniciens.find(x => x.id === tId);
      if (t) await window.fbSave('techniciens', tId, t);
    };

    // Démarrer la sync
    syncInit().then(() => startListeners());
    showFBBadge('⟳ SYNC...', true);

    console.log('[EGT Firebase] Opérationnel ✓');
  }, 800);
});
