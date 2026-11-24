const gasUrl = 'https://script.google.com/macros/s/AKfycbzkTdgX39OsESb5rOkq6GdjJgEYy1zCH0dwtdMD8vtCrNDYZSlcR48dHQf8zxdpgvLjTA/exec';
const LIFF_ID = '2012765015-hIKLZ2oX';

// ?mock=1 を付けて開くと、LINE と GAS を使わずに画面だけ試せる（動作確認用）
const MOCK = new URLSearchParams(location.search).get('mock') === '1';

let globalMaterialList = [];
let siteMasterData = [];
let currentStaff = null;     // { name, company, phone }
let lastSavedJson = '';      // 最後に一時保存した内容（変更があるかの判定用）
let submitted = false;

document.addEventListener('DOMContentLoaded', () => {
    setToday();
    addExtraRow();
    startApp();
});

function setToday() {
    const now = new Date();
    const y = now.getFullYear(), m = ("0" + (now.getMonth() + 1)).slice(-2), d = ("0" + now.getDate()).slice(-2);
    const dateInput = document.getElementById('date');
    if (dateInput) dateInput.value = `${y}-${m}-${d}`;
}

/* ===================== 起動・LINEログイン ===================== */

async function startApp() {
    try {
        if (!MOCK) {
            setLoading('LINEに接続中...');
            await liff.init({ liffId: LIFF_ID });
            if (!liff.isLoggedIn()) { liff.login({ redirectUri: location.href }); return; }
        }
        setLoading('読み込み中...');
        const [, me] = await Promise.all([fetchLists(), callApi('getMe')]);
        currentStaff = me.staff;
        if (!currentStaff) {
            showRegister(false);
        } else {
            showForm();
            let pending = null;
            try { pending = JSON.parse(localStorage.getItem('pendingDraft')); } catch (e) {}
            if (pending) {
                // ログインし直す前の入力内容を戻して、サーバーにも保存しておく
                localStorage.removeItem('pendingDraft');
                applyDraft(pending);
                await saveDraft(true);
                toast('ログインし直しました。入力内容は残っています');
            } else if (me.draft) {
                showDraftBanner(me.draft);
            }
        }
    } catch (e) {
        console.error(e);
        setLoading('読み込みに失敗しました。電波の良い場所で開き直してください。');
    }
}

function getIdToken() {
    return MOCK ? 'mock' : liff.getIDToken();
}

/** GAS に action 付きで送り、JSON を受け取る */
async function callApi(action, extra = {}) {
    if (MOCK) return mockApi(action, extra);
    const res = await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // プリフライトを起こさない
        body: JSON.stringify(Object.assign({ action, idToken: getIdToken() }, extra))
    });
    const data = await res.json();
    if (!data.ok) {
        // ログインの有効期限切れ → 1回だけログインし直す
        if (data.error === 'invalid_token' && !sessionStorage.getItem('relogin')) {
            sessionStorage.setItem('relogin', '1');
            // 入力中の内容を消さないよう、ログインし直す前に端末に一時退避する
            if (document.getElementById('reportForm').style.display !== 'none') {
                try { localStorage.setItem('pendingDraft', JSON.stringify(collectDraft())); } catch (e) {}
                alert('ログインの有効期限が切れたため、LINEにログインし直します。入力内容はそのまま残ります。');
            }
            liff.logout();
            liff.login({ redirectUri: location.href });
            await new Promise(() => {}); // 画面遷移を待つ
        }
        throw new Error(data.error || 'api_error');
    }
    sessionStorage.removeItem('relogin');
    return data;
}

/* ===================== 画面切り替え ===================== */

function setLoading(text) {
    document.getElementById('loadingView').style.display = 'block';
    document.getElementById('registerView').style.display = 'none';
    document.getElementById('reportForm').style.display = 'none';
    document.getElementById('loadingText').innerText = text;
}

function showRegister(isEdit) {
    document.getElementById('loadingView').style.display = 'none';
    document.getElementById('reportForm').style.display = 'none';
    document.getElementById('registerView').style.display = 'block';
    document.getElementById('registerLead').innerHTML = isEdit
        ? '登録情報を変更できます。'
        : 'はじめての方は、スタッフ登録をお願いします。<br>次回から記入者が自動で入ります。';
    document.getElementById('registerCancelBtn').style.display = isEdit ? 'block' : 'none';
    if (currentStaff) {
        document.getElementById('regName').value = currentStaff.name;
        document.getElementById('regCompany').value = currentStaff.company;
        document.getElementById('regPhone').value = currentStaff.phone;
    }
    window.scrollTo(0, 0);
}

function showForm() {
    document.getElementById('loadingView').style.display = 'none';
    document.getElementById('registerView').style.display = 'none';
    document.getElementById('reportForm').style.display = 'block';
    document.getElementById('record').value = currentStaff.name;
    document.getElementById('staffLabel').innerText = `${currentStaff.name}（${currentStaff.company}）`;
    if (!lastSavedJson) lastSavedJson = JSON.stringify(collectDraft());
}

document.getElementById('registerBtn').onclick = async () => {
    const staff = {
        name: document.getElementById('regName').value.trim(),
        company: document.getElementById('regCompany').value.trim(),
        phone: document.getElementById('regPhone').value.trim()
    };
    if (!staff.name || !staff.company || !staff.phone) { alert('名前・企業名・電話番号をすべて入力してください'); return; }
    if (!/^[0-9０-９\-－ー()（）+\s]{10,}$/.test(staff.phone)) { alert('電話番号を確認してください'); return; }

    const btn = document.getElementById('registerBtn');
    btn.disabled = true; btn.innerText = '登録中...';
    try {
        const res = await callApi('registerStaff', { staff });
        currentStaff = res.staff;
        showForm();
        toast('スタッフ登録が完了しました');
    } catch (e) {
        alert('登録に失敗しました。もう一度お試しください。');
    } finally {
        btn.disabled = false; btn.innerText = '登録する';
    }
};

document.getElementById('editStaffBtn').onclick = () => showRegister(true);
document.getElementById('registerCancelBtn').onclick = () => showForm();

/* ===================== 一時保存（下書き） ===================== */

/** 今の入力内容をまるごと取り出す */
function collectDraft() {
    const v = (id) => document.getElementById(id).value;
    return {
        v: 1,
        fields: {
            date: v('date'), startTime: v('startTime'), endTime: v('endTime'), weather: v('weather'),
            site: v('site'), otherSiteInput: v('otherSiteInput'),
            director: v('director'), primeContractor: v('primeContractor'), salesStaff: v('salesStaff'),
            memberCount: v('memberCount'), memberName: v('memberName'), support: v('support'), notes: v('notes')
        },
        locations: Array.from(document.querySelectorAll('#locationContainer .dynamic-row')).map(r => ({
            loc: r.querySelector('.loc-input').value,
            qty: r.querySelector('.loc-qty').value,
            unit: r.querySelector('.loc-unit').value
        })),
        materials: Array.from(document.querySelectorAll('#materialContainer .dynamic-row')).map(r => ({
            sel: r.querySelector('.mat-input').value,
            other: r.querySelector('.other-mat-name').value,
            qty: r.querySelector('.mat-qty').value
        })),
        extras: Array.from(document.querySelectorAll('#extraContainer .dynamic-row')).map(r => ({
            item: r.querySelector('.extra-item').value,
            qty: r.querySelector('.extra-qty').value
        }))
    };
}

/** 下書きを画面に戻す */
function applyDraft(d) {
    const f = d.fields || {};
    const set = (id, val) => { const el = document.getElementById(id); if (el && val !== undefined && val !== null) el.value = val; };

    ['date', 'startTime', 'endTime', 'weather', 'director', 'primeContractor', 'salesStaff',
     'memberCount', 'memberName', 'support', 'notes'].forEach(id => set(id, f[id]));

    // 現場：一覧から消えた現場は「その他」の直接入力に入れる
    const siteSelect = document.getElementById('site');
    const otherInput = document.getElementById('otherSiteInput');
    let site = f.site || '', other = f.otherSiteInput || '';
    if (site && site !== 'OTHER_SITE' && !hasOption(siteSelect, site)) { other = site; site = 'OTHER_SITE'; }
    siteSelect.value = site;
    otherInput.value = other;
    otherInput.style.display = site === 'OTHER_SITE' ? 'block' : 'none';
    otherInput.required = site === 'OTHER_SITE';

    // 施工内容（固定6行）
    const locRows = document.querySelectorAll('#locationContainer .dynamic-row');
    (d.locations || []).forEach((l, i) => {
        if (!locRows[i]) return;
        locRows[i].querySelector('.loc-input').value = l.loc || '';
        locRows[i].querySelector('.loc-qty').value = l.qty || '';
        locRows[i].querySelector('.loc-unit').value = l.unit || '';
    });

    // 使用材料
    const matBox = document.getElementById('materialContainer');
    matBox.innerHTML = '';
    const mats = (d.materials && d.materials.length) ? d.materials : [{}];
    mats.forEach(m => {
        addDynamicRow('materialContainer', 'mat-input');
        const row = matBox.lastElementChild;
        const select = row.querySelector('.mat-input');
        let sel = m.sel || '', other = m.other || '';
        if (sel && sel !== 'OTHER_MAT' && !hasOption(select, sel)) { other = sel; sel = 'OTHER_MAT'; }
        select.value = sel;
        checkMatOther(select);
        row.querySelector('.other-mat-name').value = other;
        row.querySelector('.mat-qty').value = m.qty || '';
    });

    // その他（経費・備品）
    const extraBox = document.getElementById('extraContainer');
    extraBox.innerHTML = '';
    const extras = (d.extras && d.extras.length) ? d.extras : [{}];
    extras.forEach(x => {
        addExtraRow();
        const row = extraBox.lastElementChild;
        row.querySelector('.extra-item').value = x.item || '';
        row.querySelector('.extra-qty').value = x.qty || '';
    });

    showStep(1);
}

function hasOption(select, value) {
    return Array.from(select.options).some(o => o.value === value);
}

function showDraftBanner(draft) {
    const banner = document.getElementById('draftBanner');
    banner.innerHTML = `
        <p>📝 ${draft.savedAt} に一時保存した日報があります。</p>
        <div class="form-row">
            <button type="button" id="discardDraftBtn" class="secondary-btn">破棄する</button>
            <button type="button" id="resumeDraftBtn" class="primary-btn flex-grow">続きから入力</button>
        </div>`;
    banner.style.display = 'block';

    document.getElementById('resumeDraftBtn').onclick = () => {
        applyDraft(draft.data);
        lastSavedJson = JSON.stringify(collectDraft());
        banner.style.display = 'none';
        toast('一時保存した内容を読み込みました');
    };
    document.getElementById('discardDraftBtn').onclick = async () => {
        if (!confirm('一時保存した日報を破棄します。よろしいですか？')) return;
        banner.style.display = 'none';
        try { await callApi('deleteDraft'); } catch (e) { /* 次回また表示されるだけなので無視 */ }
    };
}

async function saveDraft(silent) {
    const draft = collectDraft();
    const json = JSON.stringify(draft);
    if (silent && json === lastSavedJson) return; // 変更なし
    const btns = document.querySelectorAll('.draft-btn');
    if (!silent) btns.forEach(b => { b.disabled = true; b.innerText = '保存中...'; });
    try {
        const res = await callApi('saveDraft', { draft });
        lastSavedJson = json;
        document.getElementById('draftBanner').style.display = 'none';
        if (!silent) toast(`一時保存しました（${res.savedAt}）`);
    } catch (e) {
        if (!silent) alert('一時保存に失敗しました。電波の良い場所でもう一度お試しください。');
    } finally {
        btns.forEach(b => { b.disabled = false; b.innerText = '一時保存'; });
    }
}

document.querySelectorAll('.draft-btn').forEach(b => b.onclick = () => saveDraft(false));

// アプリを閉じる・切り替える時、変更があれば自動で一時保存する
document.addEventListener('visibilitychange', () => {
    if (document.visibilityState !== 'hidden' || submitted || !currentStaff) return;
    if (document.getElementById('reportForm').style.display === 'none') return;
    const draft = collectDraft();
    const json = JSON.stringify(draft);
    if (json === lastSavedJson) return;
    // 下書きバナーが出たまま（まだ読み込むか決めていない）なら、上書きしない
    if (document.getElementById('draftBanner').style.display !== 'none') return;
    lastSavedJson = json;
    if (MOCK) { mockApi('saveDraft', { draft }); return; }
    fetch(gasUrl, {
        method: 'POST', keepalive: true,
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'saveDraft', idToken: getIdToken(), draft })
    }).catch(() => {});
});

function toast(msg) {
    const t = document.getElementById('toast');
    t.innerText = msg;
    t.classList.add('show');
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => t.classList.remove('show'), 2500);
}

/* ===================== ここから下は Ver5.0 の処理をそのまま利用 ===================== */

async function fetchLists() {
    try {
        const res = MOCK ? null : await fetch(gasUrl);
        const data = MOCK ? mockLists() : await res.json();
        siteMasterData = data.sites || [];
        globalMaterialList = data.materials || [];

        const siteSelect = document.getElementById('site');
        if (siteSelect) {
            siteSelect.innerHTML = '<option value="" disabled selected>現場を選択</option>';
            siteMasterData.forEach(s => {
                const opt = document.createElement('option');
                opt.value = opt.textContent = s.name;
                siteSelect.appendChild(opt);
            });
            siteSelect.innerHTML += '<option value="OTHER_SITE">その他</option>';
        }
        addDynamicRow('materialContainer', 'mat-input');
    } catch (e) {
        console.error('データ取得失敗');
    }
}

// 現場名→自動反映
document.getElementById('site').addEventListener('change', (e) => {
    const otherInput = document.getElementById('otherSiteInput');
    const info = siteMasterData.find(s => s.name === e.target.value);

    if (e.target.value === "OTHER_SITE") {
        if(otherInput) {
            otherInput.style.display = 'block';
            otherInput.required = true;
        }
        document.getElementById('director').value = "";
        document.getElementById('primeContractor').value = "";
        document.getElementById('salesStaff').value = "";
    } else {
        if(otherInput) {
            otherInput.style.display = 'none';
            otherInput.required = false;
            otherInput.value = "";
        }
        if (info) {
            document.getElementById('director').value = info.director || "";
            document.getElementById('primeContractor').value = info.prime || "";
            document.getElementById('salesStaff').value = info.staff || "";
            document.getElementById('memberCount').value = info.Count || "";
            document.getElementById('memberName').value = info.Name || "";
        }
    }
});

// 材料行の追加
function addDynamicRow(containerId, inputClass) {
    const container = document.getElementById(containerId);
    const div = document.createElement('div');
    div.className = 'dynamic-row';
    div.style.flexWrap = 'wrap';

    let options = '<option value="" disabled selected>材料を選択</option>';
    globalMaterialList.forEach(m => { options += `<option value="${m}">${m}</option>`; });
    options += '<option value="OTHER_MAT">その他（直接入力）</option>';

    div.innerHTML = `
        <select class="${inputClass} flex-grow" onchange="checkMatOther(this)">${options}</select>
        <input type="text" class="other-mat-name" style="display:none; flex:1 1 100%; margin-top:5px;" placeholder="材料名を入力">
        <input type="text" style="width: 60px;" class="mat-qty" placeholder="数量">
        <button type="button" class="remove-btn">×</button>
    `;
    container.appendChild(div);
    updateRemoveButtons(containerId);
}

// その他（経費・備品など）の行追加
function addExtraRow() {
    const container = document.getElementById('extraContainer');
    if(!container) return;
    const div = document.createElement('div');
    div.className = 'dynamic-row';
    div.innerHTML = `
        <input type="text" class="extra-item flex-grow" placeholder="材料名">
        <input type="text" style="width: 80px;" class="extra-qty" placeholder="数量">
        <button type="button" class="remove-btn">×</button>
    `;
    container.appendChild(div);

    const updateExtraBtns = () => {
        const rows = container.querySelectorAll('.dynamic-row');
        rows.forEach(row => {
            const btn = row.querySelector('.remove-btn');
            btn.onclick = () => { if (rows.length > 1) row.remove(); updateExtraBtns(); };
            btn.disabled = (rows.length === 1);
            btn.style.opacity = (rows.length === 1) ? "0.3" : "1";
        });
    };
    updateExtraBtns();
}

document.getElementById('addExtraBtn').onclick = addExtraRow;

function checkMatOther(select) {
    const otherInput = select.parentElement.querySelector('.other-mat-name');
    if (select.value === "OTHER_MAT") {
        otherInput.style.display = 'block';
    } else {
        otherInput.style.display = 'none';
        otherInput.value = "";
    }
}

function updateRemoveButtons(containerId) {
    const rows = document.getElementById(containerId).querySelectorAll('.dynamic-row');
    rows.forEach(row => {
        const btn = row.querySelector('.remove-btn');
        if (btn) {
            btn.onclick = () => { if (rows.length > 1) row.remove(); updateRemoveButtons(containerId); };
            btn.disabled = (rows.length === 1);
            btn.style.opacity = (rows.length === 1) ? "0.3" : "1";
        }
    });
}

const showStep = (s) => {
    ['step1', 'step2', 'step3'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });
    document.getElementById(`step${s}`).style.display = 'block';
    window.scrollTo(0, 0);
};

// Step1 → 2
// 必須の項目（ここに足せば、送信前のチェックにも入る）
const REQUIRED_FIELDS = [
    { id: 'record', label: '記入者' },
    { id: 'date', label: '日付' },
    { id: 'startTime', label: '開始時刻' },
    { id: 'endTime', label: '終了時刻' },
    { id: 'weather', label: '天気' },
    { id: 'site', label: '現場名' },
    { id: 'memberCount', label: '人工' },
    { id: 'memberName', label: '作業者名' }
];

// まだ入力されていない必須項目を返す
function findMissing() {
    const missing = [];
    REQUIRED_FIELDS.forEach(f => {
        const el = document.getElementById(f.id);
        if (el && !String(el.value).trim()) missing.push({ el, label: f.label });
    });
    const other = document.getElementById('otherSiteInput');
    if (document.getElementById('site').value === 'OTHER_SITE' && other && !other.value.trim()) {
        missing.push({ el: other, label: '現場名（その他）' });
    }
    const locs = Array.from(document.querySelectorAll('.loc-input'));
    if (!locs.some(i => i.value.trim() !== '')) missing.push({ el: locs[0], label: '施工箇所（1つ以上）' });
    return missing;
}

// 未入力の項目を赤くして、1ページ目のその場所へ戻す。未入力が無ければ true
function checkRequired() {
    document.querySelectorAll('.missing').forEach(el => el.classList.remove('missing'));
    const missing = findMissing();
    if (missing.length === 0) return true;
    missing.forEach(m => m.el && m.el.classList.add('missing'));
    alert('まだ入力されていない項目があります。\n\n・' + missing.map(m => m.label).join('\n・') +
          '\n\n1ページ目の赤い枠のところを入力してください。');
    showStep(1);
    if (missing[0].el) missing[0].el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    return false;
}

// 入力したら赤い枠を外す
document.addEventListener('input', e => { if (e.target.classList) e.target.classList.remove('missing'); });
document.addEventListener('change', e => { if (e.target.classList) e.target.classList.remove('missing'); });

// Step1 → 2：未入力があっても進める（送信前の「内容確認へ」で止める）
document.getElementById('nextBtn1').onclick = () => {
    const n = findMissing().length;
    showStep(2);
    if (n > 0) toast(`未入力が${n}件あります。「内容確認へ」で確認します`);
};

// Step2 → 3
document.getElementById('nextBtn2').onclick = () => {
    if (!checkRequired()) return; // 未入力があれば、確認画面には進ませない
    const getVal = (id) => {
        const el = document.getElementById(id);
        return el ? el.value : "";
    };

    const setConf = (id, val) => {
        const el = document.getElementById(id);
        if(el) el.innerText = val;
    };

    setConf('conf-record', getVal('record'));
    setConf('conf-date', getVal('date'));
    setConf('conf-weather', getVal('weather'));
    setConf('conf-time', `${getVal('startTime')} 〜 ${getVal('endTime')}`);

    const siteVal = getVal('site');
    setConf('conf-site', (siteVal === "OTHER_SITE") ? getVal('otherSiteInput') : siteVal);

    setConf('conf-director', getVal('director'));
    setConf('conf-primeContractor', getVal('primeContractor'));
    setConf('conf-salesStaff', getVal('salesStaff'));
    setConf('conf-name', getVal('memberName'));
    setConf('conf-count', getVal('memberCount'));

    const support = getVal('supportMember') || getVal('support');
    setConf('conf-support', support || "なし");
    setConf('conf-notes', getVal('notes') || "なし");

    renderConfirmLists();
    showStep(3);
};

function renderConfirmLists() {
    // 施工箇所（備考含む）
    const locList = document.getElementById('conf-locations-list');
    locList.innerHTML = '<strong>【施工箇所】</strong>';
    document.querySelectorAll('.loc-input').forEach(i => {
        if (i.value) {
            const div = document.createElement('div');
            const parent = i.parentElement;
            const qty = parent.querySelector('.loc-qty').value || 0;
            const unit = parent.querySelector('.loc-unit').value || "";
            const note = parent.querySelector('.loc-notes') ? parent.querySelector('.loc-notes').value : "";
            div.innerText = `・${i.value} ： ${qty}${unit} ${note ? '('+note+')' : ''}`;
            locList.appendChild(div);
        }
    });

    // 材料
    const matList = document.getElementById('conf-materials-list');
    matList.innerHTML = '<strong>【使用材料】</strong>';
    document.querySelectorAll('.mat-input').forEach(select => {
        let name = select.value;
        if (name === "OTHER_MAT") name = select.parentElement.querySelector('.other-mat-name').value;
        if (name && name !== "OTHER_MAT") {
            const div = document.createElement('div');
            const qty = select.parentElement.querySelector('.mat-qty').value || 0;
            div.innerText = `・${name} ： ${qty}`;
            matList.appendChild(div);
        }
    });

    // その他項目の表示
    const extraList = document.getElementById('conf-extra-list') || document.createElement('div');
    extraList.id = 'conf-extra-list';
    extraList.innerHTML = '<strong>【その他項目】</strong>';
    document.querySelectorAll('.extra-item').forEach(input => {
        if (input.value) {
            const div = document.createElement('div');
            const qty = input.parentElement.querySelector('.extra-qty').value || "";
            div.innerText = `・${input.value} ： ${qty}`;
            extraList.appendChild(div);
        }
    });
    matList.after(extraList);
}

document.getElementById('addMaterialBtn').onclick = () => addDynamicRow('materialContainer', 'mat-input');
document.getElementById('backBtn1').onclick = () => showStep(1);
document.getElementById('backBtn2').onclick = () => showStep(2);

// 送信
document.getElementById('reportForm').onsubmit = async function(e) {
    e.preventDefault();
    if (!checkRequired()) return; // 念のため、送信の直前にももう一度確かめる
    const btn = document.getElementById('submitBtn');
    btn.disabled = true; btn.innerText = "送信中...";

    // 送る前にログインが有効か確かめる（切れていたら入力を残したままログインし直す）
    if (!MOCK) {
        try { await callApi('getMe'); }
        catch (err) {
            alert('通信に失敗しました。電波の良い場所でもう一度お試しください。');
            btn.disabled = false; btn.innerText = "この内容で送信する";
            return;
        }
    }

    const getVal = (id) => document.getElementById(id) ? document.getElementById(id).value : "";

    const siteVal = getVal('site');
    const finalSiteName = (siteVal === "OTHER_SITE") ? getVal('otherSiteInput') : siteVal;

    const locArray = Array.from(document.querySelectorAll('.loc-input')).map(v => {
        const p = v.parentElement;
        return {
            loc: v.value,
            qty: p.querySelector('.loc-qty').value || 0,
            unit: p.querySelector('.loc-unit').value || "",
            notes: p.querySelector('.loc-notes') ? p.querySelector('.loc-notes').value : ""
        };
    }).filter(i => i.loc !== "");

    const matArray = Array.from(document.querySelectorAll('.mat-input')).map(select => {
        let name = select.value;
        if (name === "OTHER_MAT") name = select.parentElement.querySelector('.other-mat-name').value;
        return { name: name, qty: select.parentElement.querySelector('.mat-qty').value || 0 };
    }).filter(i => i.name && i.name !== "OTHER_MAT" && i.name !== "");

    const extraArray = Array.from(document.querySelectorAll('.extra-item')).map(input => {
        return { item: input.value, qty: input.parentElement.querySelector('.extra-qty').value || "" };
    }).filter(i => i.item !== "");

    const payload = {
        record: getVal('record'),
        date: getVal('date'),
        startTime: getVal('startTime'),
        endTime: getVal('endTime'),
        weather: getVal('weather'),
        site: finalSiteName,
        director: getVal('director'),
        primeContractor: getVal('primeContractor'),
        salesStaff: getVal('salesStaff'),
        name: getVal('memberName'),
        count: getVal('memberCount'),
        supportMember: getVal('support') || getVal('supportMember'),
        notes: getVal('notes'),
        locationArray: locArray,
        materialArray: matArray,
        extraArray: extraArray,
        idToken: MOCK ? '' : getIdToken()   // 送った本人にPDFをLINEで送るため
    };

    try {
        let result = { result: 'success', line: 'sent' };
        if (!MOCK) {
            const res = await fetch(gasUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(payload)
            });
            result = await res.json();
        }
        if (result.result !== 'success') {
            // 記録に失敗した時は下書きを消さずに、理由を出す
            alert('日報の送信に失敗しました。\n' + (result.message || '') + '\n\n「一時保存」してから、管理者に連絡してください。');
            btn.disabled = false; btn.innerText = "この内容で送信する";
            return;
        }
        submitted = true;
        try { await callApi('deleteDraft'); } catch (err) { /* 下書きが残るだけなので無視 */ }
        if (result.line === 'sent') {
            alert('日報の送信が完了しました！\nLINEに日報PDFをお送りします。');
        } else if (result.line) {
            alert('日報の送信が完了しました！\n\n※ LINEへのPDF送付はできませんでした。\n理由：' + result.line);
        } else {
            alert('日報の送信が完了しました！\n\n※ LINEへのPDF送付の結果が分かりませんでした（GASが古い版のままの可能性があります）。');
        }
        if (!MOCK && liff.isInClient()) {
            liff.closeWindow();
        } else {
            location.reload();
        }
    } catch (e) {
        alert('送信失敗');
        btn.disabled = false; btn.innerText = "この内容で送信する";
    }
};

/* ===================== 動作確認用（?mock=1 の時だけ使う） ===================== */

function mockLists() {
    return {
        sites: [
            { name: '〇〇ビル新築工事', director: '田中', prime: '△△建設', staff: '佐藤', Count: 3, Name: '鈴木、高橋、伊藤' },
            { name: '□□邸改修工事', director: '山本', prime: '自社', staff: '中村', Count: 2, Name: '小林、加藤' }
        ],
        materials: ['コンクリート', '鉄筋 D13', '型枠合板']
    };
}

function mockApi(action, extra) {
    const read = (k) => { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } };
    const write = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} };
    const stamp = () => { const d = new Date(); return `${d.getMonth() + 1}/${d.getDate()} ${('0' + d.getHours()).slice(-2)}:${('0' + d.getMinutes()).slice(-2)}`; };
    switch (action) {
        case 'getMe': return { ok: true, staff: read('mock_staff'), draft: read('mock_draft') };
        case 'registerStaff': write('mock_staff', extra.staff); return { ok: true, staff: extra.staff };
        case 'saveDraft': { const savedAt = stamp(); write('mock_draft', { savedAt, data: extra.draft }); return { ok: true, savedAt }; }
        case 'deleteDraft': localStorage.removeItem('mock_draft'); return { ok: true };
    }
}
