
'use strict';

const $ = id => document.getElementById(id);
const scopes = ['User.Read', 'Files.Read'];

const config = {
    auth: {
        clientId: 'e2cb54f1-a84b-471a-8bcf-adf024a95ff3',
        authority: 'https://login.microsoftonline.com/consumers',
        redirectUri: 'https://ljmead-web.github.io/metal-inventory-app/'
    },
    cache: {
        cacheLocation: 'sessionStorage'
    }
};

const path = 'https://graph.microsoft.com/v1.0/me/drive/root:/Documents/Metal%20Inventory%203.xlsx';

const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
}[c]));

const fmt = (v, n = 2) =>
    typeof v === 'number' && Number.isFinite(v)
        ? v.toLocaleString('en-US', {
            minimumFractionDigits: n,
            maximumFractionDigits: n
        })
        : '—';

const money = v =>
    typeof v === 'number' && Number.isFinite(v)
        ? v.toLocaleString('en-US', {
            style: 'currency',
            currency: 'USD'
        })
        : '—';

const pct = v =>
    typeof v === 'number' ? fmt(v * 100, 1) + '%' : '—';

const line = (a, b) =>
    `<div class="line">
        <span>${esc(a)}</span>
        <strong>${esc(b)}</strong>
    </div>`;

let auth, account, products = [];

function status(s) {
    $('status').textContent = s;
}

function controls(busy = false) {
    $('login').hidden = !!account;
    $('logout').hidden = !account;
    $('refresh').disabled = !account || busy;
    $('login').disabled = busy;
    $('logout').disabled = busy;
}

function render(wb) {
    const d = wb.Sheets['Enhanced Dashboard'];
    const g = wb.Sheets.Guide;

    if (!d || !g) {
        throw Error('Expected Enhanced Dashboard and Guide worksheets.');
    }

    const val = (s, c) =>
        s[c]?.t === 'e' ? null : s[c]?.v;

    const v = c => val(d, c);

    if (typeof v('C6') !== 'number' ||
        typeof v('C7') !== 'number') {
        throw Error('Saved dashboard values missing. Recalculate and save in Excel.');
    }

    const total = v('C6') + v('C7');

    const goal = (name, actual, target) =>
        `<div class="panel">
            ${line(name, pct(target > 0 ? actual / target : null))}
            <progress max="1"
                value="${target > 0
                    ? Math.max(0, Math.min(1, actual / target))
                    : 0}">
            </progress>
            ${line('Goal', name.includes('value')
                ? money(target)
                : fmt(target, 2) + ' oz')}
        </div>`;

    $('holdings').innerHTML =
        `<div class="panel">
            <small>Gold + silver market value</small>
            <div class="big">${money(total)}</div>
        </div>` +

        [
            ['Gold', 'B2', 'B7', 'C7', 'B11'],
            ['Silver', 'B3', 'B6', 'C6', 'B10']
        ].map(([n, s, w, m, a]) =>
            `<div class="panel">
                <h2>${n}</h2>
                ${line('Market value', money(v(m)))}
                ${line('Holdings', fmt(v(w), 4) + ' oz')}
                ${line('Spot / oz', money(v(s)))}
                ${line('Average paid / oz', money(v(a)))}
            </div>`
        ).join('') +

        line('Gold / silver ratio', fmt(v('B5'))) +
        goal('Market value goal', total, v('G3')) +
        goal('Silver goal', v('B6'), v('G5')) +
        goal('Gold goal', v('B7'), v('G7'));

    $('activity').innerHTML =
        `<div class="panel">
            <h2>Purchases · ${esc(v('B32'))}</h2>
            <div class="big">${money(v('F46'))}</div>
            ${line('Gold acquired', fmt(v('B46'), 4) + ' oz')}
            ${line('Gold spent', money(v('C46')))}
            ${line('Silver acquired', fmt(v('D46'), 4) + ' oz')}
            ${line('Silver spent', money(v('E46')))}
        </div>

        <div class="panel scroll">
            <table>
                <thead>
                    <tr>
                        ${[
                            'Month',
                            'Gold oz',
                            'Gold cost',
                            'Silver oz',
                            'Silver cost',
                            'Total'
                        ].map(x => `<th>${x}</th>`).join('')}
                    </tr>
                </thead>
                <tbody>
                    ${Array.from({length: 13}, (_, i) => i + 34)
                        .map(r => `
                            <tr>
                                <td>${esc(v('A' + r))}</td>
                                <td>${fmt(v('B' + r), 4)}</td>
                                <td>${money(v('C' + r))}</td>
                                <td>${fmt(v('D' + r), 4)}</td>
                                <td>${money(v('E' + r))}</td>
                                <td>${money(v('F' + r))}</td>
                            </tr>
                        `).join('')}
                </tbody>
            </table>
        </div>

        <small>Year follows the selection saved in Excel.</small>`;

    products = XLSX.utils.sheet_to_json(g, {
        header: 1,
        range: 3,
        defval: null,
        raw: true
    }).filter(r =>
        ['Gold', 'Silver'].includes(r[0]) && r[1]
    );

    renderGuide();
}

function renderGuide() {
    const q = $('search').value.toLowerCase();

    $('products').innerHTML = products
        .filter(r =>
            r.slice(0, 2).join(' ').toLowerCase().includes(q)
        )
        .map(r => `
            <div class="panel">
                <small>${esc(r[0])} · ${fmt(r[2], 5)} ozt</small>
                <h3>${esc(r[1])}</h3>

                <div class="grid">
                    <div>
                        <small>Melt</small>
                        <strong>${money(r[3])}</strong>
                    </div>

                    <div>
                        <small>Strong buy</small>
                        <strong class="good">${money(r[4])}</strong>
                        <small>${pct(r[6])}</small>
                    </div>

                    <div>
                        <small>Acceptable</small>
                        <strong>${money(r[5])}</strong>
                        <small>${pct(r[7])}</small>
                    </div>
                </div>
            </div>
        `).join('') || '<p>No matching products.</p>';
}

async function refresh() {
    if (!account) return;

    controls(true);
    status('Reading OneDrive…');

    try {
        let token;

        try {
            token = await auth.acquireTokenSilent({
                account,
                scopes
            });
        } catch (e) {
            if (e instanceof msal.InteractionRequiredAuthError) {
                await auth.acquireTokenRedirect({
                    account,
                    scopes
                });
                return;
            }

            throw e;
        }

        const res = await fetch(
            path + '?$select=name,lastModifiedDateTime,@microsoft.graph.downloadUrl',
            {
                headers: {
                    Authorization: 'Bearer ' + token.accessToken
                },
                cache: 'no-store'
            }
        );

        if (!res.ok) {
            throw Error(
                'OneDrive lookup failed (HTTP ' +
                res.status +
                '). Check file path and permission.'
            );
        }

        const meta = await res.json();
        const url = meta['@microsoft.graph.downloadUrl'];

        if (!url) {
            throw Error('OneDrive did not provide a download URL.');
        }

        const file = await fetch(url, {
            cache: 'no-store',
            credentials: 'omit'
        });

        if (!file.ok) {
            throw Error(
                'Workbook download failed (HTTP ' +
                file.status + ').'
            );
        }

        render(XLSX.read(await file.arrayBuffer(), {
            type: 'array'
        }));

        status(
            'Loaded ' + new Date().toLocaleString() +
            ' · Excel saved ' +
            (meta.lastModifiedDateTime
                ? new Date(meta.lastModifiedDateTime).toLocaleString()
                : 'unknown')
        );

    } catch (e) {
        status(
            'Refresh failed: ' + e.message +
            ' — any displayed values may be old.'
        );
    } finally {
        controls();
    }
}

async function init() {
    try {
        if (typeof msal === 'undefined' ||
            typeof XLSX === 'undefined') {
            throw Error('Required libraries did not load.');
        }

        auth = new msal.PublicClientApplication(config);

        await auth.initialize();

        const result = await auth.handleRedirectPromise();

        account =
            result?.account ||
            auth.getActiveAccount() ||
            auth.getAllAccounts()[0];

        if (account) {
            auth.setActiveAccount(account);
            await refresh();
        } else {
            status('Sign in to access your private workbook.');
        }

    } catch (e) {
        status('Initialization failed: ' + e.message);
    } finally {
        controls();
    }
}

$('login').onclick = async () => {
    try {
        await auth.loginRedirect({
            scopes,
            prompt: 'select_account'
        });
    } catch (e) {
        status(e.message);
    }
};

$('refresh').onclick = refresh;

$('logout').onclick = async () => {
    try {
        await auth.logoutRedirect({
            account,
            postLogoutRedirectUri: config.auth.redirectUri
        });
    } catch (e) {
        status(e.message);
    }
};

$('search').oninput = renderGuide;

document.querySelectorAll('[data-tab]').forEach(b =>
    b.onclick = () => {
        document.querySelectorAll('[data-tab]').forEach(x => {
            x.classList.toggle('active', x === b);
            $(x.dataset.tab).hidden = x !== b;
        });
    }
);

init();
