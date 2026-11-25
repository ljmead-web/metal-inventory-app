
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
    typeof v === 'number' && Number.isFinite(v)
        ? fmt(v * 100, 1) + '%'
        : '—';

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

    function goalRow(name, actual, target, currency = false) {

        const ratio =
            typeof actual === 'number' &&
            typeof target === 'number' &&
            target > 0
                ? actual / target
                : null;

        const progress = ratio === null
            ? 0
            : Math.max(0, Math.min(1, ratio));

        return `
            <tr>
                <td>${esc(name)}</td>
                <td>${currency
                    ? money(target)
                    : fmt(target, 0) + ' oz'}</td>
                <td class="goal-progress">
                    ${pct(ratio)}
                    <progress max="1" value="${progress}">
                    </progress>
                </td>
            </tr>`;
    }

    // COMPACT HOLDINGS TABLE

    $('holdings').innerHTML = `

        <div class="total-value">
            <strong>Total market value</strong>
            <span class="big">${money(total)}</span>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Holdings</th>
                    <th class="gold">Gold</th>
                    <th class="silver">Silver</th>
                </tr>
            </thead>

            <tbody>
                <tr>
                    <td>Spot price</td>
                    <td>${money(v('B2'))}</td>
                    <td>${money(v('B3'))}</td>
                </tr>

                <tr>
                    <td>Ounces held</td>
                    <td>${fmt(v('B7'), 4)}</td>
                    <td>${fmt(v('B6'), 2)}</td>
                </tr>

                <tr>
                    <td>Market value</td>
                    <td>${money(v('C7'))}</td>
                    <td>${money(v('C6'))}</td>
                </tr>

                <tr>
                    <td>Avg. paid / oz</td>
                    <td>${money(v('B11'))}</td>
                    <td>${money(v('B10'))}</td>
                </tr>

                <tr>
                    <td>Gold / Silver ratio</td>
                    <td colspan="2">${fmt(v('B5'))}</td>
                </tr>
            </tbody>
        </table>

        <table>
            <thead>
                <tr>
                    <th>Goals</th>
                    <th>Target</th>
                    <th>Progress</th>
                </tr>
            </thead>

            <tbody>
                ${goalRow('Market value', total, v('G3'), true)}
                ${goalRow('Gold', v('B7'), v('G7'))}
                ${goalRow('Silver', v('B6'), v('G5'))}
            </tbody>
        </table>
    `;

    // MONTHLY AND ANNUAL PURCHASE ACTIVITY

    const months = [
        'Jan', 'Feb', 'Mar', 'Apr',
        'May', 'Jun', 'Jul', 'Aug',
        'Sep', 'Oct', 'Nov', 'Dec'
    ];

    $('activity').innerHTML = `

        <h2>${esc(v('B32'))} Purchase Activity</h2>

        <div class="scroll">
            <table class="activity-table">

                <thead>
                    <tr>
                        <th rowspan="2">Month</th>
                        <th colspan="2" class="gold"
                            style="text-align:center">Gold</th>
                        <th colspan="2" class="silver"
                            style="text-align:center">Silver</th>
                        <th rowspan="2">Total $</th>
                    </tr>

                    <tr>
                        <th style="text-align:right">Oz</th>
                        <th>$</th>
                        <th>Oz</th>
                        <th>$</th>
                    </tr>
                </thead>

                <tbody>
                    ${months.map((month, i) => {

                        const r = i + 34;

                        return `
                            <tr>
                                <td>${month}</td>
                                <td>${fmt(v('B' + r), 4)}</td>
                                <td>${money(v('C' + r))}</td>
                                <td>${fmt(v('D' + r), 2)}</td>
                                <td>${money(v('E' + r))}</td>
                                <td>${money(v('F' + r))}</td>
                            </tr>`;
                    }).join('')}

                    <tr class="total-row">
                        <td>Total</td>
                        <td>${fmt(v('B46'), 4)}</td>
                        <td>${money(v('C46'))}</td>
                        <td>${fmt(v('D46'), 2)}</td>
                        <td>${money(v('E46'))}</td>
                        <td>${money(v('F46'))}</td>
                    </tr>
                </tbody>

            </table>
        </div>
    `;

    // READ GOLD AND SILVER GUIDE PRODUCTS

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

// COMPACT BUYING GUIDE

function renderGuide() {

    const q = $('search').value.toLowerCase().trim();

    const filtered = products.filter(r =>
        r.slice(0, 2).join(' ').toLowerCase().includes(q)
    );

    if (!filtered.length) {
        $('products').innerHTML =
            '<p>No matching products.</p>';
        return;
    }

    function metalRows(metal) {

        const rows = filtered.filter(r => r[0] === metal);

        if (!rows.length) return '';

        return `
            <tr class="group-row">
                <td colspan="6" class="${metal.toLowerCase()}">
                    ${metal}
                </td>
            </tr>

            ${rows.map(r => `
                <tr>
                    <td>${esc(r[1])}</td>
                    <td>${fmt(r[2], 4)}</td>
                    <td>${money(r[3])}</td>
                    <td class="good">${money(r[4])}</td>
                    <td>${money(r[5])}</td>
                    <td>
                        ${pct(r[6])} / ${pct(r[7])}
                    </td>
                </tr>
            `).join('')}
        `;
    }

    $('products').innerHTML = `

        <div class="scroll">
            <table class="guide-table">
                <thead>
                    <tr>
                        <th>Product</th>
                        <th>Oz</th>
                        <th>Melt</th>
                        <th>Strong</th>
                        <th>Accept.</th>
                        <th>Premiums</th>
                    </tr>
                </thead>

                <tbody>
                    ${metalRows('Gold')}
                    ${metalRows('Silver')}
                </tbody>
            </table>
        </div>

        <small>Premiums: Strong Buy / Acceptable Buy</small>
    `;
}

// WORKING ONEDRIVE CONNECTION

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

        const headers = {
            Authorization: 'Bearer ' + token.accessToken
        };

        const res = await fetch(path, {
            headers,
            cache: 'no-store'
        });

        if (!res.ok) {
            throw Error(
                'OneDrive lookup failed (HTTP ' +
                res.status +
                '). Check file path and permission.'
            );
        }

        let meta = await res.json();

        let url = meta['@microsoft.graph.downloadUrl'];

        if (!url && meta.id) {

            status('Workbook located. Retrieving download URL…');

            const itemPath =
                'https://graph.microsoft.com/v1.0/me/drive/items/' +
                encodeURIComponent(meta.id);

            const itemRes = await fetch(itemPath, {
                headers,
                cache: 'no-store'
            });

            if (!itemRes.ok) {
                throw Error(
                    'File ID lookup failed (HTTP ' +
                    itemRes.status + ').'
                );
            }

            const itemMeta = await itemRes.json();

            meta = {
                ...meta,
                ...itemMeta
            };

            url = meta['@microsoft.graph.downloadUrl'];
        }

        if (!url) {
            throw Error(
                'Workbook located, but OneDrive did not provide a download URL.'
            );
        }

        status('Downloading Excel workbook…');

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

        status('Reading Excel worksheets…');

        render(XLSX.read(await file.arrayBuffer(), {
            type: 'array'
        }));

        status(
            'Refreshed: ' + new Date().toLocaleString() +
            ' · Excel saved: ' +
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

// MICROSOFT AUTHENTICATION

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

// TWO-TAB NAVIGATION

document.querySelectorAll('[data-tab]').forEach(b => {

    b.onclick = () => {

        const dashboard = b.dataset.tab === 'holdings';

        document.querySelectorAll('[data-tab]').forEach(x => {
            x.classList.toggle('active', x === b);
        });

        $('holdings').hidden = !dashboard;
        $('activity').hidden = !dashboard;
        $('guide').hidden = dashboard;
    };
});

init();
