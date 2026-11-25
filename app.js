
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

const numeric = v =>
    typeof v === 'number' && Number.isFinite(v)
        ? v : 0;

let auth;
let account;
let products = [];
let purchases = [];
let purchaseYears = [];
let selectedYear = null;
let selectedMetal = 'All';
let workbookLoaded = false;

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

// READ EXCEL WORKBOOK

function render(wb) {

    const d = wb.Sheets['Enhanced Dashboard'];
    const g = wb.Sheets.Guide;
    const inv = wb.Sheets.Inventory_Master;

    if (!d || !g || !inv) {
        throw Error(
            'Expected Enhanced Dashboard, Guide and Inventory_Master worksheets.'
        );
    }

    const val = (s, c) =>
        s[c]?.t === 'e' ? null : s[c]?.v;

    const v = c => val(d, c);

    if (typeof v('C6') !== 'number' ||
        typeof v('C7') !== 'number') {

        throw Error(
            'Saved dashboard values missing. Recalculate and save in Excel.'
        );
    }

    const total = v('C6') + v('C7');

    // SPOT PRICES

    $('guideSpot').innerHTML = `
        <span class="gold">
            Gold: ${money(v('B2'))}
        </span>
        <span class="silver">
            Silver: ${money(v('B3'))}
        </span>
    `;

    // GOAL PROGRESS

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
            </tr>
        `;
    }

    // DASHBOARD

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

    // READ PURCHASE RECORDS

    purchases = [];

    const date1904 = !!wb.Workbook?.WBProps?.date1904;

    // Match the Excel SUMIFS range: rows 2 through 2000.

    const lastRow = Math.min(
        2000,
        XLSX.utils.decode_range(inv['!ref'] || 'A1').e.r + 1
    );

    for (let r = 2; r <= lastRow; r++) {

        const metal = String(
            val(inv, 'D' + r) ?? ''
        ).toLowerCase();

        if (metal !== 'gold' && metal !== 'silver') {
            continue;
        }

        const rawDate = val(inv, 'C' + r);

        if (typeof rawDate !== 'number' ||
            !Number.isFinite(rawDate)) {
            continue;
        }

        const date = XLSX.SSF.parse_date_code(
            rawDate,
            { date1904 }
        );

        if (!date || date.m < 1 || date.m > 12) {
            continue;
        }

        purchases.push({
            year: date.y,
            month: date.m,
            metal: metal,
            weight: numeric(val(inv, 'AH' + r)),
            cost: numeric(val(inv, 'Q' + r))
        });
    }

    // BUILD AVAILABLE YEARS

    const currentYear = new Date().getFullYear();
    const savedYear = Number(v('B32'));

    const years = new Set(
        purchases.map(p => p.year)
    );

    years.add(currentYear);

    if (Number.isInteger(savedYear) &&
        savedYear >= 1900 &&
        savedYear <= 9999) {
        years.add(savedYear);
    }

    purchaseYears = Array.from(years)
        .sort((a, b) => b - a);

    if (!purchaseYears.includes(selectedYear)) {

        selectedYear = purchaseYears.includes(savedYear)
            ? savedYear
            : currentYear;
    }

    renderActivity();

    // LOAD ALL GOLD AND SILVER GUIDE PRODUCTS

    products = XLSX.utils.sheet_to_json(g, {
        header: 1,
        range: 3,
        defval: null,
        raw: true
    })
    .map(r => {

        const metal = String(r[0] ?? '')
            .trim()
            .toLowerCase();

        return [
            metal === 'gold' ? 'Gold' :
            metal === 'silver' ? 'Silver' :
            metal,
            ...r.slice(1)
        ];
    })
    .filter(r =>
        ['Gold', 'Silver'].includes(r[0]) &&
        String(r[1] ?? '').trim() !== ''
    );

    workbookLoaded = true;

    renderGuide();
}

// MONTHLY PURCHASE ACTIVITY

function renderActivity() {

    const months = [
        'Jan', 'Feb', 'Mar', 'Apr',
        'May', 'Jun', 'Jul', 'Aug',
        'Sep', 'Oct', 'Nov', 'Dec'
    ];

    const monthly = months.map(() => ({
        goldOz: 0,
        goldCost: 0,
        silverOz: 0,
        silverCost: 0
    }));

    // Calculate selected year from inventory records.

    purchases.forEach(p => {

        if (p.year !== selectedYear) return;

        const m = monthly[p.month - 1];

        if (p.metal === 'gold') {

            m.goldOz += p.weight;
            m.goldCost += p.cost;

        } else if (p.metal === 'silver') {

            m.silverOz += p.weight;
            m.silverCost += p.cost;
        }
    });

    // Calculate annual totals.

    const totals = monthly.reduce((t, m) => {

        t.goldOz += m.goldOz;
        t.goldCost += m.goldCost;
        t.silverOz += m.silverOz;
        t.silverCost += m.silverCost;

        return t;

    }, {
        goldOz: 0,
        goldCost: 0,
        silverOz: 0,
        silverCost: 0
    });

    $('activity').innerHTML = `

        <div class="activity-heading">

            <h2>Purchase Activity</h2>

            <select id="purchaseYear"
                    aria-label="Purchase year">

                ${purchaseYears.map(year => `
                    <option value="${year}"
                        ${year === selectedYear
                            ? 'selected' : ''}>
                        ${year}
                    </option>
                `).join('')}

            </select>

        </div>

        <div class="scroll">

            <table class="activity-table">

                <thead>
                    <tr>
                        <th rowspan="2">Month</th>

                        <th colspan="2"
                            class="gold"
                            style="text-align:center">
                            Gold
                        </th>

                        <th colspan="2"
                            class="silver"
                            style="text-align:center">
                            Silver
                        </th>

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

                    ${monthly.map((m, i) => `

                        <tr>
                            <td>${months[i]}</td>

                            <td>
                                ${fmt(m.goldOz, 4)}
                            </td>

                            <td>
                                ${money(m.goldCost)}
                            </td>

                            <td>
                                ${fmt(m.silverOz, 2)}
                            </td>

                            <td>
                                ${money(m.silverCost)}
                            </td>

                            <td>
                                ${money(
                                    m.goldCost +
                                    m.silverCost
                                )}
                            </td>
                        </tr>

                    `).join('')}

                    <tr class="total-row">

                        <td>Total</td>

                        <td>
                            ${fmt(totals.goldOz, 4)}
                        </td>

                        <td>
                            ${money(totals.goldCost)}
                        </td>

                        <td>
                            ${fmt(totals.silverOz, 2)}
                        </td>

                        <td>
                            ${money(totals.silverCost)}
                        </td>

                        <td>
                            ${money(
                                totals.goldCost +
                                totals.silverCost
                            )}
                        </td>

                    </tr>

                </tbody>

            </table>

        </div>
    `;

    $('purchaseYear').onchange = e => {

        selectedYear = Number(e.target.value);

        renderActivity();
    };
}

// QUICK AUCTION SEARCH

function searchText(r) {

    const name = String(r[1] ?? '').toLowerCase();

    let aliases = '';

    if (/\bage\b/.test(name)) {
        aliases += ' american gold eagle';
    }

    if (/\base\b/.test(name)) {
        aliases += ' american silver eagle';
    }

    if (name.includes('peso')) {
        aliases += ' mexican mexico';
    }

    if (name.includes('pre 1933') ||
        name.includes('pre-1933')) {
        aliases += ' pre33 pre-33 pre 33 liberty indian';
    }

    if (name.includes('rooster')) {
        aliases += ' french france franc francs';
    }

    if (name.includes('gram')) {
        aliases += ' gold bar';
    }

    return `${r[0]} ${name} ${aliases}`.toLowerCase();
}

// COMPACT BUYING GUIDE

function renderGuide() {

    if (!workbookLoaded) return;

    const q = $('search').value
        .toLowerCase()
        .trim();

    const terms = q.split(/\s+/).filter(Boolean);

    const filtered = products.filter(r => {

        const metalMatch =
            selectedMetal === 'All' ||
            r[0] === selectedMetal;

        const searchable = searchText(r);

        const textMatch = terms.every(
            term => searchable.includes(term)
        );

        return metalMatch && textMatch;
    });

    $('productCount').textContent =
        filtered.length + ' of ' +
        products.length + ' products · USD';

    if (!filtered.length) {

        $('products').innerHTML =
            '<p>No matching products.</p>';

        return;
    }

    function price(v) {

        if (typeof v !== 'number' ||
            !Number.isFinite(v)) {
            return '—';
        }

        return fmt(v, 2);
    }

    // DISPLAY PRICE WITH PREMIUM BELOW

    function priceWithPremium(amount, premium) {

        const displayedPrice = price(amount);

        if (displayedPrice === '—') {
            return '—';
        }

        return `
            ${displayedPrice}
            <span class="premium">
                ${pct(premium)}
            </span>
        `;
    }

    function metalRows(metal) {

        const rows = filtered.filter(
            r => r[0] === metal
        );

        if (!rows.length) return '';

        return `

            <tr class="group-row">
                <td colspan="4"
                    class="${metal.toLowerCase()}">

                    ${metal.toUpperCase()}

                </td>
            </tr>

            ${rows.map(r => `

                <tr>

                    <td>
                        ${esc(r[1])}
                    </td>

                    <td>
                        ${price(r[3])}
                    </td>

                    <td class="good">
                        ${priceWithPremium(
                            r[4], r[6]
                        )}
                    </td>

                    <td class="acceptable">
                        ${priceWithPremium(
                            r[5], r[7]
                        )}
                    </td>

                </tr>

            `).join('')}
        `;
    }

    $('products').innerHTML = `

        <table class="guide-table">

            <thead>
                <tr>
                    <th>Product</th>
                    <th>Melt</th>
                    <th>Strong</th>
                    <th>Accept.</th>
                </tr>
            </thead>

            <tbody>
                ${metalRows('Gold')}
                ${metalRows('Silver')}
            </tbody>

        </table>
    `;
}

// ONEDRIVE CONNECTION

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

            status(
                'Workbook located. Retrieving download URL…'
            );

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
            'Refreshed: ' +
            new Date().toLocaleString() +
            ' · Excel saved: ' +
            (meta.lastModifiedDateTime
                ? new Date(
                    meta.lastModifiedDateTime
                  ).toLocaleString()
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

            throw Error(
                'Required libraries did not load.'
            );
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

            status(
                'Sign in to access your private workbook.'
            );
        }

    } catch (e) {

        status(
            'Initialization failed: ' + e.message
        );

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
            postLogoutRedirectUri:
                config.auth.redirectUri
        });

    } catch (e) {

        status(e.message);
    }
};

// INSTANT SEARCH

$('search').oninput = renderGuide;

// GOLD / SILVER FILTERS

document.querySelectorAll('[data-metal]').forEach(b => {

    b.onclick = () => {

        selectedMetal = b.dataset.metal;

        document.querySelectorAll('[data-metal]')
            .forEach(x => {

                x.classList.toggle('active', x === b);
            });

        renderGuide();
    };
});

// TWO-TAB NAVIGATION

document.querySelectorAll('[data-tab]').forEach(b => {

    b.onclick = () => {

        const dashboard =
            b.dataset.tab === 'holdings';

        document.querySelectorAll('[data-tab]')
            .forEach(x => {

                x.classList.toggle('active', x === b);
            });

        $('holdings').hidden = !dashboard;
        $('guide').hidden = dashboard;
    };
});

init();
