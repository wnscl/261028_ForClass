'use strict';

// HWPX 및 DOCX 문서는 ZIP + XML 구조
// JSZip으로 압축을 해제하고 DOMParser로 XML을 분석한다.

const $ = (id) => document.getElementById(id);

const PX = 1 / 75;

const MIME = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    svg: 'image/svg+xml'
};


// =====================================================
// XML 공통 함수
// =====================================================

const kids = (el, n) =>
    el ? Array.from(el.children).filter((c) => c.localName === n) : [];

const first = (el, n) =>
    kids(el, n)[0];

const all = (el, n) =>
    el ? Array.from(el.getElementsByTagNameNS('*', n)) : [];

const num = (v) =>
    Number(v) || 0;


function parseXml(text, label) {

    const doc = new DOMParser().parseFromString(
        text,
        'application/xml'
    );

    if (doc.getElementsByTagName('parsererror').length) {
        throw new Error(`${label} 파싱 실패`);
    }

    return doc;
}


const findEntry = (zip, path) => {

    const p = path.replace(/^\.?\//, '').toLowerCase();

    return Object.values(zip.files).find(
        (f) => !f.dir && f.name.toLowerCase() === p
    );
};


async function readXml(zip, path) {

    const f = findEntry(zip, path);

    if (!f) {
        throw new Error(`${path} 파일이 없습니다`);
    }

    return parseXml(await f.async('string'), path);
}


const baseName = (p) =>
    p.split('/').pop().replace(/\.[^.]+$/, '');


// =====================================================
// HWPX 1단계 : 스타일 분석
// =====================================================

function parseHeader(doc) {

    const charPr = new Map();
    const paraPr = new Map();

    for (const el of all(doc, 'charPr')) {

        const ul = first(el, 'underline');
        const st = first(el, 'strikeout');

        charPr.set(el.getAttribute('id'), {

            size: num(el.getAttribute('height')) / 100,

            color: el.getAttribute('textColor'),

            bold: !!first(el, 'bold'),

            italic: !!first(el, 'italic'),

            underline:
                !!ul && ul.getAttribute('type') !== 'NONE',

            strike:
                !!st && st.getAttribute('shape') !== 'NONE'
        });
    }

    for (const el of all(doc, 'paraPr')) {

        const align =
            all(el, 'align')[0]?.getAttribute('horizontal');

        const ls = all(el, 'lineSpacing')[0];

        paraPr.set(el.getAttribute('id'), {

            align: {
                CENTER: 'center',
                RIGHT: 'right',
                JUSTIFY: 'justify',
                LEFT: 'left'
            }[align],

            lineHeight:
                ls && ls.getAttribute('type') === 'PERCENT'
                    ? num(ls.getAttribute('value')) / 100
                    : 0
        });
    }

    return { charPr, paraPr };
}


// =====================================================
// HWPX 2단계 : 문서 구조 및 이미지 분석
// =====================================================

async function loadManifest(zip) {

    const f = findEntry(zip, 'Contents/content.hpf');

    const items = new Map();

    let spine = [];

    if (f) {

        const doc = parseXml(
            await f.async('string'),
            'content.hpf'
        );

        for (const it of all(doc, 'item')) {

            items.set(
                it.getAttribute('id'),
                it.getAttribute('href')
            );
        }

        spine = all(doc, 'itemref')
            .map((r) => items.get(r.getAttribute('idref')))
            .filter((h) => /section\d+\.xml$/i.test(h || ''));
    }

    if (!spine.length) {

        spine = Object.keys(zip.files)
            .filter((n) =>
                /^contents\/section\d+\.xml$/i.test(n)
            )
            .sort();
    }

    return { items, spine };
}


async function loadImages(zip, manifest) {

    const urls = new Map();

    for (const f of Object.values(zip.files)) {

        if (f.dir || !/^bindata\//i.test(f.name)) {
            continue;
        }

        const ext = f.name.split('.').pop().toLowerCase();

        const blob = new Blob(
            [await f.async('uint8array')],
            { type: MIME[ext] || '' }
        );

        urls.set(
            baseName(f.name),
            URL.createObjectURL(blob)
        );
    }

    for (const [id, href] of manifest.items) {

        const u = href && urls.get(baseName(href));

        if (u) {
            urls.set(id, u);
        }
    }

    return urls;
}


// =====================================================
// HWPX 3단계 : 텍스트 렌더링
// =====================================================

function renderText(t, span) {

    for (const n of t.childNodes) {

        if (n.nodeType === Node.TEXT_NODE) {

            span.append(n.data);

        }
        else if (n.localName === 'lineBreak') {

            span.append(document.createElement('br'));

        }
        else if (n.localName === 'tab') {

            span.append('\u2003');

        }
        else if (n.localName === 'nbSpace') {

            span.append('\u00A0');

        }
        else if (n.localName === 'fwSpace') {

            span.append('\u3000');
        }
    }
}


function applyChar(span, c) {

    if (!c) {
        return;
    }

    if (c.size) {
        span.style.fontSize = c.size + 'pt';
    }

    if (/^#[0-9a-f]{6}$/i.test(c.color || '')) {
        span.style.color = c.color;
    }

    if (c.bold) {
        span.style.fontWeight = 'bold';
    }

    if (c.italic) {
        span.style.fontStyle = 'italic';
    }

    const deco = [
        c.underline && 'underline',
        c.strike && 'line-through'
    ]
        .filter(Boolean)
        .join(' ');

    if (deco) {
        span.style.textDecoration = deco;
    }
}


function missing(msg) {

    const s = document.createElement('span');

    s.className = 'missing';

    s.textContent = msg;

    return s;
}


// =====================================================
// HWPX 4단계 : 이미지 렌더링
// =====================================================

function renderPic(pic, ctx) {

    const ref =
        all(pic, 'img')[0]?.getAttribute('binaryItemIDRef');

    const url = ctx.urls.get(ref);

    if (!url) {
        return missing(`[이미지 없음: ${ref}]`);
    }

    const w =
        num(first(pic, 'curSz')?.getAttribute('width')) ||
        num(first(pic, 'sz')?.getAttribute('width'));

    const img = new Image();

    img.src = url;

    img.alt = '문서 이미지';

    img.style.maxWidth = '100%';

    if (w) {
        img.style.width = w * PX + 'px';
    }

    img.onerror = () =>
        img.replaceWith(
            missing('[표시할 수 없는 이미지 형식]')
        );

    return img;
}


// =====================================================
// HWPX 5단계 : 표 렌더링
// =====================================================

function renderTable(tbl, ctx) {

    const t = document.createElement('table');

    t.className = 'hwpx-tbl';

    const sz = first(tbl, 'sz');

    if (sz) {
        t.style.width =
            num(sz.getAttribute('width')) * PX + 'px';
    }

    for (const tr of kids(tbl, 'tr')) {

        const row = t.insertRow();

        for (const tc of kids(tr, 'tc')) {

            const td = row.insertCell();

            const span = first(tc, 'cellSpan');

            if (span) {

                td.colSpan =
                    num(span.getAttribute('colSpan')) || 1;

                td.rowSpan =
                    num(span.getAttribute('rowSpan')) || 1;
            }

            const cw =
                num(first(tc, 'cellSz')?.getAttribute('width'));

            if (cw) {
                td.style.width = cw * PX + 'px';
            }

            for (const p of kids(first(tc, 'subList') || tc, 'p')) {

                td.append(renderParagraph(p, ctx));
            }
        }
    }

    return t;
}


// =====================================================
// HWPX 6단계 : 문단 렌더링
// =====================================================

function renderParagraph(p, ctx) {

    const box = document.createElement('div');

    box.className = 'hwpx-p';

    const pp =
        ctx.styles.paraPr.get(
            p.getAttribute('paraPrIDRef')
        );

    if (pp?.align) {
        box.style.textAlign = pp.align;
    }

    if (pp?.lineHeight) {
        box.style.lineHeight = pp.lineHeight;
    }

    for (const run of kids(p, 'run')) {

        const cp =
            ctx.styles.charPr.get(
                run.getAttribute('charPrIDRef')
            );

        for (const node of run.children) {

            if (node.localName === 't') {

                const span = document.createElement('span');

                applyChar(span, cp);

                renderText(node, span);

                box.append(span);

            }
            else if (node.localName === 'tbl') {

                box.append(renderTable(node, ctx));

            }
            else if (node.localName === 'pic') {

                box.append(renderPic(node, ctx));
            }
        }
    }

    if (!box.childNodes.length) {

        box.append(document.createElement('br'));
    }

    return box;
}


function renderSection(doc, ctx) {

    const sheet = document.createElement('section');

    sheet.className = 'sheet';

    const pg = all(doc, 'pagePr')[0];

    const mg = first(pg, 'margin');

    if (pg) {

        sheet.style.width =
            num(pg.getAttribute('width')) * PX + 'px';
    }

    if (mg) {

        sheet.style.padding =
            ['top', 'right', 'bottom', 'left']
                .map((k) =>
                    num(mg.getAttribute(k)) * PX + 'px'
                )
                .join(' ');
    }

    for (const p of kids(doc.documentElement, 'p')) {

        sheet.append(renderParagraph(p, ctx));
    }

    return sheet;
}


// =====================================================
// HWPX 7단계 : 파일 열기
// =====================================================

let currentUrls = [];

async function openHwpx(file) {

    const status = $('status');

    const viewer = $('viewer');

    status.className = '';

    status.textContent = `${file.name} 여는 중...`;

    currentUrls.forEach(URL.revokeObjectURL);

    currentUrls = [];

    viewer.replaceChildren();

    $('structure').hidden = true;

    try {

        if (/\.hwp$/i.test(file.name)) {

            throw new Error(
                '.hwp(바이너리) 형식은 지원하지 않습니다. .hwpx로 저장해 주세요.'
            );
        }

        const zip = await JSZip.loadAsync(file);

        const styles =
            parseHeader(
                await readXml(zip, 'Contents/header.xml')
            );

        const manifest = await loadManifest(zip);

        const urls = await loadImages(zip, manifest);

        currentUrls = [...new Set(urls.values())];

        const ctx = { styles, urls };

        for (const path of manifest.spine) {

            viewer.append(
                renderSection(
                    await readXml(zip, path),
                    ctx
                )
            );
        }

        $('tree').textContent =
            Object.values(zip.files)
                .map((f) =>
                    f.dir
                        ? f.name
                        : `${f.name}  (${f._data?.uncompressedSize ?? '?'} B)`
                )
                .join('\n');

        $('structure').hidden = false;

        status.textContent =
            `${file.name} - 섹션 ${manifest.spine.length}개, 이미지 ${urls.size}개`;

    }
    catch (e) {

        status.className = 'error';

        status.textContent =
            `열 수 없습니다: ${e.message}`;
    }
}


// =====================================================
// DOCX 1단계 : 텍스트 및 서식 렌더링
// =====================================================

function renderDocxRun(run) {

    const span = document.createElement('span');

    const rPr = first(run, 'rPr');

    if (rPr) {

        const bold = first(rPr, 'b');

        if (bold && bold.getAttribute('val') !== '0' &&
            bold.getAttribute('val') !== 'false') {

            span.style.fontWeight = 'bold';
        }

        const italic = first(rPr, 'i');

        if (italic && italic.getAttribute('val') !== '0' &&
            italic.getAttribute('val') !== 'false') {

            span.style.fontStyle = 'italic';
        }

        const underline = first(rPr, 'u');

        if (underline &&
            underline.getAttribute('val') !== 'none') {

            span.style.textDecoration = 'underline';
        }

        const size = first(rPr, 'sz');

        if (size) {

            const pt = num(size.getAttribute('val')) / 2;

            if (pt) {
                span.style.fontSize = pt + 'pt';
            }
        }

        const color = first(rPr, 'color');

        if (color) {

            const value = color.getAttribute('val');

            if (/^[0-9a-f]{6}$/i.test(value || '')) {

                span.style.color = '#' + value;
            }
        }
    }

    for (const node of run.children) {

        if (node.localName === 't') {

            span.append(document.createTextNode(node.textContent));

        }
        else if (node.localName === 'tab') {

            span.append('\u2003');

        }
        else if (node.localName === 'br' ||
                 node.localName === 'cr') {

            span.append(document.createElement('br'));
        }
    }

    return span;
}


// =====================================================
// DOCX 2단계 : 문단 렌더링
// =====================================================

function renderDocxParagraph(p) {

    const div = document.createElement('div');

    div.className = 'hwpx-p';

    const pPr = first(p, 'pPr');

    const jc = first(pPr, 'jc');

    if (jc) {

        const align = jc.getAttribute('val');

        const alignMap = {
            left: 'left',
            center: 'center',
            right: 'right',
            both: 'justify'
        };

        if (alignMap[align]) {

            div.style.textAlign = alignMap[align];
        }
    }

    for (const node of p.children) {

        if (node.localName === 'r') {

            div.append(renderDocxRun(node));

        }
        else if (node.localName === 'hyperlink') {

            for (const run of kids(node, 'r')) {

                div.append(renderDocxRun(run));
            }
        }
    }

    if (!div.childNodes.length) {

        div.append(document.createElement('br'));
    }

    return div;
}


// =====================================================
// DOCX 3단계 : 표 렌더링
// =====================================================

function renderDocxTable(tbl) {

    const table = document.createElement('table');

    table.className = 'hwpx-tbl';

    for (const tr of kids(tbl, 'tr')) {

        const row = table.insertRow();

        for (const tc of kids(tr, 'tc')) {

            const td = row.insertCell();

            const tcPr = first(tc, 'tcPr');

            const gridSpan = first(tcPr, 'gridSpan');

            if (gridSpan) {

                td.colSpan =
                    num(gridSpan.getAttribute('val')) || 1;
            }

            for (const node of tc.children) {

                if (node.localName === 'p') {

                    td.append(renderDocxParagraph(node));

                }
                else if (node.localName === 'tbl') {

                    td.append(renderDocxTable(node));
                }
            }
        }
    }

    return table;
}


// =====================================================
// DOCX 4단계 : 파일 열기
// =====================================================

async function openDocx(file) {

    const status = $('status');

    const viewer = $('viewer');

    status.className = '';

    status.textContent = `${file.name} 여는 중...`;

    currentUrls.forEach(URL.revokeObjectURL);

    currentUrls = [];

    viewer.replaceChildren();

    $('structure').hidden = true;

    try {

        // DOCX 압축 해제
        const zip = await JSZip.loadAsync(file);

        // DOCX 본문 XML 읽기
        const doc = await readXml(zip, 'word/document.xml');

        // XML 본문 찾기
        const body = first(doc.documentElement, 'body');

        if (!body) {

            throw new Error(
                'DOCX 본문을 찾을 수 없습니다.'
            );
        }

        // 문서를 출력할 영역 생성
        const sheet = document.createElement('section');

        sheet.className = 'sheet';

        sheet.style.padding = '40px';

        sheet.style.minHeight = '600px';

        let paragraphCount = 0;
        let tableCount = 0;

        // DOCX 내부 문단 및 표 렌더링
        for (const node of body.children) {

            if (node.localName === 'p') {

                sheet.append(
                    renderDocxParagraph(node)
                );

                paragraphCount++;

            }
            else if (node.localName === 'tbl') {

                sheet.append(
                    renderDocxTable(node)
                );

                tableCount++;
            }
        }

        // 브라우저에 출력
        viewer.append(sheet);

        // ZIP 내부 구조 출력
        $('tree').textContent =
            Object.values(zip.files)
                .map((f) =>
                    f.dir
                        ? f.name
                        : `${f.name}  (${f._data?.uncompressedSize ?? '?'} B)`
                )
                .join('\n');

        $('structure').hidden = false;

        status.textContent =
            `${file.name} - DOCX 문서 열기 완료 (문단 ${paragraphCount}개, 표 ${tableCount}개)`;

    }
    catch (e) {

        status.className = 'error';

        status.textContent =
            `열 수 없습니다: ${e.message}`;
    }
}


// =====================================================
// 파일 형식 자동 판별
// =====================================================

async function openDocument(file) {

    const ext = file.name
        .split('.')
        .pop()
        .toLowerCase();

    if (ext === 'hwpx') {

        await openHwpx(file);

    }
    else if (ext === 'docx') {

        await openDocx(file);

    }
    else {

        alert('지원하지 않는 파일 형식입니다.');
    }
}


// =====================================================
// 파일 선택 이벤트
// =====================================================

$('file').addEventListener('change', (e) => {

    const file = e.target.files[0];

    if (file) {

        openDocument(file);
    }

    // 동일한 파일을 다시 선택할 수 있도록 초기화
    e.target.value = '';
});


// =====================================================
// 드래그 앤 드롭 이벤트
// =====================================================

const drop = $('drop');

['dragenter', 'dragover'].forEach((ev) => {

    drop.addEventListener(ev, (e) => {

        e.preventDefault();

        drop.classList.add('over');
    });
});


['dragleave', 'drop'].forEach((ev) => {

    drop.addEventListener(ev, (e) => {

        e.preventDefault();

        drop.classList.remove('over');
    });
});


drop.addEventListener('drop', (e) => {

    const file = e.dataTransfer.files[0];

    if (file) {

        openDocument(file);
    }
});