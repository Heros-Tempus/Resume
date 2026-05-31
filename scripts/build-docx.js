#!/usr/bin/env node
// scripts/build-docx.js
//
// Converts resume.json → Timothy_Miller_Resume.docx
//
// Usage (from repo root):
//   node scripts/build-docx.js
//
// Output: resume.docx  (repo root, next to index.html)
//
// Requirements:
//   npm install docx
//   (one-time; no package.json needed — installs locally in node_modules/)

'use strict';

const fs   = require('fs');
const path = require('path');

// Resolve paths relative to repo root (one level up from scripts/)
const REPO_ROOT  = path.resolve(__dirname, '..');
const JSON_PATH  = path.join(REPO_ROOT, 'resume.json');
const OUT_PATH   = path.join(REPO_ROOT, 'resume.docx');

const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, LevelFormat, ExternalHyperlink,
    BorderStyle, WidthType, ShadingType, VerticalAlign,
    TabStopType, UnderlineType,
} = require('docx');

// ─── Palette (mirrors style.css CSS variables) ────────────────────────────────
const ACCENT        = '1A3A5C';   // --accent
const ACCENT_LIGHT  = '2A5A8C';   // --accent-light
const TEXT_MUTED    = '555555';   // --text-muted
const BG_ALT        = 'E9EBEE';   // --bg-alt

// ─── Page geometry ────────────────────────────────────────────────────────────
// US Letter, 0.75" margins
const PAGE_W    = 12240;
const PAGE_H    = 15840;
const MARGIN    = 1080;                    // 0.75 inches in DXA
const CONTENT_W = PAGE_W - MARGIN * 2;    // 10080 DXA

// ─── Shared helpers ───────────────────────────────────────────────────────────

/** Thick accent-coloured horizontal rule (bottom border on an empty paragraph). */
function accentRule() {
    return new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 1 } },
        spacing: { before: 0, after: 80 },
        children: [],
    });
}

/** Section heading: uppercase, spaced letters, accent colour, underline rule. */
function sectionHeading(label) {
    return new Paragraph({
        spacing: { before: 280, after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: ACCENT, space: 4 } },
        children: [
            new TextRun({
                text: label.toUpperCase(),
                bold: true,
                size: 22,
                color: ACCENT,
                font: 'Arial',
                characterSpacing: 80,
            }),
        ],
    });
}

/** Dimmed uppercase sublabel (e.g. "ADDITIONAL PROJECTS"). */
function subLabel(text) {
    return new Paragraph({
        spacing: { before: 200, after: 80 },
        children: [
            new TextRun({
                text: text.toUpperCase(),
                bold: true,
                size: 18,
                color: TEXT_MUTED,
                font: 'Arial',
                characterSpacing: 80,
            }),
        ],
    });
}

/** Clickable hyperlink run. */
function hyperlink(text, url) {
    return new ExternalHyperlink({
        link: url,
        children: [
            new TextRun({
                text,
                color: ACCENT_LIGHT,
                size: 18,
                font: 'Arial',
                underline: { type: UnderlineType.SINGLE },
            }),
        ],
    });
}

/** Inline tag rendered as [Tag] in accent colour. */
function tagRun(text) {
    return new TextRun({
        text: `[${text}]`,
        color: ACCENT_LIGHT,
        size: 16,
        font: 'Arial',
        bold: true,
    });
}

/** Empty paragraph used purely for vertical spacing. */
function spacer(before = 0, after = 0) {
    return new Paragraph({ spacing: { before, after }, children: [] });
}

// ─── Section builders ─────────────────────────────────────────────────────────

function buildHeader(data) {
    const ghUrl = `https://github.com/${data.github}`;
    return [
        new Paragraph({
            spacing: { before: 0, after: 40 },
            children: [
                new TextRun({ text: data.name, bold: true, size: 52, font: 'Arial', color: ACCENT }),
            ],
        }),
        new Paragraph({
            spacing: { before: 0, after: 60 },
            children: [
                new TextRun({ text: data.tagline, size: 24, font: 'Arial', color: TEXT_MUTED }),
            ],
        }),
        new Paragraph({
            spacing: { before: 0, after: 200 },
            children: [hyperlink(`github.com/${data.github}`, ghUrl)],
        }),
        accentRule(),
    ];
}

function buildSummary(s) {
    return [
        sectionHeading(s.label),
        new Paragraph({
            spacing: { before: 80, after: 0 },
            children: [new TextRun({ text: s.text, size: 20, font: 'Arial' })],
        }),
    ];
}

function buildSkills(s) {
    // Two-column table: spaced-caps category name | dot-joined items
    // Odd rows get a light grey background (BG_ALT), even rows stay white — Option B.
    const noBorder  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
    const colL      = Math.round(CONTENT_W * 0.28);
    const colR      = CONTENT_W - colL;

    const rows = s.groups.map((group, idx) => {
        const fill    = idx % 2 === 0 ? BG_ALT : 'FFFFFF';
        const shading = { fill, type: ShadingType.CLEAR };

        return new TableRow({
            children: [
                new TableCell({
                    borders: noBorders,
                    shading,
                    width: { size: colL, type: WidthType.DXA },
                    margins: { top: 80, bottom: 80, left: 100, right: 120 },
                    verticalAlign: VerticalAlign.TOP,
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: group.category.toUpperCase(),
                                    bold: true,
                                    size: 16,
                                    font: 'Arial',
                                    color: TEXT_MUTED,
                                    characterSpacing: 60,
                                }),
                            ],
                        }),
                    ],
                }),
                new TableCell({
                    borders: noBorders,
                    shading,
                    width: { size: colR, type: WidthType.DXA },
                    margins: { top: 80, bottom: 80, left: 0, right: 100 },
                    verticalAlign: VerticalAlign.TOP,
                    children: [
                        new Paragraph({
                            children: [
                                new TextRun({
                                    text: group.items.join('  ·  '),
                                    size: 19,
                                    font: 'Arial',
                                }),
                            ],
                        }),
                    ],
                }),
            ],
        });
    });

    return [
        sectionHeading(s.label),
        spacer(80, 0),
        new Table({
            width: { size: CONTENT_W, type: WidthType.DXA },
            columnWidths: [colL, colR],
            rows,
        }),
    ];
}

function buildProjectCard(p, featured) {
    const paras = [];

    // Collect link objects
    const links = [];
    if (p.repo)  links.push({ label: 'GitHub', url: `https://github.com/Heros-Tempus/${p.repo}` });
    if (p.repos) p.repos.forEach(r => links.push({ label: r.label, url: `https://github.com/Heros-Tempus/${r.repo}` }));
    if (p.live)  links.push({ label: 'Live', url: p.live });

    // Name (left) + links (right) on one line via right-aligned tab stop
    const linkRuns = [];
    links.forEach((lk, i) => {
        if (i > 0) linkRuns.push(new TextRun({ text: '  ', size: 18, font: 'Arial' }));
        linkRuns.push(hyperlink(lk.label, lk.url));
    });

    paras.push(
        new Paragraph({
            spacing: { before: featured ? 160 : 80, after: 40 },
            tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
            children: [
                new TextRun({ text: p.name, bold: true, size: featured ? 22 : 20, font: 'Arial', color: ACCENT }),
                ...(linkRuns.length
                    ? [new TextRun({ text: '\t', size: 18, font: 'Arial' }), ...linkRuns]
                    : []),
            ],
        })
    );

    if (p.description) {
        paras.push(
            new Paragraph({
                spacing: { before: 0, after: 40 },
                children: [new TextRun({ text: p.description, size: 19, font: 'Arial' })],
            })
        );
    }

    if (p.tags && p.tags.length) {
        const tagRuns = [];
        p.tags.forEach((t, i) => {
            if (i > 0) tagRuns.push(new TextRun({ text: '  ', size: 16, font: 'Arial' }));
            tagRuns.push(tagRun(t));
        });
        paras.push(
            new Paragraph({
                spacing: { before: 0, after: featured ? 120 : 60 },
                children: tagRuns,
            })
        );
    }

    return paras;
}

function buildProjects(s) {
    const paras    = [sectionHeading(s.label)];
    const featured = s.items.filter(p => p.featured);
    const minor    = s.items.filter(p => !p.featured && !p.archive);

    featured.forEach(p => paras.push(...buildProjectCard(p, true)));

    if (minor.length) {
        paras.push(subLabel('Additional Projects'));
        minor.forEach(p => paras.push(...buildProjectCard(p, false)));
    }

    return paras;
}

function buildExperience(s) {
    const paras = [sectionHeading(s.label)];

    s.items.forEach((e, idx) => {
        paras.push(
            new Paragraph({
                spacing: { before: idx === 0 ? 100 : 200, after: 30 },
                tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
                children: [
                    new TextRun({ text: e.title, bold: true, size: 22, font: 'Arial' }),
                    new TextRun({ text: '\t', size: 20, font: 'Arial' }),
                    new TextRun({ text: e.date, size: 19, font: 'Arial', color: TEXT_MUTED }),
                ],
            })
        );
        paras.push(
            new Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    new TextRun({
                        text: e.location ? `${e.org} · ${e.location}` : e.org,
                        size: 19,
                        font: 'Arial',
                        color: ACCENT_LIGHT,
                        bold: true,
                    }),
                ],
            })
        );
        (e.bullets || []).forEach(b =>
            paras.push(
                new Paragraph({
                    numbering: { reference: 'bullets', level: 0 },
                    spacing: { before: 0, after: 20 },
                    children: [new TextRun({ text: b, size: 19, font: 'Arial' })],
                })
            )
        );
    });

    return paras;
}

function buildEducation(s) {
    const paras = [sectionHeading(s.label)];

    s.items.forEach((e, idx) => {
        paras.push(
            new Paragraph({
                spacing: { before: idx === 0 ? 100 : 200, after: 30 },
                tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
                children: [
                    new TextRun({ text: e.degree, bold: true, size: 22, font: 'Arial' }),
                    new TextRun({ text: '\t', size: 20, font: 'Arial' }),
                    new TextRun({ text: e.date, size: 19, font: 'Arial', color: TEXT_MUTED }),
                ],
            })
        );
        paras.push(
            new Paragraph({
                spacing: { before: 0, after: 40 },
                children: [
                    new TextRun({
                        text: e.institution,
                        size: 19,
                        font: 'Arial',
                        color: ACCENT_LIGHT,
                        bold: true,
                    }),
                ],
            })
        );
        (e.details || []).forEach(d =>
            paras.push(
                new Paragraph({
                    numbering: { reference: 'bullets', level: 0 },
                    spacing: { before: 0, after: 20 },
                    children: [new TextRun({ text: d, size: 19, font: 'Arial' })],
                })
            )
        );
    });

    return paras;
}

function buildCertifications(s) {
    const noBorder  = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
    const col       = Math.round(CONTENT_W / 2);

    const rows = [];
    for (let i = 0; i < s.items.length; i += 2) {
        const pair = [s.items[i], s.items[i + 1]].filter(Boolean);
        rows.push(
            new TableRow({
                children: pair.map(c =>
                    new TableCell({
                        borders: noBorders,
                        width: { size: col, type: WidthType.DXA },
                        shading: { fill: BG_ALT, type: ShadingType.CLEAR },
                        margins: { top: 80, bottom: 80, left: 100, right: 100 },
                        children: [
                            new Paragraph({
                                spacing: { before: 0, after: 20 },
                                children: [
                                    new TextRun({ text: c.name, bold: true, size: 19, font: 'Arial' }),
                                ],
                            }),
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: `${c.issuer} · ${c.date}`,
                                        size: 17,
                                        font: 'Arial',
                                        color: TEXT_MUTED,
                                    }),
                                ],
                            }),
                        ],
                    })
                ),
            })
        );
    }

    return [
        sectionHeading(s.label),
        spacer(80, 0),
        new Table({
            width: { size: CONTENT_W, type: WidthType.DXA },
            columnWidths: [col, col],
            rows,
        }),
    ];
}

// ─── Dispatch table ───────────────────────────────────────────────────────────

const RENDERERS = {
    summary:        buildSummary,
    skills:         buildSkills,
    projects:       buildProjects,
    experience:     buildExperience,
    education:      buildEducation,
    certifications: buildCertifications,
};

// ─── Main ─────────────────────────────────────────────────────────────────────

const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));

const children = [
    ...buildHeader(data),
    ...data.sections.flatMap(s => {
        const fn = RENDERERS[s.type];
        if (!fn) { console.warn(`Warning: no renderer for section type "${s.type}" — skipped.`); return []; }
        return fn(s);
    }),
];

const doc = new Document({
    numbering: {
        config: [{
            reference: 'bullets',
            levels: [{
                level: 0,
                format: LevelFormat.BULLET,
                text: '\u2022',
                alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 480, hanging: 240 } } },
            }],
        }],
    },
    styles: {
        default: { document: { run: { font: 'Arial', size: 20 } } },
    },
    sections: [{
        properties: {
            page: {
                size:   { width: PAGE_W, height: PAGE_H },
                margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
            },
        },
        children,
    }],
});

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(OUT_PATH, buf);
    console.log(`✓  Written: ${OUT_PATH}`);
}).catch(err => {
    console.error('Build failed:', err.message);
    process.exit(1);
});
