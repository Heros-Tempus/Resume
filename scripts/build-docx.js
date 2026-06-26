#!/usr/bin/env node
// scripts/build-docx.js
//
// Converts resume.json → resume.docx
//
// Usage (from repo root):
//   node scripts/build-docx.js
//
// Requirements:
//   npm install docx

'use strict';

const fs   = require('fs');
const path = require('path');

const REPO_ROOT = path.resolve(__dirname, '..');
const JSON_PATH = path.join(REPO_ROOT, 'resume.json');
const OUT_PATH  = path.join(REPO_ROOT, 'resume.docx');

const {
    Document, Packer, Paragraph, TextRun,
    AlignmentType, LevelFormat, ExternalHyperlink,
    BorderStyle, TabStopType,
} = require('docx');

// ─── Constants ────────────────────────────────────────────────────────────────
const FONT       = 'Garamond';
const TEXT_MUTED = '555555';   // dates, muted text
const RULE_COLOR = 'AAAAAA';   // thin horizontal rules

// US Letter, 0.75" margins
const PAGE_W    = 12240;
const PAGE_H    = 15840;
const MARGIN    = 1080;
const CONTENT_W = PAGE_W - MARGIN * 2;   // 10080 DXA

// ─── Helpers ──────────────────────────────────────────────────────────────────

// Thin grey rule drawn as a bottom-border on an empty paragraph
function rule() {
    return new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE_COLOR, space: 1 } },
        spacing: { before: 0, after: 80 },
        children: [],
    });
}

// Uppercase section heading with a thin rule below
function sectionHeading(label) {
    return new Paragraph({
        spacing: { before: 260, after: 60 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE_COLOR, space: 4 } },
        children: [
            new TextRun({
                text: label.toUpperCase(),
                bold: true,
                size: 20,
                font: FONT,
                characterSpacing: 60,
            }),
        ],
    });
}

// Clickable hyperlink rendered in black (no blue, matches PDF appearance)
function link(text, url) {
    return new ExternalHyperlink({
        link: url,
        children: [
            new TextRun({ text, size: 18, font: FONT, color: '000000' }),
        ],
    });
}

function spacer(before = 0, after = 0) {
    return new Paragraph({ spacing: { before, after }, children: [] });
}

// Selects the best URL to display beside a project name.
// Prefers the live URL; falls back to the GitHub repo URL.
function projectUrl(p) {
    if (p.live) return p.live;
    if (p.repo) return `https://github.com/Heros-Tempus/${p.repo}`;
    if (p.repos && p.repos.length) return `https://github.com/Heros-Tempus/${p.repos[0].repo}`;
    return null;
}

// ─── Section builders ─────────────────────────────────────────────────────────

// Centered name + GitHub URL, no tagline — matches PDF header layout
function buildHeader(data) {
    const ghUrl = `https://github.com/${data.github}`;
    return [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 40 },
            children: [
                new TextRun({ text: data.name, bold: true, size: 48, font: FONT }),
            ],
        }),
        new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 0, after: 140 },
            children: [link(`github.com/${data.github}`, ghUrl)],
        }),
        rule(),
    ];
}

function buildSummary(s) {
    return [
        sectionHeading(s.label),
        new Paragraph({
            spacing: { before: 80, after: 0 },
            children: [new TextRun({ text: s.text, size: 19, font: FONT })],
        }),
    ];
}

// Inline "Bold Category: items" — one paragraph per group, no table
function buildSkills(s) {
    const paras = [sectionHeading(s.label), spacer(60, 0)];
    s.groups.forEach(group => {
        paras.push(
            new Paragraph({
                spacing: { before: 0, after: 30 },
                children: [
                    new TextRun({ text: group.category + ': ', bold: true, size: 19, font: FONT }),
                    new TextRun({ text: group.items.join(', '), size: 19, font: FONT }),
                ],
            })
        );
    });
    return paras;
}

// Project name (bold, black) + full URL on the same line; bullets below; Stack line last
function buildProjectCard(p, featured) {
    const paras = [];
    const url   = projectUrl(p);

    paras.push(
        new Paragraph({
            spacing: { before: featured ? 140 : 80, after: 30 },
            children: [
                new TextRun({ text: p.name, bold: true, size: featured ? 22 : 20, font: FONT }),
                ...(url ? [
                    new TextRun({ text: ' ', size: 19, font: FONT }),
                    link(url, url),
                ] : []),
            ],
        })
    );

    (p.bullets || []).forEach(b =>
        paras.push(
            new Paragraph({
                numbering: { reference: 'bullets', level: 0 },
                spacing: { before: 0, after: 20 },
                children: [new TextRun({ text: b, size: 19, font: FONT })],
            })
        )
    );

    if (p.tags && p.tags.length) {
        const isBootDev = p.tags.includes('Boot.dev');
        const mainTags  = p.tags.filter(t => t !== 'Boot.dev');
        const stackText = 'Stack: ' + mainTags.join(', ') + (isBootDev ? ' (Boot.dev)' : '');
        paras.push(
            new Paragraph({
                numbering: { reference: 'bullets', level: 0 },
                spacing: { before: 0, after: featured ? 100 : 60 },
                children: [new TextRun({ text: stackText, size: 19, font: FONT })],
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
        paras.push(
            new Paragraph({
                spacing: { before: 180, after: 60 },
                children: [
                    new TextRun({
                        text: 'ADDITIONAL PROJECTS',
                        bold: true,
                        size: 18,
                        font: FONT,
                        color: TEXT_MUTED,
                        characterSpacing: 60,
                    }),
                ],
            })
        );
        minor.forEach(p => paras.push(...buildProjectCard(p, false)));
    }

    return paras;
}

// Title (bold) | Date, then Org (italic) | Location (italic) — two separate lines with tab stops
function buildExperience(s) {
    const paras = [sectionHeading(s.label)];

    s.items.forEach((e, idx) => {
        paras.push(
            new Paragraph({
                spacing: { before: idx === 0 ? 100 : 180, after: 20 },
                tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
                children: [
                    new TextRun({ text: e.title, bold: true, size: 22, font: FONT }),
                    new TextRun({ text: '\t', size: 19, font: FONT }),
                    new TextRun({ text: e.date, size: 19, font: FONT, color: TEXT_MUTED }),
                ],
            })
        );
        paras.push(
            new Paragraph({
                spacing: { before: 0, after: 40 },
                tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
                children: [
                    new TextRun({ text: e.org, size: 19, font: FONT, italics: true }),
                    ...(e.location ? [
                        new TextRun({ text: '\t', size: 19, font: FONT }),
                        new TextRun({ text: e.location, size: 19, font: FONT, italics: true }),
                    ] : []),
                ],
            })
        );
        (e.bullets || []).forEach(b =>
            paras.push(
                new Paragraph({
                    numbering: { reference: 'bullets', level: 0 },
                    spacing: { before: 0, after: 20 },
                    children: [new TextRun({ text: b, size: 19, font: FONT })],
                })
            )
        );
    });

    return paras;
}

// Institution (bold, alone) | degree+subtitle (italic) | date — then plain-text details.
// Boot.dev has no degree field: its first detail serves as the italic subtitle with the date.
function buildEducation(s) {
    const paras = [sectionHeading(s.label)];

    s.items.forEach((e, idx) => {
        // Institution heading — no date on this line
        paras.push(
            new Paragraph({
                spacing: { before: idx === 0 ? 100 : 200, after: 20 },
                children: [
                    new TextRun({ text: e.institution, bold: true, size: 22, font: FONT }),
                ],
            })
        );

        if (e.degree) {
            // Degree + optional honours subtitle on the next line; date right-aligned
            const degreeText = e.subtitle ? `${e.degree}. ${e.subtitle}` : e.degree;
            paras.push(
                new Paragraph({
                    spacing: { before: 0, after: 40 },
                    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
                    children: [
                        new TextRun({ text: degreeText, size: 19, font: FONT, italics: true }),
                        new TextRun({ text: '\t', size: 19, font: FONT }),
                        new TextRun({ text: e.date, size: 19, font: FONT, color: TEXT_MUTED }),
                    ],
                })
            );
            // Details as plain (non-italic) paragraphs
            (e.details || []).forEach(d =>
                paras.push(
                    new Paragraph({
                        spacing: { before: 0, after: 20 },
                        children: [new TextRun({ text: d, size: 19, font: FONT })],
                    })
                )
            );
        } else {
            // No degree (e.g. Boot.dev): first detail is italic with date; rest are plain
            (e.details || []).forEach((d, di) =>
                paras.push(
                    new Paragraph({
                        spacing: { before: 0, after: 20 },
                        ...(di === 0 ? { tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }] } : {}),
                        children: [
                            new TextRun({ text: d, size: 19, font: FONT, italics: true }),
                            ...(di === 0 ? [
                                new TextRun({ text: '\t', size: 19, font: FONT }),
                                new TextRun({ text: e.date, size: 19, font: FONT, color: TEXT_MUTED }),
                            ] : []),
                        ],
                    })
                )
            );
        }
    });

    return paras;
}

// Plain bullet list: "Name, Issuer, Date" — matches the PDF's simple cert list
function buildCertifications(s) {
    const paras = [sectionHeading(s.label), spacer(60, 0)];
    s.items.forEach(c =>
        paras.push(
            new Paragraph({
                numbering: { reference: 'bullets', level: 0 },
                spacing: { before: 0, after: 30 },
                children: [
                    new TextRun({ text: `${c.name}, ${c.issuer}, ${c.date}`, size: 19, font: FONT }),
                ],
            })
        )
    );
    return paras;
}

// ─── Dispatch ─────────────────────────────────────────────────────────────────

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
        if (!fn) { console.warn(`Warning: no renderer for type "${s.type}" — skipped.`); return []; }
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
                text: '•',
                alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 480, hanging: 240 } } },
            }],
        }],
    },
    styles: {
        default: { document: { run: { font: FONT, size: 20 } } },
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
