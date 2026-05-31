fetch('resume.json')
    .then(r => r.json())
    .then(render)
    .catch(() => {
        document.querySelector('main').innerHTML =
            '<p style="padding:2rem;color:#c00;">Failed to load resume.json.</p>';
    });

function render(data) {
    renderHeader(data);
    renderNav(data.sections);
    renderSections(data.sections);
    renderFooter(data);
    document.title = data.name + ' — Resume';
}

// ─── Header ──────────────────────────────────────────────────────────────────

function renderHeader(data) {
    document.querySelector('h1').textContent = data.name;
    document.querySelector('.tagline').textContent = data.tagline;

    const link = document.querySelector('.github-link');
    link.href = 'https://github.com/' + data.github;
    link.textContent = 'github.com/' + data.github;
}

// ─── Nav ─────────────────────────────────────────────────────────────────────

function renderNav(sections) {
    const nav = document.querySelector('nav');
    sections.forEach(s => {
        const a = el('a');
        a.href = '#' + s.id;
        a.textContent = s.label;
        nav.appendChild(a);
    });

    const dl = el('a', 'nav-download');
    dl.href = 'resume.docx';
    dl.download = 'Timothy_Miller_Resume.docx';
    dl.textContent = '⬇ Download CV';
    nav.appendChild(dl);
}

// ─── Sections ────────────────────────────────────────────────────────────────

function renderSections(sections) {
    const main = document.querySelector('main');
    sections.forEach(s => main.appendChild(buildSection(s)));
}

function buildSection(s) {
    const section = el('section');
    section.id = s.id;

    const h2 = el('h2');
    h2.textContent = s.label;
    section.appendChild(h2);

    const renderers = {
        summary:        renderSummary,
        skills:         renderSkills,
        projects:       renderProjects,
        experience:     renderExperience,
        education:      renderEducation,
        certifications: renderCertifications,
    };

    const renderer = renderers[s.type];
    if (renderer) {
        const content = renderer(s);
        if (content) section.appendChild(content);
    }

    return section;
}

// ─── Summary ─────────────────────────────────────────────────────────────────

function renderSummary(s) {
    const p = el('p', 'summary-text');
    p.textContent = s.text;
    return p;
}

// ─── Skills ──────────────────────────────────────────────────────────────────

function renderSkills(s) {
    const grid = el('div', 'skills-grid');
    s.groups.forEach(group => {
        const div = el('div', 'skill-group');

        const h3 = el('h3');
        h3.textContent = group.category;
        div.appendChild(h3);

        const ul = el('ul');
        group.items.forEach(item => {
            const li = el('li');
            li.textContent = item;
            ul.appendChild(li);
        });
        div.appendChild(ul);

        grid.appendChild(div);
    });
    return grid;
}

// ─── Projects ────────────────────────────────────────────────────────────────

function renderProjects(s) {
    const frag = document.createDocumentFragment();

    const featured = s.items.filter(p => p.featured);
    const minor    = s.items.filter(p => !p.featured && !p.archive);
    const archive  = s.items.filter(p => p.archive);

    featured.forEach(p => frag.appendChild(buildProjectCard(p, 'featured')));

    if (minor.length) {
        frag.appendChild(subLabel('Additional Projects'));
        const grid = el('div', 'project-list');
        minor.forEach(p => grid.appendChild(buildProjectCard(p, 'minor')));
        frag.appendChild(grid);
    }

    if (archive.length) {
        frag.appendChild(subLabel('Course Archives'));
        const grid = el('div', 'project-list');
        archive.forEach(p => grid.appendChild(buildProjectCard(p, 'archive')));
        frag.appendChild(grid);
    }

    return frag;
}

function buildProjectCard(p, size) {
    const card = el('div', 'project-card ' + size);

    // ── Header row ──
    const header = el('div', 'project-header');

    const nameTag = size === 'featured' ? 'h3' : 'h4';
    const nameEl = el(nameTag);
    nameEl.textContent = p.name;

    if (p.archive) {
        const badge = el('span', 'archive-badge');
        badge.textContent = 'archive';
        nameEl.appendChild(badge);
    }
    header.appendChild(nameEl);

    const links = el('div', 'project-links');

    if (p.repo) {
        links.appendChild(repoLink(p.repo, 'GitHub'));
    }
    if (p.repos) {
        p.repos.forEach(r => links.appendChild(repoLink(r.repo, r.label)));
    }
    if (p.live) {
        const a = el('a');
        a.href = p.live;
        a.textContent = 'Live';
        a.target = '_blank';
        a.rel = 'noopener';
        links.appendChild(a);
    }

    if (links.children.length) header.appendChild(links);
    card.appendChild(header);

    // ── Description ──
    if (p.description) {
        const desc = el('p');
        desc.textContent = p.description;
        card.appendChild(desc);
    }

    // ── Tags ──
    if (p.tags && p.tags.length) {
        const tags = el('div', 'tags');
        p.tags.forEach(t => {
            const span = el('span');
            span.textContent = t;
            tags.appendChild(span);
        });
        card.appendChild(tags);
    }

    return card;
}

function repoLink(repo, label) {
    const a = el('a');
    a.href = 'https://github.com/Heros-Tempus/' + repo;
    a.textContent = label;
    a.target = '_blank';
    a.rel = 'noopener';
    return a;
}

// ─── Experience ──────────────────────────────────────────────────────────────

function renderExperience(s) {
    if (!s.items.length) return placeholder('Work experience coming soon.');
    const frag = document.createDocumentFragment();
    s.items.forEach(e => {
        const entry = el('div', 'experience-entry');

        const header = el('div', 'entry-header');
        const title = el('h3');
        title.textContent = e.title;
        const date = el('span', 'entry-date');
        date.textContent = e.date;
        header.appendChild(title);
        header.appendChild(date);
        entry.appendChild(header);

        const org = el('p', 'entry-org');
        org.textContent = e.location ? e.org + ' · ' + e.location : e.org;
        entry.appendChild(org);

        if (e.bullets && e.bullets.length) {
            const body = el('div', 'entry-body');
            const ul = el('ul');
            e.bullets.forEach(b => {
                const li = el('li');
                li.textContent = b;
                ul.appendChild(li);
            });
            body.appendChild(ul);
            entry.appendChild(body);
        }

        frag.appendChild(entry);
    });
    return frag;
}

// ─── Education ───────────────────────────────────────────────────────────────

function renderEducation(s) {
    if (!s.items.length) return placeholder('Education coming soon.');
    const frag = document.createDocumentFragment();
    s.items.forEach(e => {
        const entry = el('div', 'education-entry');

        const header = el('div', 'entry-header');
        const degree = el('h3');
        degree.textContent = e.degree;
        const date = el('span', 'entry-date');
        date.textContent = e.date;
        header.appendChild(degree);
        header.appendChild(date);
        entry.appendChild(header);

        const inst = el('p', 'entry-org');
        inst.textContent = e.institution;
        entry.appendChild(inst);

        if (e.details && e.details.length) {
            const body = el('div', 'entry-body');
            const ul = el('ul');
            e.details.forEach(d => {
                const li = el('li');
                li.textContent = d;
                ul.appendChild(li);
            });
            body.appendChild(ul);
            entry.appendChild(body);
        }

        frag.appendChild(entry);
    });
    return frag;
}

// ─── Certifications ──────────────────────────────────────────────────────────

function renderCertifications(s) {
    if (!s.items.length) return placeholder('Certifications coming soon.');
    const grid = el('div', 'cert-grid');
    s.items.forEach(c => {
        const card = el('div', 'cert-card');
        const name = el('p', 'cert-name');
        name.textContent = c.name;
        const meta = el('p', 'cert-meta');
        meta.textContent = c.issuer + ' · ' + c.date;
        card.appendChild(name);
        card.appendChild(meta);
        grid.appendChild(card);
    });
    return grid;
}

// ─── Footer ──────────────────────────────────────────────────────────────────

function renderFooter(data) {
    const a = el('a');
    a.href = 'https://github.com/' + data.github;
    a.textContent = 'github.com/' + data.github;
    a.target = '_blank';
    a.rel = 'noopener';
    document.querySelector('footer').appendChild(a);
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function el(tag, className) {
    const e = document.createElement(tag);
    if (className) e.className = className;
    return e;
}

function placeholder(msg) {
    const p = el('p', 'placeholder-notice');
    p.textContent = msg;
    return p;
}

function subLabel(text) {
    const p = el('p', 'subsection-label');
    p.textContent = text;
    return p;
}
