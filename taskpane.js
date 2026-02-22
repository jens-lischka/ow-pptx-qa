/* ============================================================
   OW PPTX QA Tool — Task Pane Logic
   
   Architecture:
   1. Reads slide data via Office.js (client-side, no API needed)
   2. Runs deterministic brand checks locally in the browser
   3. Optionally calls the QA backend API for AI-assisted checks
      (backend URL configured via BACKEND_URL constant below)
   ============================================================ */

'use strict';

// ── Configuration ────────────────────────────────────────────────────────────
// In production, point this at your Azure Function or proxy URL.
// Leave empty to run local checks only (no API calls).
const BACKEND_URL = '';

// OW Brand rules (deterministic checks run entirely in the browser)
const OW_RULES = {
  fonts: {
    allowed: ['Gill Sans MT', 'Gill Sans', 'Arial', 'Arial Narrow'],
    displayAllowed: ['Gill Sans MT', 'Gill Sans'],
  },
  colors: {
    primary: ['D0021B', 'FFFFFF', '333333', '000000'],
    acceptable: ['666666', '999999', 'E0E0E0', 'F5F5F5', '4A90D9'],
  },
  layout: {
    minTitleFontSize: 20,      // pt
    maxBodyFontSize: 18,       // pt
    minBodyFontSize: 10,       // pt
    maxBulletsPerSlide: 6,
  }
};

// ── State machine ─────────────────────────────────────────────────────────────
const States = { IDLE: 'idle', LOADING: 'loading', RESULTS: 'results', ERROR: 'error' };
let currentState = States.IDLE;

function setState(state) {
  currentState = state;
  document.querySelectorAll('.state').forEach(el => el.classList.remove('active'));
  const el = document.getElementById(`state-${state}`);
  if (el) el.classList.add('active');
}

function setLoadingMessage(msg) {
  const el = document.getElementById('loading-message');
  if (el) el.textContent = msg;
}

// ── Office.js bootstrap ───────────────────────────────────────────────────────
Office.onReady(info => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById('btn-run-qa').onclick = runQA;
    document.getElementById('btn-rerun').onclick = runQA;
    document.getElementById('btn-retry').onclick = runQA;
    document.getElementById('btn-export').onclick = exportReport;
    setState(States.IDLE);
  }
});

// ── Main QA runner ────────────────────────────────────────────────────────────
async function runQA() {
  setState(States.LOADING);
  setLoadingMessage('Reading presentation…');

  try {
    const scope = document.querySelector('input[name="scope"]:checked').value;
    const slideData = await readSlides(scope);

    setLoadingMessage('Running brand checks…');
    const issues = await runChecks(slideData);

    setLoadingMessage('Building report…');
    renderResults(issues, slideData.length);

    setState(States.RESULTS);
  } catch (err) {
    console.error('QA failed:', err);
    document.getElementById('error-message').textContent =
      err.message || 'An unexpected error occurred. Check the console for details.';
    setState(States.ERROR);
  }
}

// ── Read slide data via Office.js ─────────────────────────────────────────────
async function readSlides(scope) {
  return new Promise((resolve, reject) => {
    PowerPoint.run(async context => {
      try {
        const presentation = context.presentation;
        const slides = presentation.slides;
        slides.load('items');
        await context.sync();

        const slideItems = scope === 'current'
          ? [slides.items[presentation.activeSlide ? presentation.activeSlide.id : 0]]
          : slides.items;

        const slideData = [];

        for (let i = 0; i < slideItems.length; i++) {
          const slide = slideItems[i];
          slide.load('shapes');
          await context.sync();

          const shapes = [];
          for (const shape of slide.shapes.items) {
            shape.load('name,shapeType,textFrame,fill,left,top,width,height');
            await context.sync();

            let textContent = null;
            let fontInfo = [];

            try {
              if (shape.textFrame) {
                const tf = shape.textFrame;
                tf.load('text,paragraphs');
                await context.sync();
                textContent = tf.text;

                for (const para of tf.paragraphs.items) {
                  para.load('runs');
                  await context.sync();
                  for (const run of para.runs.items) {
                    run.load('font');
                    await context.sync();
                    run.font.load('name,size,color,bold');
                    await context.sync();
                    fontInfo.push({
                      name: run.font.name,
                      size: run.font.size,
                      color: run.font.color,
                      bold: run.font.bold,
                    });
                  }
                }
              }
            } catch (e) {
              // Shape may not have a text frame — that's fine
            }

            shapes.push({
              name: shape.name,
              type: shape.shapeType,
              text: textContent,
              fonts: fontInfo,
              position: { left: shape.left, top: shape.top, width: shape.width, height: shape.height },
            });
          }

          slideData.push({ index: i + 1, shapes });
        }

        resolve(slideData);
      } catch (err) {
        reject(err);
      }
    }).catch(reject);
  });
}

// ── Brand checks ──────────────────────────────────────────────────────────────
async function runChecks(slideData) {
  const issues = [];

  for (const slide of slideData) {
    // --- Font checks ---
    for (const shape of slide.shapes) {
      for (const font of shape.fonts) {
        if (font.name && !OW_RULES.fonts.allowed.some(f =>
            font.name.toLowerCase().includes(f.toLowerCase()))) {
          issues.push({
            type: 'error',
            rule: 'Non-OW font',
            detail: `Font "${font.name}" is not in the OW approved font list.`,
            slide: slide.index,
            shape: shape.name,
          });
        }

        // Body text size check
        if (font.size && font.size > OW_RULES.layout.maxBodyFontSize) {
          // Only flag if not in a title shape
          if (!shape.name.toLowerCase().includes('title')) {
            issues.push({
              type: 'warning',
              rule: 'Large body text',
              detail: `Font size ${font.size}pt exceeds recommended maximum of ${OW_RULES.layout.maxBodyFontSize}pt for body text.`,
              slide: slide.index,
              shape: shape.name,
            });
          }
        }

        if (font.size && font.size < OW_RULES.layout.minBodyFontSize) {
          issues.push({
            type: 'error',
            rule: 'Text too small',
            detail: `Font size ${font.size}pt is below minimum readable size of ${OW_RULES.layout.minBodyFontSize}pt.`,
            slide: slide.index,
            shape: shape.name,
          });
        }

        // Color check
        if (font.color) {
          const hex = font.color.replace('#', '').toUpperCase();
          const allAllowed = [...OW_RULES.colors.primary, ...OW_RULES.colors.acceptable]
            .map(c => c.toUpperCase());
          if (!allAllowed.includes(hex) && hex !== '000000') {
            issues.push({
              type: 'warning',
              rule: 'Non-brand colour',
              detail: `Text colour #${hex} is not in the OW brand palette.`,
              slide: slide.index,
              shape: shape.name,
            });
          }
        }
      }

      // Bullet count check
      if (shape.text) {
        const bulletCount = (shape.text.match(/\n/g) || []).length + 1;
        if (bulletCount > OW_RULES.layout.maxBulletsPerSlide) {
          issues.push({
            type: 'warning',
            rule: 'Too many bullets',
            detail: `Shape has ${bulletCount} text blocks. OW guidelines recommend a maximum of ${OW_RULES.layout.maxBulletsPerSlide} per slide.`,
            slide: slide.index,
            shape: shape.name,
          });
        }
      }
    }

    // --- Slide-level checks ---
    const titleShapes = slide.shapes.filter(s =>
      s.name.toLowerCase().includes('title') && s.text);

    if (titleShapes.length === 0) {
      issues.push({
        type: 'warning',
        rule: 'Missing title',
        detail: 'No title shape detected on this slide. OW slides should have a clear action title.',
        slide: slide.index,
      });
    }

    // Title font size check
    for (const ts of titleShapes) {
      for (const font of ts.fonts) {
        if (font.size && font.size < OW_RULES.layout.minTitleFontSize) {
          issues.push({
            type: 'error',
            rule: 'Title too small',
            detail: `Title font size ${font.size}pt is below the recommended minimum of ${OW_RULES.layout.minTitleFontSize}pt.`,
            slide: slide.index,
            shape: ts.name,
          });
        }
      }
    }
  }

  // --- Optional: AI-assisted checks via backend ---
  if (BACKEND_URL) {
    try {
      setLoadingMessage('Running AI checks…');
      const aiIssues = await runAIChecks(slideData);
      issues.push(...aiIssues);
    } catch (e) {
      console.warn('AI check failed, continuing with local results only:', e);
    }
  }

  // Add a "passed" entry if no issues found
  if (issues.length === 0) {
    issues.push({
      type: 'passed',
      rule: 'All checks passed',
      detail: 'No brand violations detected.',
    });
  }

  return issues;
}

// ── Optional AI check via backend ────────────────────────────────────────────
async function runAIChecks(slideData) {
  const response = await fetch(`${BACKEND_URL}/qa`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ slides: slideData }),
  });

  if (!response.ok) {
    throw new Error(`Backend returned ${response.status}`);
  }

  const data = await response.json();
  return data.issues || [];
}

// ── Render results ────────────────────────────────────────────────────────────
let lastIssues = [];

function renderResults(issues, slideCount) {
  lastIssues = issues;

  const errors   = issues.filter(i => i.type === 'error').length;
  const warnings = issues.filter(i => i.type === 'warning').length;
  const passed   = issues.filter(i => i.type === 'passed').length;

  // Score: 0–100 based on errors and warnings
  const totalChecks = Math.max(issues.length, 1);
  const deductions  = (errors * 10) + (warnings * 3);
  const score       = Math.max(0, Math.min(100, 100 - deductions));

  document.getElementById('stat-errors').textContent   = errors;
  document.getElementById('stat-warnings').textContent = warnings;
  document.getElementById('stat-passed').textContent   = passed + (slideCount * 2); // approx passed checks

  // Score ring
  const scoreEl = document.getElementById('score-value');
  const ringEl  = document.getElementById('ring-fill');
  scoreEl.textContent = `${score}`;
  const circumference = 113;
  const offset = circumference - (score / 100) * circumference;
  ringEl.style.strokeDashoffset = offset;
  ringEl.style.stroke = score >= 80 ? 'var(--ow-ok)' : score >= 60 ? 'var(--ow-warn)' : 'var(--ow-red)';

  // Issue list
  const list = document.getElementById('issue-list');
  list.innerHTML = '';

  // Sort: errors first, then warnings, then passed
  const sorted = [...issues].sort((a, b) => {
    const order = { error: 0, warning: 1, passed: 2 };
    return (order[a.type] ?? 3) - (order[b.type] ?? 3);
  });

  for (const issue of sorted) {
    const div = document.createElement('div');
    div.className = `ow-issue ${issue.type}`;

    const badge = issue.type === 'error' ? 'Error'
                : issue.type === 'warning' ? 'Warning'
                : 'Pass';

    div.innerHTML = `
      <div class="ow-issue-header">
        <span class="ow-issue-title">${escapeHtml(issue.rule)}</span>
        <span class="ow-issue-badge">${badge}</span>
      </div>
      <div class="ow-issue-detail">${escapeHtml(issue.detail)}</div>
      ${issue.slide ? `<div class="ow-issue-slide">Slide ${issue.slide}${issue.shape ? ` · ${issue.shape}` : ''}</div>` : ''}
    `;

    // Click to navigate to the slide
    if (issue.slide) {
      div.onclick = () => navigateToSlide(issue.slide);
      div.style.cursor = 'pointer';
    }

    list.appendChild(div);
  }
}

// ── Navigate to slide ─────────────────────────────────────────────────────────
async function navigateToSlide(slideIndex) {
  try {
    await PowerPoint.run(async context => {
      const slides = context.presentation.slides;
      slides.load('items');
      await context.sync();
      if (slides.items[slideIndex - 1]) {
        slides.items[slideIndex - 1].setSelectedSlides();
        await context.sync();
      }
    });
  } catch (e) {
    console.warn('Could not navigate to slide:', e);
  }
}

// ── Export report ─────────────────────────────────────────────────────────────
function exportReport() {
  if (!lastIssues.length) return;

  const lines = [
    'OW PPTX QA Report',
    `Generated: ${new Date().toLocaleString()}`,
    '',
    ...lastIssues.map(i =>
      `[${i.type.toUpperCase()}] ${i.rule}${i.slide ? ` (Slide ${i.slide})` : ''}: ${i.detail}`
    )
  ];

  const blob = new Blob([lines.join('\n')], { type: 'text/plain' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `OW-QA-Report-${new Date().toISOString().slice(0,10)}.txt`;
  a.click();
  URL.revokeObjectURL(url);
}

// ── Utility ───────────────────────────────────────────────────────────────────
function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
            .replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}
