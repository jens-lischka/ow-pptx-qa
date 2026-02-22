# OW PPTX QA Tool

A PowerPoint task pane add-in that validates presentations against Oliver Wyman brand guidelines.

## Architecture

```
manifest.xml              ← Office Add-in manifest (sideloaded or deployed via M365 admin)
src/
  taskpane/
    taskpane.html         ← Task pane UI
    taskpane.css          ← OW brand styles
    taskpane.js           ← Office.js logic + local brand checks
  commands/
    commands.html         ← Ribbon command support
qa/
  main.py                 ← FastAPI backend (AI-assisted checks via Claude)
  requirements.txt
.github/
  workflows/
    ci-deploy.yml         ← CI tests + GitHub Pages deployment
```

## Quick start

### 1. Task pane (front end)
Served via GitHub Pages. See the step-by-step guide for setup.

### 2. Backend (optional, for AI checks)
```bash
cd qa
cp ../.env.example .env
# Edit .env and add your ANTHROPIC_API_KEY
pip install -r requirements.txt
uvicorn main:app --reload
```

### 3. Sideload the add-in
See the step-by-step guide: `OW-PPTX-QA-Setup-Guide.docx`

## GitHub Secrets required
| Secret | Description |
|--------|-------------|
| `ANTHROPIC_API_KEY` | Your Anthropic API key from console.anthropic.com |

## Brand rules checked
- **Fonts**: Gill Sans MT, Gill Sans, Arial, Arial Narrow only
- **Colours**: OW brand palette (#D0021B, #333333, #666666, etc.)
- **Text size**: Body 10–18pt, titles 20pt+
- **Bullet count**: Max 6 per slide
- **Action titles**: AI-assisted check for descriptive vs. label titles
- **Content quality**: AI-assisted tone and clarity review

## Deployment
The task pane is deployed to GitHub Pages automatically on every push to `main`.
The backend requires separate hosting (Azure App Service recommended for OW).
