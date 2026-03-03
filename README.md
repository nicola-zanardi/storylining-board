# Storylining & Ghost-Decking Board

A keyboard-first single-page app for structuring storylines into sections, slides, and bullets. Built for fast consulting-style ghost-deck workflows.

## Tech Stack

- React 19
- Vite 7
- Tailwind CSS 4
- @dnd-kit/core + @dnd-kit/sortable (drag and drop)
- lucide-react (icons)
- pptxgenjs (PowerPoint export)

## Prerequisites

- Node.js 20+ (recommended)
- npm 10+

## Install

```bash
npm install
```

## Compile and Run

### Development (hot reload)

```bash
npm run dev
```

Then open the local URL shown in terminal (usually `http://localhost:5173`).

### Production build (compile)

```bash
npm run build
```

Build output is generated in `dist/`.

### Run production build locally

```bash
npm run preview
```

### Lint

```bash
npm run lint
```

## How to Use

### 1. Structure the board

- The board is organized as **Section Rows**.
- Each section has a purple left anchor with the section title.
- Slides appear to the right in a wrapping card layout.

### 2. Edit content

- Storyline title at top is editable.
- Section titles, slide titles, and bullets are editable inline.
- Changes are saved automatically.

### 3. Keyboard-first flow

- `Enter` in a **slide title**: move focus to first bullet.
- `Enter` in a **bullet**: create a bullet below and focus it.
- `Backspace` on an **empty bullet**: delete bullet and move focus up.
- `Ctrl + Enter` (Windows/Linux) or `Cmd + Enter` (macOS): create a new slide below current.
- Empty bullet lines are cleaned up when focus leaves a slide card.

### 4. Drag and drop

- Drag **sections** to reorder the storyline.
- Drag **slides** within a section or across sections.

### 5. Manage content quickly

- Hover a section or slide to reveal Duplicate and Delete actions.

### 6. Project Manager

Use the top-right Project Manager to:

- Switch between saved projects
- Create a new project
- Rename current project

All data is persisted to browser `localStorage`.

### 7. Export to PowerPoint

- Click **Export PPTX** to export the full board into a **single 16:9 slide**.
- Use **Scale to Fit** toggle if content is dense.
- If the board is too large at native scale, a warning is shown.

## Visual Theme

Current palette:

- `strategy-blue`: `#d7e4ee`
- `deep-purple`: `#4d217a`

Configured in `tailwind.config.js` and used in both UI and PPTX export rendering.

## Troubleshooting

### Tailwind/PostCSS error about plugin

This project uses Tailwind v4 PostCSS adapter:

- `@tailwindcss/postcss`

If dependencies are missing:

```bash
npm install
```

## Project Scripts

- `npm run dev` - Start Vite dev server
- `npm run build` - Compile production build
- `npm run preview` - Serve compiled build
- `npm run lint` - Run ESLint


