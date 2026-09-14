# Handoff: Sales Report Recolor — Team CVS & BHX

## Overview

Recolor / visual redesign of an existing internal sales-reporting web tool used by "Team Cam Giang". The tool ingests two Excel workbooks (BHX + WinMart hubs · CVS & NPP stores), auto-allocates targets across employees, computes % attainment per store and per rep, and lets the user snapshot the styled tables to image or export a formatted `.xlsx` matching `Team_CamGiang_Report.xlsx`.

**All parsing / allocation / export logic is unchanged.** The design work in this bundle is purely a **color-system swap** — moving the original soft-pink pastel palette to a professional, sales-energy palette across five variants so stakeholders can pick a direction.

The original file was **`original_index.html`** (soft-pink pastel, "Cute Chic" theme). Five redesign variants were produced.

## About the Design Files

The HTML files in this bundle are **design references**, not production code to ship as-is.

The original file is a single-page vanilla-HTML app that uses three CDN libraries at runtime — **SheetJS** (xlsx parsing), **html2canvas** (table → PNG snapshot), and **ExcelJS** (styled `.xlsx` export). If the target codebase already has a UI framework (React / Vue / SwiftUI / etc.), the developer should **rebuild the report screen inside that framework using its existing component library**, importing the design tokens from this handoff. If no framework exists yet, the single-file HTML can be adopted directly as the production build target — it is self-contained and works offline once the three CDN scripts are vendored.

The five HTML files each represent a full-app rebuild in a different palette. **Pick one** based on the "Recommendation" section below, then extract that variant's tokens for implementation.

## Fidelity

**High-fidelity (hifi).** Every hex value, spacing token, radius, shadow, font-weight, and interaction state is finalized in the HTML source. The developer should reproduce these values **pixel-exact**. All five variants share identical layout / DOM structure / component sizes — the only differences are colors, one accent gradient direction, and (for the two Command Center variants) a monospace font on numeric columns.

The one intentional lo-fi region is the `<h1>` brand text and the emoji icons (📊, ⚡) inline in the navbar — these are placeholders that a designer or brand team should replace with the real Team Cam Giang logotype and a proper glyph icon before production.

## Screens / Views

The app is a **single scrollable screen** composed of the following vertical sections. Every variant shares this layout.

### 1. Sticky Top Navbar (`.navbar`)
- **Height:** auto (padding `18px 28px`), sticky at `top: 0`, `z-index: 100`
- **Layout:** flex row, `justify-content: space-between`, wraps at narrow widths, gap `16px`
- **Left cluster (`.brand`):** flex row, gap `14px`
  - Square icon box `.brand-icon` `46×46px`, radius `12px`, contains an emoji
  - Title stack: `<h1>` 18px/800 tracking `-0.01em` + subtitle `<p>` 12.5px/500
- **Right cluster (`.nav-meta`):** flex row, wraps, gap `10px`
  - Pill badges (`.pill-badge`): white/dark surface, 1.5px border, radius `9999px`, 13px/600, padding `7px 15px`
  - Timer variant (`.pill-badge.timegone`): amber surface for "time gone" indicator
- **Border-bottom:** 1.5px solid var(--border-strong)
- **Shadow:** subtle drop, tinted with the primary color at low alpha

### 2. Target / KPI Banner (`.target-banner`)
Full-width strip below navbar summarizing daily/monthly target status. Flex row, wraps.
- Left: emoji `.target-icon` (28px) + heading `.target-summary-text h4` (15px/700) + subtitle `<p>` (13px/500)
- Right: pill badges or CTA button
- Padding `16px 20px`, radius `14px`, border `1.5px solid var(--border-strong)`

### 3. Unified Upload Dropzone (`.unified-dropzone`)
Large centered dashed-border card where the user drops both workbooks at once, OR uses the two split dropzones (`.upload-grid` → 2× `.dropzone`) for BHX file and CVS/NPP file separately.
- Radius `18px`, padding `30px 20px`
- Border `2.5px dashed` (primary color), with a subtle gradient fill
- Hover state: darker border, translate `-2px`, stronger shadow
- Loaded state (`.dropzone.loaded`, `.file-detected-card.loaded`): switches to a green success background + border regardless of variant

Below the dropzone is a **file-detection card row** (`.file-detected-card`) — two horizontal cards showing which files were parsed, with a 42×42 icon box on the left and a pending/success badge on the right.

### 4. KPI Summary Grid (`.kpi-grid`)
CSS grid, `repeat(auto-fit, minmax(210px, 1fr))`, gap `16px`. Contains 5 KPI cards:
- `.kpi-card.bhx` — BHX actual (accent color: emerald/lime depending on variant)
- `.kpi-card.cvs` — CVS actual (amber)
- `.kpi-card.total` — Total actual (primary accent)
- `.kpi-card.missing` — Missing stores (red/coral)
- `.kpi-card.pct` — Overall % attainment (variant-specific accent)

Each card:
- Padding `18px 20px`, radius `14px`
- 1.5px border, subtle drop shadow
- `::before` pseudo-element: `5px` wide vertical stripe on the left edge, color per-card (see list above)
- `.kpi-title` — 11.5px/700 uppercase, tracking `0.04em`
- `.kpi-value` — 23px/800, tabular numerals, tracking `-0.02em` (26px + JetBrains Mono in Command Center variants)
- `.kpi-sub` — 12px/500 muted

Hover: translate `-2px`, stronger shadow.

### 5. Status Banner (`.status-banner`)
Full-width strip signaling overall pass/warn/fail of the day. Three modifier classes:
- `.status-banner.pass` — green
- `.status-banner.warn` — amber
- `.status-banner.fail` — rose

Padding `16px 20px`, radius `14px`, 1.5px border, flex row wrapping.

### 6. Tab Bar (`.tabs-nav`) + Tab Panels
Horizontal scrollable tab strip. Bottom border 2px, gap `6px`. Each `.tab-btn`:
- Padding `11px 18px`, 13.5px/600
- Active state: darker text + tinted background + `::after` underline (3px solid primary, radius top-only)
- Command Center variants swap tab-btn to JetBrains Mono 12px uppercase and give the active underline a glow (`box-shadow: 0 0 12px currentColor`).

### 7. Toolbar + Data Tables (`.sheet-toolbar` + `.table-container` + `<table>`)
Each tab panel contains:
- Top toolbar (`.sheet-toolbar`): flex row with left cluster (filter label + select + count badge) and right cluster (search input + capture buttons)
- Data table wrapped in `.table-container` (radius `14px`, 1.5px border, `overflow-x: auto`)

Table styling:
- 13px body, 12px uppercase th (10.5–11px monospace in Command Center variants)
- `th` padding `12px 14px`, `td` padding `11px 14px`
- Zebra: `tbody tr:nth-child(even) td` gets tinted background
- Hover: row background changes to primary-tint
- `tfoot` is the "TOTAL" row — dark surface with white bold text and a colored top border (see per-variant footer colors)
- Chip badges inside cells (`.chip-pass` / `.chip-warn` / `.chip-fail`) — 11.5px/700, radius `6px`, padding `3px 9px`. Command Center variants add a glowing 6px dot indicator via `::before`.

### 8. Modals (`.modal-overlay` + `.modal-content`)
Fixed overlay, `rgba(...)` backdrop with `backdrop-filter: blur(4–8px)`. Modal body: max-width `650px`, radius `18px`, padding `26px`.

Spinner (`.spinner`): 40×40px, 3px border ring, `border-top-color` = primary, rotates 0.8s linear infinite.

## Interactions & Behavior

**All behavior is inherited from the original file and unchanged.** Documented here for the developer rebuilding in a framework:

- **Drag & drop** on `.dropzone` and `.unified-dropzone` — highlight with `.dragover` class, parse dropped file via SheetJS on drop
- **Click-to-upload** — same dropzones open a native file picker on click
- **File detection** — after parsing, `.dropzone.loaded` / `.file-detected-card.loaded` switches to green success state
- **Sheet filter selects** — dropdowns above each table filter rows (employee filter, store filter, etc.)
- **Search input** (`.search-input`) — live-filters table rows client-side
- **Capture buttons** (`.btn-capture` / `.btn-capture-outline`) — invoke `html2canvas` on the table clone at 100% width, download PNG or copy to clipboard
- **Excel export button** (`.btn-excel`) — builds a styled `.xlsx` via ExcelJS matching the `Team_CamGiang_Report.xlsx` template
- **Tab click** — swaps active tab and shows the corresponding panel (standard tabs pattern)
- **Modal open/close** — `.active` class toggles opacity + `pointer-events` + `transform: scale(0.97 → 1)`, `0.2s ease` transition
- **KPI hover** — `translateY(-2px)` + stronger shadow, `0.2s ease`
- **Button hover** — `translateY(-1px)` + brighter background, `0.15s ease`

### Command Center variants — extra behavior
- `.brand-icon` has a continuous `2.4s ease-in-out infinite` pulse animation (`@keyframes cc-pulse`) — box-shadow ring breathes from `0.06 → 0.18` alpha lime. **Skip this if reduced-motion is preferred by the user.**
- Active tab `::after` underline has a glow (`box-shadow: 0 0 12px currentColor, 0 0 4px currentColor`).
- Chip status dots have a `box-shadow: 0 0 8px <dot-color>` glow.
- Command Center Hybrid: row-hover on light tables shows a `3px inset lime shadow-left` on `td` — this is the "trace-the-row" affordance.

## State Management

Purely local UI state — no backend, no auth, no persistence beyond the current page session:
- Two file blobs (BHX file, CVS/NPP file) held in memory after parse
- Derived: array of employees, array of stores, per-employee allocation, per-store % attainment
- Filter selections (employee, store, tab) held in local component state
- Modal open/close flags
- Snapshot-in-progress flag (during html2canvas render)

If rebuilding in React: 5–7 `useState` hooks in a top-level page component are sufficient. No global store needed.

## Design Tokens

Five palettes. Every variant uses the SAME token names (`--primary`, `--surface`, `--text-main`, etc.) with different values, so a framework rebuild can define one set of CSS custom properties per variant and switch at the theme layer.

### Shared across all variants
```
--radius-sm: 8px
--radius:    14px
--radius-lg: 18px

Font family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif
             (weights 300–800)
Command Center variants also use: 'JetBrains Mono' (weights 400–700) for numeric columns, KPI values,
                                   chip labels, tab labels, and headers.

Standard easings:
  ease            → default hovers (0.15s)
  cubic-bezier(0.4, 0, 0.2, 1) → dropzone dragover (0.2s)
```

### Variant 1 — Modern Teal (Corporate Clean)
```
--primary:        #14b8a6      teal-500
--primary-dark:   #0f766e      teal-700
--primary-light:  #99f6e4      teal-200
--primary-soft:   #f0fdfa      teal-50
--accent:         #2dd4bf      teal-400
--success:        #059669      emerald-600
--warning:        #d97706      amber-600
--danger:         #e11d48      rose-600
--bg-main:        #f8fafc      slate-50 (page)
--surface:        #ffffff      pure white (cards)
--text-main:      #0f172a      slate-900
--text-muted:     #475569      slate-600
--text-subtle:    #0f766e      teal-700
--border:         #ccfbf1      teal-100
--border-strong:  #99f6e4      teal-200

Shadows (all shadows tinted with primary at low alpha):
--shadow-sm: 0 2px 6px  rgba(15, 23, 42, 0.04)
--shadow-md: 0 6px 18px rgba(15, 23, 42, 0.06)
--shadow-lg: 0 12px 28px rgba(20, 184, 166, 0.12)

Modal backdrop: rgba(15, 23, 42, 0.5)
KPI stripe (.pct::before): #14b8a6
Nav gradient: linear-gradient(135deg, #f1f5f9 0%, #ccfbf1 45%, #f0fdfa 100%)
Primary button: solid #14b8a6 → hover #0d9488 (flat, no gradient)
```

### Variant 2 — Teal + Coral Energy (Sales Team)
```
--primary:        #0d9488      teal-600
--primary-dark:   #e0533a      coral-deep
--primary-light:  #a7f3d0      emerald-200
--primary-soft:   #ecfdf5      emerald-50
--accent:         #ff8a65      soft coral
--success:        #059669
--warning:        #d97706
--danger:         #e11d48
--bg-main:        #f8fafc      slate-50
--surface:        #ffffff
--text-main:      #0f172a
--text-muted:     #475569
--text-subtle:    #0d9488
--border:         #d1fae5      emerald-100
--border-strong:  #a7f3d0      emerald-200

Primary CTA button: solid #ff6b47 (coral) → hover #e0533a — the "action" color
Neutral action button: solid #0d9488 (teal) → hover #0f766e — the "info" color
KPI stripe (.pct::before): #ff6b47
Modal backdrop: rgba(19, 78, 74, 0.55)
```

### Variant 3 — Deep Teal + Amber (Executive)
```
--primary:        #0f766e      teal-700
--primary-dark:   #134e4a      teal-900
--primary-light:  #5eead4      teal-300
--primary-soft:   #f0fdfa
--accent:         #f59e0b      amber-500
--success:        #059669
--warning:        #d97706
--danger:         #e11d48
--bg-main:        #fafaf9      stone-50 (warm neutral)
--surface:        #ffffff
--text-main:      #1c1917      stone-900
--text-muted:     #57534e      stone-600
--text-subtle:    #0f766e
--border:         #ccfbf1
--border-strong:  #5eead4

Primary button: solid #0f766e → hover #134e4a
Amber accent used for numeric highlights and KPI .pct stripe
KPI stripe (.pct::before): #f59e0b
Modal backdrop: rgba(28, 25, 23, 0.55)
```

### Variant 4 — Command Center Dark (Bloomberg × Linear × Vercel)
```
--primary:        #c6f24e      electric lime
--primary-dark:   #84cc16      lime-600
--primary-light:  #1e2938      slate-800 (used as prominent dark border)
--primary-soft:   #131a25      elevated surface
--accent:         #5eead4      iridescent teal
--success:        #34d399      emerald-400
--warning:        #fbbf24      amber-400
--danger:         #f87171      red-400
--bg-main:        #0a0e1a      midnight navy
--surface:        #101725      card surface
--text-main:      #e6edf3      off-white
--text-muted:     #8b98a9      slate-400
--text-subtle:    #5eead4
--border:         #171f2b      subtle dark border
--border-strong:  #1e2938      prominent dark border

Body background gradient (add radial atmospheric glows):
  radial-gradient(ellipse 1200px 800px at 20% -10%, rgba(94, 234, 212, 0.06), transparent 50%),
  radial-gradient(ellipse 1000px 700px at 80% 110%, rgba(198, 242, 78, 0.05), transparent 50%),
  #0a0e1a

Shadows (deep on dark, with inset hairline highlights):
  --shadow-sm: 0 1px 2px rgba(0,0,0,.6),  inset 0 0 0 1px rgba(255,255,255,.02)
  --shadow-md: 0 4px 12px rgba(0,0,0,.5), inset 0 0 0 1px rgba(255,255,255,.03)
  --shadow-lg: 0 12px 32px rgba(0,0,0,.6), inset 0 0 0 1px rgba(198,242,78,.06)

Primary button: solid #c6f24e text #0a0e1a
              hover #d4f56e + glow  box-shadow: 0 0 0 1px rgba(198,242,78,.7),
                                                0 6px 20px rgba(198,242,78,.4),
                                                0 0 30px rgba(198,242,78,.2)
Excel button: dark green fill #1a5f4a text #6ee7b7 border 1px #34d399
Outline button: transparent + 1px #2a3548 border, hover text #c6f24e border #c6f24e

Tab active underline: 2px #c6f24e + box-shadow: 0 0 12px #c6f24e, 0 0 4px #c6f24e
Modal backdrop: rgba(0, 0, 0, 0.75) + backdrop-filter: blur(8px)

Scrollbar (webkit):
  track  #0a0e1a
  thumb  #1e2938 (hover #2a3548)
  10px wide

Chip dots (::before, 6px circle):
  pass  #34d399  box-shadow: 0 0 8px #34d399
  warn  #fbbf24  box-shadow: 0 0 8px #fbbf24
  fail  #f87171  box-shadow: 0 0 8px #f87171

Brand-icon pulse:
  @keyframes cc-pulse {
    0%,100% { box-shadow: 0 0 0 3px rgba(198,242,78,0.06), inset 0 1px 0 rgba(255,255,255,0.04); }
    50%     { box-shadow: 0 0 0 3px rgba(198,242,78,0.18), inset 0 1px 0 rgba(255,255,255,0.04); }
  }
  animation: cc-pulse 2.4s ease-in-out infinite;
```

### Variant 5 — Command Center Hybrid (RECOMMENDED)
**Same tokens as Variant 4 for all chrome** (navbar, KPI cards, dropzones, tabs, buttons, sheet toolbar, modals, spinner, search input at page level, scrollbar). **Override for `.table-container` region only**:

```
.table-container:
  background:  #ffffff
  border:      1px solid #d8dee6
  box-shadow:  0 8px 24px rgba(0,0,0,.5), 0 0 0 1px rgba(198,242,78,.08)
  radius:      12px

th:
  background:  #f1f5f9
  color:       #334155
  border:      1px solid #d8dee6
  border-bottom: 2px solid #94a3b8
  font:        11px / 700 JetBrains Mono uppercase, tracking 0.06em
  padding:     12px 14px

th[highlight-blue] (was .th[style*="#2E75B6"] in source):
  bg #dbeafe · border #93c5fd · color #1e3a8a
th[highlight-red] (was .th[style*="#DD2323"]):
  bg #fee2e2 · border #fca5a5 · color #7f1d1d

td:
  background: transparent
  color:      #0f172a
  border:     1px solid #e2e8f0
  padding:    11px 14px

td.text-right:
  font: JetBrains Mono, tabular-nums, weight 500

Zebra:
  tr:nth-child(odd)  td → background #ffffff
  tr:nth-child(even) td → background #f8fafc

Row hover:
  background: #f0fdf4
  color:      #0f172a
  inset box-shadow: 3px 0 0 #84cc16   ← the signature lime trace-line

Footer (KEEP DARK for pop):
  tfoot { background: #0f172a }
  tfoot td { background: linear-gradient(180deg, #101828 0%, #0a0e1a 100%);
             color: #c6f24e;
             border: 1px solid #1e2938;
             border-top: 2px solid #c6f24e;
             font: JetBrains Mono 700; }

Chips on light table (override the dark chips):
  chip-pass  bg #ecfdf5  text #065f46  border #a7f3d0  dot #10b981 glow rgba(16,185,129,.6)
  chip-warn  bg #fffbeb  text #92400e  border #fde68a  dot #f59e0b glow rgba(245,158,11,.6)
  chip-fail  bg #fef2f2  text #991b1b  border #fecaca  dot #ef4444 glow rgba(239,68,68,.6)

Search input inside .card:
  bg #ffffff · text #0f172a · border 1px #d8dee6 · JetBrains Mono
  focus: border #84cc16 + box-shadow 0 0 0 3px rgba(132,204,22,.2)

Print media override — flatten to light for @media print:
  body { background: #ffffff }
  .navbar, .card, .kpi-card { background: #ffffff; color: #0f172a }
```

### Recommendation
**Ship Variant 5 (Command Center Hybrid).** It is the only variant tested against the "reads well when data actually loads" criterion — the user explicitly rejected pure dark mode after seeing hundreds of rows rendered. Variants 1–3 are safe corporate fallbacks if the audience prefers a lighter overall feel; Variant 4 is presentation-only.

### Spacing scale (used across all variants)
```
2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32
Card padding:       24px (16–20px on smaller UI)
Section vertical:   22px (margin-bottom between cards)
Grid gap:           16–22px
Table cell padding: 11–12px vertical, 14px horizontal
```

### Typography scale
```
Nav h1 title       18 / 800 / -0.01em tracking
Card title         16.5 / 700
Section header     15 / 700 / 800
Body               13 / 400–500
Small / caption    12 / 500
KPI value          23 / 800 / -0.02em tracking   (26 in Command Center)
KPI title (label)  11.5 / 700 uppercase / 0.04em tracking
th                 12 / 700 uppercase / 0.03em tracking
                   (11 mono / 700 / 0.06em in Command Center)
Chip               11.5 / 700
```

## Assets

- **Fonts (loaded from Google Fonts):**
  - `Plus Jakarta Sans` weights 300, 400, 500, 600, 700, 800
  - `JetBrains Mono` weights 400, 500, 600, 700 (Command Center variants only)
- **CDN JS libraries (vendor these for production):**
  - SheetJS `xlsx-0.20.1` — Excel parse
  - html2canvas `1.4.1` — table → PNG
  - ExcelJS `4.4.0` — styled xlsx export
- **Images:** none. All iconography is emoji-in-text (📊, ⚡, 🏢, 📦, 🏆, 👥). A brand team should replace `.brand-icon` and inline emoji with real icon glyphs (Lucide, Phosphor, or custom SVG) before production.
- **Report template reference:** `Team_CamGiang_Report.xlsx` (owned by the user's team — not included in this bundle, needed by the ExcelJS export step).

## Files

Located in this handoff folder:

- `original_index.html` — the pre-recolor source (soft-pink pastel, "Cute Chic"). Kept for diff reference so a developer can see exactly which properties changed.
- `Modern Teal.html` — Variant 1
- `Teal Coral Energy.html` — Variant 2
- `Deep Teal Amber.html` — Variant 3
- `Command Center Dark.html` — Variant 4 (presentation only)
- `Command Center Hybrid.html` — **Variant 5 (recommended)**
- `chooser.html` — landing page comparing all 5 variants side-by-side (was `index.html` in the design project)

Each variant is a complete, standalone HTML file that runs offline once the three CDN libraries are vendored. The developer can open any variant in a browser, drop the two Excel files, and see the full styled interactive prototype before starting the framework rebuild.
