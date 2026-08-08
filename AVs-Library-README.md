# AV's Library — Design & Build Guide

An interactive "galaxy map" of my to‑be‑read (TBR) list. Every book is a star; stars
cluster and glow by genre, forming a soft nebula you can pan, zoom, search, and read from.
Inspired by the [Open Syllabus Galaxy](https://galaxy.opensyllabus.org).

This document explains **what it is**, **how it works**, and **how to change it** — so the
look can be refined later without disturbing the stable, working app.

> [!IMPORTANT]
> **Note for AI assistants:** When implementing requested features or tweaks, always prioritize making **minimal changes** to the existing code structure. Do not rewrite large portions of the layout or rendering logic unless explicitly instructed. Always update this README and other documentation files with every change made.

---

## 1. The idea in one paragraph

I have ~1,435 books, comics, and manga on a TBR list in a spreadsheet. Instead of a boring
list, I want a **living star‑map**: books grouped into genre "galaxies" I can drift around,
hover to learn a region's name, search, pick a random book to "read today," mark books as
read, and **rate them 1–5 stars** (then filter by my rating). Each dot's **size reflects the
book's page length**. There are **two looks**: a **light** mode like a black‑and‑white
astronomical photo plate (dots in blues + black on white), and a **dark** mode like a color
galaxy photo (glowing colored stars on deep space). The chrome stays out of the way —
everything lives behind a ☰ menu.

**Design north star:** calm, spacious, a little cosmic. Light mode = observatory star‑plate;
dark mode = the universe.

---

## 2. What it does (features)

- **Galaxy map** — ~1,549 books as dots clustered into 24 genres (1,435 TBR + your read/rated
  Goodreads books). **Dot size = page length.** **Light mode — Brain** = blues + black on white
  with a **thin black brain outline** wrapping the cluster (the books read as neurons inside a
  brain); **Dark mode — universe** = glowing colored stars on deep space with a faint starfield.
- **📚 Goodreads‑synced read/unread** — your Goodreads *read* shelf and ratings are baked in, so
  read books show read (and rate‑able) out of the box. See §5b to refresh from a new export.
- **Light / dark toggle** — the ☀️/🌙 button in the top bar; the choice is remembered.
- **⭐ Ratings (Goodreads‑style)** — open a book and click its **1–5 stars**; rating a book also
  marks it read. Filter the map by **"Rating ≥ N"** in the menu. The top bar shows your average.
- **Minimal chrome** — a top bar with a ☰ hamburger (menu drawer), logo, read counter,
  theme toggle, and a 🔍 search shortcut. Nothing else covers the map.
- **Labels on demand** — genre names are hidden by default; they appear when you **hover** a
  region or **zoom in**.
- **Search** — title or author, with a results dropdown (shows rating); picking one flies to it.
- **✦ Read Today** — picks a random *unread* book (respecting your filters), flies to it, opens it.
- **Details panel** — click any star for genre, author, format, page count, your star rating,
  "more by this author," and "explore this genre."
- **Mark read / unread** — saved in the browser; read stars fade (marking unread clears its rating).
- **Filters** — All / Unread / Read · eBook / Bookshelf · Rating ≥ N · toggle any genre.
- **Connections** (off by default) — optional faint lines linking books by the same author.
- **Import read list** — paste titles you've read; add an optional rating per line
  (`Title | 4`, `Title ★★★★`, or `Title 4/5`) to bulk‑import ratings too.
- **Keyboard** — `/` search · `r` read‑today · `f` fit · `m` menu · `Esc` close.

State that matters — **which books are read, your ratings, and your theme** — is stored in the
browser's `localStorage`, so it persists between visits on the same browser. Everything else is
just the file — no server, no internet needed. Double‑click `AVs-Library.html` to open it.

---

## 3. How it works (the concept)

1. **Books → genres.** Each book is classified into one of 24 genres by a Python script
   (author lookups + title rules). See §5.
2. **Genres → galaxies.** Each genre gets a position on a big 2‑D plane (a spiral packing where
   clusters are allowed to overlap slightly, so they read as one connected landmass). Within a
   genre, books are scattered organically — same‑author books clump together, denser toward the
   center — using a gaussian spread. The whole field is stretched horizontally so it fills a
   wide screen.
3. **Dots → look.** Each dot's radius comes from the book's page count (`pagesToRad`). In
   **dark** mode an additive glow sprite is stamped under each star (the universe look) over a
   deep‑space gradient + static starfield; in **light** mode dots are flat blues/black with a
   soft grey "plate bloom" under the biggest ones (the astrophoto look).
4. **Interaction.** A single HTML `<canvas>` is redrawn on demand. A simple camera (pan/zoom
   with easing) maps world coordinates to the screen. Hovering finds the nearest star and the
   region under the cursor; clicking selects a star.
5. **Reading & rating.** Marking a book read adds its id to a list in `localStorage`; rating it
   1–5 stars stores `{id: stars}` and also marks it read. Read stars render faded.

Rendering order each frame (top of `draw()`): background (white, or space gradient + starfield)
→ glow/bloom pass → (optional connections) → **star dots** → **brain outline** (light mode only)
→ selection/hover ring → zoomed‑in book labels → **genre labels** (hover/zoom‑gated).

---

## 4. Project files

```
Downloads/Claude/
├── Books TBR.xlsx                 ← source data (sheet "ebooks": A=type, B=title, C=author, D=pages opt.)
├── goodreads_library_export.csv   ← your Goodreads export (read/to-read shelves, ratings, pages)
├── AVs-Library.html               ← THE APP. Self-contained. This is the "stable code."
├── AVs-Library-README.md          ← this document
└── pipeline/                      ← how the app is generated (only needed to change data/genres)
    ├── extract.py            ← Books TBR.xlsx           → books_raw.json
    ├── merge_goodreads.py    ← + goodreads CSV          → books_raw.json  (adds read/rated books, real pages)
    ├── classify.py           ← books_raw.json           → books_classified.json   (genres, page estimates)
    ├── build.py              ← books_classified.json    → ../AVs-Library.html      (the app + Goodreads seed)
    ├── books_raw.json        ← intermediate: cleaned titles + authors + type
    └── books_classified.json ← intermediate: adds a "genre" to every book
```

**Two mental models:**
- `AVs-Library.html` **is the product.** It embeds the book data and all the CSS/JS. For pure
  *look‑and‑feel* tweaks you can edit this file directly and just reopen it (§7, Path A).
- `pipeline/` **regenerates the product.** Use it when the *data or genres* change, or when you
  want a design change baked in permanently (§7, Path B).

---

## 5. The build pipeline (how to execute it)

One‑time dependency (for reading the spreadsheet):

```bash
pip3 install openpyxl
```

Run the steps from the `pipeline/` folder (merge is optional — skip it if you have no Goodreads file):

```bash
cd "/Users/abhinavverma/Downloads/Claude/pipeline"
python3 extract.py         # Books TBR.xlsx        -> books_raw.json
python3 merge_goodreads.py # + goodreads CSV       -> books_raw.json   (read/rated books + real pages)
python3 classify.py        # books_raw.json        -> books_classified.json  (prints genre counts)
python3 build.py           # books_classified.json -> ../AVs-Library.html
```

`build.py` overwrites `AVs-Library.html` in the parent folder. Reopen it in a browser to see the
result. (Your read/unread progress is stored in the browser, not the file, so rebuilding does
**not** erase it.)

### What each script does

- **extract.py** — reads the spreadsheet. If you rename the sheet or move columns, adjust the
  `ws = wb['ebooks']` line and the column indices at the top.
- **merge_goodreads.py** — folds `goodreads_library_export.csv` into `books_raw.json`: matches each
  Goodreads row to a TBR book (fuzzy title + author surname), attaches its **rating**, **read shelf**,
  and **real page count**; and **adds** any *read/rated* Goodreads book the TBR didn't already have.
  Idempotent (re‑running rebuilds the additions). Prints matched / added counts.
- **classify.py** — the "brain." It (a) cleans messy filenames (e.g. the long Anna's‑Archive /
  z‑library names), extracts authors, and (b) assigns a genre using, in priority order:
  manga list → comics keywords → exact title map → author→genre map → title keywords. It prints
  a genre distribution and any `UNCLASSIFIED` books at the end. Every book starts **unread**.
- **build.py** — holds the entire HTML/CSS/JS as a template with a `__BOOKS_JSON__` placeholder,
  injects the book data, and writes the final self‑contained app.

> **Important:** the CSS/JS constants described in §7 exist **twice** — once in the live
> `AVs-Library.html` and once inside the template string in `pipeline/build.py`. Editing the
> HTML gives an instant preview; to make a change permanent (survive a rebuild), make the *same*
> edit in `build.py`.

### 5b. Syncing your Goodreads export

Export your library from Goodreads (**My Books → Import/Export → Export Library**) and save the
`goodreads_library_export.csv` next to `Books TBR.xlsx`. Re‑running the pipeline (with the
`merge_goodreads.py` step) will:

- **mark as read** every TBR book on your Goodreads *read* shelf (or with a rating), and apply the
  rating and real page count;
- **add** your read/rated Goodreads books that weren't in the TBR (so you see read *and* unread);
- **bake a one‑time seed** into the app (`SEED_READ` / `SEED_RATINGS`) that populates your read +
  ratings in the browser on first open.

**To refresh after a newer export:** replace the CSV, re‑run the pipeline, and bump
`SEED_VERSION` (e.g. `"goodreads-1"` → `"goodreads-2"`) in `build.py` so the app re‑applies the
new baseline (existing manual marks are kept — the seed only adds). Goodreads titles are matched
by normalized title + author surname; a stubborn mismatch can be fixed with a `TITLE_MAP` entry
in `classify.py`. Current merge: **156 TBR books matched, 114 read books added → 1,549 total,
170 read/rated.**

---

## 6. Adding or changing books

1. Edit **`Books TBR.xlsx`** (add rows: column A = `ebook` or `in my bookshelf`, B = title,
   C = author optional, **D = pages optional** — a real page count here overrides the estimate
   and sets that dot's size).
2. Re‑run the pipeline (§5).
3. If a new book lands in the wrong genre, teach `classify.py` (see §8, "Reclassify a book").

To mark / rate books you've already read: open the app → ☰ → **Import read list** → paste one
per line, with an optional rating (`Title | 4`, `Title ★★★★`, or `Title 4/5`). Or rate any book
directly by clicking its stars in the detail panel.

---

## 7. Refining the design — the knobs

You can change almost all of the look by editing a handful of constants. Search for the quoted
strings inside **`AVs-Library.html`** (Path A: instant, no tools) and/or **`pipeline/build.py`**
(Path B: permanent). Values below are the current defaults.

### 7.1 Colours & theme
| What | Find | Now | Notes |
|---|---|---|---|
| Dark / legend palette | `var GENRE_COLORS={` | 24 hexes | Vivid star colour per genre (dark mode + legend swatches + tooltip dot). |
| Light‑mode dot tones | `var LIGHT_TONES=[...]` | 2 blues + black | The 3 tones in light mode, picked by dot size (small→large). |
| UI colours (both themes) | `:root{` and `body.theme-dark{` (CSS) | — | Every panel colour is a CSS var; the `body.theme-dark{…}` block holds all dark overrides. |
| Default theme | `localStorage.getItem(THEME_KEY)||"light"` | light | Change `"light"` → `"dark"` to open dark by default. |
| Dark space background | the `bg.addColorStop(...)` gradient in `draw()` | deep blue‑black | The gradient behind dark mode. |
| Background starfield | `BGSTARS` / `drawBgStars()` | 220 faint stars | Ambient stars in dark mode only. |
| Top bar height | `--topbar-h:56px;` | 56px | Canvas offsets itself below this automatically. |

### 7.2 Galaxy layout (shape of the map)
| What | Find | Now | Effect |
|---|---|---|---|
| Horizontal stretch | `var STRETCHX=1.32;` | 1.32 | >1 = wider field (fills landscape). 1.0 = circular. |
| Cluster spacing | `SP=18` | 18 | Base spiral spacing between genre centers (lower = tighter). |
| Cluster overlap | `OVER=0.70` | 0.70 | <1 lets galaxies overlap into one mass. Lower = tighter/more blended. |
| Cluster size | `r=31*Math.sqrt(n)+10` | — | Bigger number = larger, sparser genres. |
| Clump tightness | `var spread=Math.min(R*0.17,` … | 0.17 | How tightly same‑author books clump. |
| Center density | `Math.pow(u,0.55)` | 0.55 | Lower = more books pulled toward each galaxy's core. |

### 7.3 Glow & bloom (the luminous pass)
| What | Find | Now | Effect |
|---|---|---|---|
| Dark star‑glow strength | `ctx.globalAlpha=isRead(...)?0.10:0.42;` | 0.42 | Brightness of each star's halo in dark mode. |
| Dark glow size | `var gz=Math.max(7, rr*5);` | 5 | Halo size relative to dot size. |
| Glow falloff | `grd.addColorStop(...)` in `glowSprite()` | — | Tighter / feathier star glow. |
| Light plate‑bloom | `ctx.globalAlpha=0.5;` + the `rr<3.2` cutoff (else branch) | 0.5 / 3.2 | Soft grey halo under big light‑mode dots; raise the cutoff to bloom fewer. |

### 7.3b Brain outline (light mode only — "Light mode — Brain")

A thin black brain is stroked around the whole cluster in light mode so the books read as neurons
inside a brain. It's drawn in world space (pans/zooms with the cluster) and skipped entirely in
dark mode (`if(!dark) drawBrain();`).

| What | Find | Now | Effect |
|---|---|---|---|
| Shape | `var BRAIN=[ … ]` (anchor points) | — | The outline is anchor points smoothed by `strokeSmooth` (Catmull‑Rom). Sub‑paths: cerebrum, Sylvian fissure, cerebellum + folia, brainstem, gyri. Edit/add points to reshape. |
| How tightly it wraps | `var targW=…*1.24, targH=…*1.42;` in `drawBrain()` | 1.24 / 1.42 | Bigger = the brain sits further outside the cluster. |
| Line colour / weight | `ctx.strokeStyle="rgba(10,12,18,0.82)"` · `ctx.lineWidth=…/S` | black, ~1.1px | Thin black lines. Lower the `1.1` for thinner. |
| Toggle name | `toast(...:"Light mode — Brain")` | — | The label shown when switching to light mode. |

### 7.4 Genre labels (hover / zoom reveal)
| What | Find | Now | Effect |
|---|---|---|---|
| Zoom fade‑in range | `var zoomA = cam.s<=0.42?0 : cam.s>=0.78?1 :` | 0.42→0.78 | Labels fade in between these zoom levels. Lower both to show them sooner. |
| Label size | `Math.min(27,m.r*cam.s*0.085)` | 27 / 0.085 | Cap and scale of label font. |
| Light‑label darkness | `var lf=(g==="Manga"\|\|g==="Comics & Graphic Novels")?0.26:0.42;` | 0.42 (warm 0.26) | Light‑mode labels are a **darker shade** of the genre colour so they read on white. Lower = darker. The warm‑coloured Manga/Comics labels use an even darker `0.26`; add other genres to that check if any stay too pale. |
| Label opacity | `(isHover?0.98:0.9)` | — | Opacity of hovered vs. zoom‑revealed labels. |
| Collision padding | the box math around `drawnBoxes` | — | Governs how aggressively overlapping labels are dropped. |

### 7.5 Stars (the dots) — size = page length
| What | Find | Now | Effect |
|---|---|---|---|
| Dot size vs pages | `function pagesToRad(pg){` → `0.55+0.135*Math.sqrt(pg)` | clamp 1.1–6.6 | Maps page count to radius. Bigger coefficient = larger spread. |
| Light tone thresholds | `r>=4.0?LIGHT_TONES[2]:(r>=2.5?...)` in `dotColor()` | 4.0 / 2.5 | Size cutoffs picking black vs dark‑ vs light‑blue in light mode. |
| Unread dot | `hexA(col,.9)` (light) / `col` (dark) | — | Dark unread dots also get a white‑hot centre when `rad>2.1`. |
| Read dot opacity | `hexA(col,dark?.36:.28)` | — | How faded a read star looks. |
| Filtered‑out dots | `"rgba(190,194,203,.4)"` / `"rgba(120,132,164,.26)"` | grey | Books hidden by a filter (light / dark). |
| Selection ring | strokeStyle in `drawAccent()` | theme‑aware | Selected vs. hover ring colour. |

### 7.5b Ratings (stars)
| What | Find | Now | Effect |
|---|---|---|---|
| Star colour (filled) | `.stars .star.on{color:#f5b301}` (CSS) | gold | The filled‑star colour everywhere. |
| Rating filter | `#ratefilterStars` / `state.minRating` | — | The "Rating ≥ N" control in the drawer; filter logic is in `passFilter`. |
| Rating input | `#dstars` / `renderDetailStars()` / `setRating()` | — | The clickable stars in the detail panel. |
| Import rating parse | `parseImportLine()` | `\| N` · `★` · `N/5` | Formats accepted when importing ratings. |

### 7.6 Camera & motion
| What | Find | Now | Effect |
|---|---|---|---|
| Fit padding | `var pad=60,` | 60 | Margin around the galaxy when framed. |
| Zoom limits | `Math.max(0.03,Math.min(16,` | 0.03–16 | Min/max zoom. |
| Intro zoom‑in | `cam.s=ts*0.72; flyTo(tx,ty,ts,1200);` | 0.72 / 1200ms | Opening animation start scale and duration. |
| Connections default | `showConn:false` | off | Set `true` to show author links by default. |

### 7.7 Data / persistence (localStorage keys)
| What | Find | Now | Effect |
|---|---|---|---|
| Read‑state key | `var LS_KEY="avs-library-read-v1";` | array of ids | Which books are read. Bump the version to reset. |
| Ratings key | `var RATE_KEY="avs-library-ratings-v1";` | `{id: 1‑5}` | Your star ratings. Rating a book also marks it read. |
| Theme key | `var THEME_KEY="avs-library-theme-v1";` | `"light"`/`"dark"` | Remembered light/dark choice. |
| Goodreads seed | `var SEED_KEY` + `SEED_VERSION="goodreads-1";` | applied once | On first load, `SEED_READ`/`SEED_RATINGS` (baked from your Goodreads export at build time) are merged into read + ratings. Bump `SEED_VERSION` after a new export to re‑apply. |
| Embedded book data | `const BOOKS =` | — | Generated by the pipeline — don't hand‑edit; change the spreadsheet / Goodreads CSV + rebuild. |

---

## 8. Common recipes

**Recolor a genre** — in `GENRE_COLORS`, change that genre's hex (do it in both
`AVs-Library.html` and `build.py`). Reopen; done.

**Make dark stars glow more / less** — nudge the `0.42` glow alpha in §7.3.

**Change how strongly page count affects dot size** — edit `pagesToRad()` (§7.5): raise the
`0.135` coefficient for a bigger spread between short and long books.

**Open in dark mode by default** — change `||"light"` to `||"dark"` on the `theme` line.

**Show genre labels sooner** — lower the two numbers in `var zoomA = cam.s<=0.42?0 : cam.s>=0.78?1`
(e.g. `0.25` and `0.5`).

**Separate the galaxies more (less blended)** — raise `OVER` toward `0.95`, or raise `SP`.

**Bulk‑import ratings** — ☰ → Import read list → paste `Title | 4` (or `Title ★★★★`, `Title 4/5`)
lines. Great for pasting a Goodreads‑style export (title + your rating).

**Rename a genre** — in `pipeline/classify.py` the genre strings appear in the mapping tables;
rename consistently, then also add a matching entry in `GENRE_COLORS` (build.py) for the new
name. Rebuild.

**Reclassify a book** — in `pipeline/classify.py`, add the book's normalized title to the
`TITLE_MAP` dict with the desired genre, or add its author to `AUTHOR_GENRE`. Rebuild. (Run
`classify.py` alone first — it prints anything still `UNCLASSIFIED`.)

**Add a whole new genre** — pick a name, give it a color in `GENRE_COLORS`, and route books to it
via `classify.py`. Rebuild.

**Refresh from a new Goodreads export** — replace `goodreads_library_export.csv`, re‑run the
pipeline, and bump `SEED_VERSION` in `build.py`. See §5b.

**Reset read progress to the Goodreads baseline** — clear the site's `localStorage`; the
Goodreads seed re‑applies on next load. (To wipe to *nothing*, also empty `SEED_READ`/
`SEED_RATINGS` or set `SEED_VERSION` and skip the merge step.)

---

## 9. Data model

Each embedded book is a tiny object:

```js
{ id: 0, t: "Elfen Lied", a: "", g: "Manga", f: "eBook", pg: 184 }
//  id   title           author  genre     format               pages (drives dot size)
```

Read status, ratings, and theme are **not** stored in the file — they live in `localStorage`:
a book is "read" iff its `id` is in the array under `LS_KEY`; ratings are `{id: 1‑5}` under
`RATE_KEY`; theme under `THEME_KEY`. `pg` (pages) is a real number if `Books TBR.xlsx` has a
"pages" column (column D) or if the book came from Goodreads, otherwise a genre‑based estimate.

**Goodreads baseline:** at build time the pipeline bakes your Goodreads read/rating history into
two arrays in the file — `SEED_READ` (ids) and `SEED_RATINGS` (`{id: 1‑5}`). On first open the
app merges those into `localStorage` once (guarded by `SEED_VERSION`), so the map shows your real
read‑vs‑unread state out of the box. Clearing `localStorage` re‑applies the seed; your later
manual marks/ratings are preserved on top. See §5b.

---

## 10. Notes, limits, and future ideas

- **Genres are inferred**, not looked up per‑book online, so a few may be arguable. Fixing one is
  a one‑line addition in `classify.py` (§8).
- The map is a **designed layout**, not a true similarity embedding (like UMAP). It looks
  organic but positions are decorative, not semantic. A future version could compute real
  positions from book descriptions/embeddings and feed `x,y` per book into `build.py`.
- **Density vs. Open Syllabus:** their map has ~1M points; this has ~1,435, so it's inherently
  sparser. Dark‑mode glow and page‑based dot sizes give it richness.
- **Page counts are estimates** unless you add a real "pages" column (column D) to the
  spreadsheet — see §6. The estimate is genre‑based with overrides for famous long/short books.
- Possible future features: reading streak / stats view, tags, cover thumbnails, shareable link
  (publish as a hosted page), shelves/collections, per‑book notes, sort/filter by rating buckets.

---

*Built as a single self‑contained HTML file so it stays simple, private, and offline‑friendly.
Keep the "stable code" (`AVs-Library.html`) working; use the knobs in §7 to make it yours.*

---

## Recent Changes

- **AI Instruction Added**: Note for AI assistants added to keep changes minimal.
- **Dark Mode Orbit & Rotating Text**: When zoomed in, stars display book details (author, genre, page count) rotating around a circular orbit path.
- **Genre Cloud Highlight**: In both light and dark modes, hovering over a book now visibly lightens and highlights the entire background genre cloud it belongs to.
- **Light Mode Vibrancy**: Updated the light mode color palette to feature richer, more vibrant blues and cyan. Increased the visibility of light mode background clouds by 10% and the dot glow by 15% for a more prominent visual effect.
- **Cyan Genre**: Changed the "Science Fiction" genre color in the Dark Mode palette to cyan.
- **Electric Light Mode**: Light mode nodes ("neurons") now inherit their genre's color. They feature a deeply shaded core (now 1.5x larger for better visibility) ensuring they stand out perfectly against the bright glowing halos and white background. Connecting network lines are thicker and brighter.
- **Light Mode High-Contrast Adjustment**: Genres with extremely pale/white colors (e.g., Manga, Comics) automatically convert to a soft grey (instead of harsh black) in Light Mode, which creates a much smoother cloud and glow effect.
- **Dark Mode Orbit**: The circular orbit paths in Dark Mode have been made much more translucent (reduced opacity) so they don't overpower the glowing stars.
- **Dark Mode Title Legibility**: Dark-colored genre titles (like History, Biography, and Business) are now automatically brightened in Dark Mode so they remain easy to read against the deep space background.
- **Brain Outline Removed**: The programmatic brain outline drawing in Light Mode has been completely removed to focus purely on the floating neuron network aesthetic.
- **Read-Only Mode**: Interactive rating inputs, the "Mark as Read" buttons, and the manual "Import read list" tool have all been removed. The UI is now strictly a read-only viewer for the database (searching and filtering are still fully functional). All book imports now run through the backend Python scripts.
