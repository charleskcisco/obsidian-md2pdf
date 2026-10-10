// Journal export template — the Typst equivalent of refs/{single,double}.docx.
//
// Every measurement below is transcribed from the reference .docx files that
// the pandoc/LibreOffice chain uses, so the two engines produce comparable
// output. Word stores lengths in twips (1440 = 1in) and half-points; the
// conversions are noted inline. Re-measure against a LibreOffice export
// before changing anything here.
//
// Parameters come from the note's YAML frontmatter via _typst_wrapper().

// ── Fonts ────────────────────────────────────────────────────────────────
//
// The reference docs ask for Times New Roman (body) and Aptos Display
// (headings). Typst bundles neither, and a minimal writerdeck has neither
// either — so list real names first and fall back through metric-compatible
// substitutes. Tinos and Liberation Serif are Times-metric; TeX Gyre Termes
// is the usual free Times clone. Aptos has no free equivalent, so headings
// fall back to the serif rather than to something arbitrary.
// De Gruyter Serif first, because it is the one face guaranteed to be
// present: it ships in fonts/ and typst is given --font-path. Everything
// after it is a fallback for a build that somehow lacks the bundle.
//
// It is not Times-metric -- Noto SemiCondensed underneath, so narrower --
// which moves line breaks and page counts. Vertical spacing does not
// move, because conf() pins the line box to explicit lengths rather than
// to whatever the font reports.
// Two backstops, not six. typst warns for every family it cannot find,
// even when an earlier one resolved, and those warnings land in the
// export log a person reads when something has gone wrong. Times covers
// macOS and Windows, Liberation covers Linux, for the case where the
// bundle is missing entirely.
#let body-fonts = (
  "De Gruyter Serif", "Times New Roman", "Liberation Serif",
)
#let math-fonts = ("De Gruyter Sans Math", "New Computer Modern Math")
#let heading-fonts = ("Aptos Display", "Aptos", ..body-fonts)
#let mono-fonts = ("Consolas", "DejaVu Sans Mono", "Liberation Mono")

// ── Vertical metrics ─────────────────────────────────────────────────────
//
// Typst's leading is the gap *between* line boxes, whereas Word's w:line is
// a multiple of the font's line height — so the two are not interchangeable
// and leading has to be derived from the line box.
//
// The line box depends on text.top-edge / bottom-edge, whose defaults are
// cap-height and baseline — a box of only ~0.66em. conf() pins them to
// explicit lengths below so the box is 1.107em (ascender 0.891em +
// descender 0.216em), which is what these two constants are derived
// against. Change one, change the other.
//
// The lengths are explicit rather than the "ascender"/"descender"
// keywords because those resolve to the font's *typo* metrics — 0.693em
// for real Times New Roman, not the 0.891em hhea ascent — which made the
// box 0.907em and put double spacing at 2.10em where LibreOffice renders
// 2.30em. Explicit lengths also make every fallback face in body-fonts
// set the same page.
//
//     single (w:line=240)  baseline-to-baseline 1.15em  -> leading 0.043em
//     double (w:line=480)  baseline-to-baseline 2.30em  -> leading 1.193em
#let box-top = 0.891em
#let box-bottom = -0.216em
#let leading-single = 0.043em
#let leading-double = 1.193em

// BodyText carries w:before=180 w:after=180 (9pt each). Word would add
// both between consecutive paragraphs, for 18pt — but the chain this
// template mirrors renders through LibreOffice, which collapses them to
// the larger of the two. Measured on an actual export, the gap between
// paragraphs is 27.6pt + 9pt, not 27.6pt + 18pt.
//
// Typst's par.spacing replaces the inter-line gap rather than adding to
// it, so the leading is folded back in to land on the same baseline
// distance.
#let para-extra = 9pt

// Compare a heading's text without caring how it was marked up.
#let _plain(body) = lower(repr(body).replace("[", "").replace("]", "")).trim()

#let _BIB_TITLES = ("bibliography", "works cited", "references")

#let conf(
  title: "",
  subtitle: "",
  author: "",
  course: "",
  instructor: "",
  date: "",
  style: "",
  spacing: "double",
  lastname: "",
  doc,
) = {
  // style, course, instructor and lastname are accepted and unused.
  // _typst_wrapper passes the same arguments to every template, and the
  // styles that needed them now have templates of their own.
  let leading = if spacing == "single" { leading-single } else { leading-double }

  // pgMar: left/right 1440tw = 1in; top 2204tw = 1.531in; bottom 2206tw =
  // 1.532in; header 1440tw = 1in from the page top.
  set page(
    paper: "us-letter",
    margin: (left: 1in, right: 1in, top: 1.531in, bottom: 1.532in),
  )

  // Pin the line box so leading-single / leading-double above are
  // measured against a known quantity rather than Typst's cap-height
  // default, and against the same quantity on every machine.
  set text(
    font: body-fonts, size: 12pt, lang: "en",
    top-edge: box-top, bottom-edge: box-bottom,
  )
  // BodyText carries w:ind w:firstLine="720" -- a half-inch indent on
  // every paragraph of the style, which Word applies to the first one
  // after a heading too, hence all: true. This was simply missed when the
  // reference docs were transcribed, so the two engines have never
  // wrapped their lines in the same places.
  set par(
    leading: leading, spacing: leading + para-extra, justify: false,
    first-line-indent: (amount: 0.5in, all: true),
  )
  show raw: set text(font: mono-fonts)
  // Operators the serif does not carry.
  show math.equation: set text(font: math-fonts)

  // Heading sizes and spacing from styles.xml (half-points -> pt, twips ->
  // pt at 20tw/pt). None of them are bold in the reference docs.
  show heading: set text(font: heading-fonts, weight: "regular")
  show heading.where(level: 1): set text(size: 20pt)
  show heading.where(level: 2): set text(size: 16pt)
  show heading.where(level: 3): set text(size: 14pt)
  show heading.where(level: 1): set block(above: 18pt, below: 4pt)
  show heading.where(level: 2): set block(above: 8pt, below: 4pt)
  show heading.where(level: 3): set block(above: 8pt, below: 4pt)

  // ── Bibliography ───────────────────────────────────────────────────────
  //
  // Citations and the reference list are rendered by pandoc's citeproc,
  // not by Typst, and arrive already formatted. pandoc emits the list as a
  // block labelled <refs>; this styles it to match what the Lua filter did
  // to the docx — page break before it, a heading, 0.5in hanging indent.
  //
  // Typst's own bibliography() cannot do this job under a note style: it
  // renders each citation as a *new* footnote, and notes in this vault put
  // citations inside footnotes already (^[Again, @key]). Footnotes do not
  // nest, so the citation silently vanished and left an empty note behind.
  // citeproc knows it is already inside a note and inlines the text.
  //
  // Declared before doc, since a show rule only affects what follows it.
  show <refs>: it => {
    pagebreak(weak: true)
    heading(level: 1, "Bibliography")
    set par(hanging-indent: 0.5in, first-line-indent: 0pt)
    it
  }

  // The same list, typed out by hand -- which is the ordinary case, since
  // most writers never touch citeproc. pandoc labels that block
  // <bibentries> (see _wrap_bibliography); a generated one arrives as
  // <refs> above. The entries need the opposite of the body's
  // first-line-indent: flush-left first line, continuations indented.
  // Without this rule they inherit the 0.5in first-line indent set for
  // prose and come out indented at the top and flush at the bottom --
  // upside down, and unmistakably wrong to anyone marking the paper.
  //
  // No heading is emitted here, unlike <refs>: the writer typed one.
  show <bibentries>: it => {
    set par(hanging-indent: 0.5in, first-line-indent: 0pt)
    it
  }

  // A reference list starts a new page. Matched on the heading's text
  // rather than its level, because whether it was typed as # or ## is the
  // writer's arbitrary choice and should not decide the layout.
  show heading: it => {
    if _BIB_TITLES.any(t => _plain(it.body) == t) {
      pagebreak(weak: true)
      it
    } else { it }
  }

  // ── Front matter ───────────────────────────────────────────────────────
  //
  // pandoc's --standalone renders title/author/date from the note's
  // metadata through the reference doc's Title, Author and Date styles,
  // so the docx chain has always produced a centred block here.
  //
  // Unconditional: mla and chicago route to their own templates now, so
  // every note reaching this one wants exactly this block. It used to be
  // guarded on `style == "" or style == "basic"`, which meant a note
  // whose style: was a typo -- anything that falls back to this template
  // -- silently lost its title, author and date altogether.
  //
  // Title is 28pt centred with w:after=80tw; Author and Date are 12pt
  // centred. The three gaps below are measured off a LibreOffice render
  // rather than derived, because Word's box model and LibreOffice's
  // interpretation of it do not agree closely enough to compute them:
  // the target is where the ink actually lands.
  //
  // Braced so the set rule stays inside it. These values are the front
  // matter's own -- single leading, no paragraph gap, no indent -- and
  // must not reach the body, which is double-spaced and indented. A set
  // rule runs to the end of its enclosing block, so at conf()'s top
  // level it would silently restyle the whole document.
  {
    set par(leading: leading-single, spacing: 0pt, first-line-indent: 0pt)
    if title != "" {
      align(center)[#text(size: 28pt)[#title]]
      v(5.2pt, weak: false)
    }
    if subtitle != "" {
      align(center)[#text(size: 18pt)[#subtitle]]
      v(5.2pt, weak: false)
    }
    if author != "" {
      align(center)[#author]
      v(4.5pt, weak: false)
    }
    if date != "" {
      align(center)[#date]
      // The body's first paragraph arrives with its own above-spacing
      // of leading + para-extra. In the docx the gap after Date is only
      // its line advance plus BodyText's 9pt w:before, which is smaller
      // -- and block spacing takes the maximum, so it cannot be reduced
      // by adding. Cancel the difference instead.
      v(9.5pt - (leading + para-extra), weak: false)
    }
  }

  doc
}
