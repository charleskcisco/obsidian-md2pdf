// MLA export template — derived from the MLA Handbook (9th ed.), not from
// refs/*.docx.
//
// This is the deliberate difference from journal.typ. That template
// transcribes the reference documents so the Typst and LibreOffice engines
// can be compared, and so it inherits a Turabian setup: no paragraph
// indent, an 18pt gap between paragraphs, a 1.53in top margin, a 20pt
// display face for headings, and a running head an inch down the page.
// Every one of those is wrong for MLA. Starting from the style guide
// instead of from the reference docs is what makes them not arise.
//
// Nothing here affects journal.typ, the docx chain, or any note that does
// not ask for it: this file is only read when a note sets template: mla.
//
// Parameters come from the note's YAML frontmatter via _typst_wrapper().

// ── Fonts ────────────────────────────────────────────────────────────────
//
// MLA asks for a legible face whose regular and italic differ clearly, and
// names Times New Roman as the usual choice. Student Macs and Windows
// machines both have it; the fallbacks are Times-metric so line breaks
// stay put on a machine that does not.
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
#let mono-fonts = ("Consolas", "DejaVu Sans Mono", "Liberation Mono")

// ── Vertical metrics ─────────────────────────────────────────────────────
//
// "Double-spaced" is a Word measurement before it is a typographic one,
// and a teacher's eye is calibrated to Word. For Times New Roman 12pt,
// Word's double spacing puts baselines 2.30em apart.
//
// Typst's leading is the gap *between* line boxes, so it depends on how
// tall the box is. conf() pins top-edge and bottom-edge to explicit
// lengths rather than to the "ascender"/"descender" keywords, because
// those keywords resolve to the font's *typo* metrics -- 0.693em for real
// Times New Roman, not the 0.891em hhea ascent -- which silently yields a
// 0.907em box and 2.10em spacing instead of 2.30em. Measured, not assumed.
//
// Explicit lengths also make the box identical across every fallback in
// body-fonts, so a machine without Times New Roman sets the same page.
//
//     0.891em + 0.216em = 1.107em line box
//     2.30em baseline-to-baseline - 1.107em = 1.193em leading
#let box-top = 0.891em
#let box-bottom = -0.216em
#let leading-double = 1.193em

// Headings that begin a reference list, and a near-enough plain-text
// reading of a heading for comparison. See turabian.typ.
#let _BIB_TITLES = ("bibliography", "works cited", "references")
#let _plain(body) = lower(repr(body).replace("[", "").replace("]", "")).trim()

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
  // ── Page ───────────────────────────────────────────────────────────────
  //
  // MLA: 1in on all four sides, and the running head half an inch from the
  // top of the page. header-ascent raises the header out of the body area
  // into the margin, so 1in - 0.5in leaves it sitting at the half-inch
  // mark. Worth checking on a printed page rather than on screen.
  set page(
    paper: "us-letter",
    margin: 1in,
    header-ascent: 0.35in,
    // Last name and page number, right-aligned, on every page including
    // the first -- MLA has no title page to exempt.
    header: if lastname != "" {
      set text(font: body-fonts, size: 12pt)
      align(right)[#lastname #context counter(page).display()]
    },
  )

  set text(
    font: body-fonts, size: 12pt, lang: "en",
    top-edge: box-top, bottom-edge: box-bottom,
  )

  // Double-spaced throughout, with NO additional space between paragraphs:
  // spacing equals leading, so a paragraph break is one more double-spaced
  // line and nothing else. The 18pt gap journal.typ adds comes from the
  // reference document's BodyText style and has no place here.
  //
  // first-line-indent applies to every paragraph -- `all: true` -- because
  // Typst otherwise skips the first paragraph after a heading or block,
  // while MLA indents that one too.
  set par(
    leading: leading-double,
    spacing: leading-double,
    first-line-indent: (amount: 0.5in, all: true),
    justify: false,
  )

  show raw: set text(font: mono-fonts)
  // Operators the serif does not carry.
  show math.equation: set text(font: math-fonts)

  // Headings sit at body size and weight. MLA does not prescribe a
  // heading scale for student papers and explicitly warns against making
  // them typographically loud; the 20pt display face journal.typ uses is
  // inherited from the reference docx, not from any style guide.
  show heading: set text(font: body-fonts, size: 12pt, weight: "regular")
  show heading: set block(above: leading-double, below: leading-double)

  // Block quotations are indented half an inch and stay double-spaced,
  // with no quotation marks added. The explicit block spacing matters:
  // Typst's default gap around a quote is wider than a line, which breaks
  // the even double-spaced rhythm exactly where a reader is most likely
  // to be checking it.
  show quote.where(block: true): set pad(left: 0.5in)
  show quote.where(block: true): set block(
    above: leading-double, below: leading-double)

  // ── Works Cited ────────────────────────────────────────────────────────
  //
  // Only reachable when a note sets bibliography:, which the student build
  // never does -- it is here so the template is correct MLA for anyone
  // upstream who does. pandoc's citeproc emits the list already formatted
  // and labelled <refs>; this places it, not the citation rendering.
  // A hand-typed reference list, which is the normal case: students type
  // these far more often than they keep a .bib. Its own page, centred
  // heading, hanging indent -- and styled regardless of the heading
  // level the writer happened to use.
  show heading: it => {
    if _BIB_TITLES.any(t => _plain(it.body) == t) {
      pagebreak(weak: true)
      block(above: leading-double, below: leading-double,
            align(center, text(weight: "regular", style: "normal", it.body)))
    } else { it }
  }

  show <bibentries>: it => {
    set par(hanging-indent: 0.5in, first-line-indent: 0pt)
    it
  }

  show <refs>: it => {
    pagebreak(weak: true)
    align(center)[Works Cited]
    set par(hanging-indent: 0.5in, first-line-indent: 0pt)
    it
  }

  // ── First page ─────────────────────────────────────────────────────────
  //
  // MLA has no title page. The first page opens with a flush-left,
  // double-spaced block -- name, instructor, course, date -- then the
  // title, centred, in plain title case. Body text follows immediately.
  //
  // These lines are not paragraphs in the running sense, so the 0.5in
  // first-line indent is switched off for them.
  {
    set par(first-line-indent: 0pt)
    let info = (author, instructor, course, date).filter(x => x != "")
    if info.len() > 0 {
      info.join(linebreak())
      linebreak()
    }
    // MLA keeps title and subtitle on one line, joined by a colon.
    if title != "" {
      let full = if subtitle == "" { title }
        else if title.ends-with(":") { title + " " + subtitle }
        else { title + ": " + subtitle }
      align(center)[#full]
    }
  }

  doc
}
