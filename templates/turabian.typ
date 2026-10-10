// Turabian export template — derived from A Manual for Writers (9th ed.),
// not from refs/*.docx.
//
// journal.typ already produces a Turabian-shaped document, because the
// reference .docx it was transcribed from is Charles's Turabian setup. The
// difference here is that the measurements come from the manual rather
// than from one person's Word styles, and that the things the manual is
// specific about -- an unnumbered title page, a half-inch paragraph
// indent, single-spaced block quotations -- are actually implemented.
//
// Nothing here affects journal.typ, the docx chain, or any note that does
// not ask for it: this file is only read when a note sets
// template: turabian.
//
// Parameters come from the note's YAML frontmatter via _typst_wrapper().

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

// See mla.typ for the derivation and for why the line box is pinned with
// explicit lengths rather than the "ascender"/"descender" keywords.
#let box-top = 0.891em
#let box-bottom = -0.216em
#let leading-single = 0.043em
#let leading-double = 1.193em

// Baseline to baseline for double-spaced body text: the line box plus
// its leading. Named because the heading spacing is expressed in whole
// lines rather than in points.
#let line-advance = 2.30em

// Headings that begin a reference list, lowercased for comparison.
#let _BIB_TITLES = ("bibliography", "works cited", "references")

// The text of a heading, near enough for comparison. Typst has no
// plain-text extractor, and a recursive walk cannot accumulate into an
// outer variable -- those are read-only inside a function. repr() of a
// simple heading is "[Bibliography]", so stripping the brackets gets
// there for the only case this needs to recognise.
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
  set text(
    font: body-fonts, size: 12pt, lang: "en",
    top-edge: box-top, bottom-edge: box-bottom,
  )

  // ── Title page ─────────────────────────────────────────────────────────
  //
  // Turabian puts the title about a third of the way down, with the
  // author, course, instructor and date lower on the same page. The title
  // page carries no page number and is not counted in the numbering the
  // reader sees, so it is emitted as its own page before numbering starts.
  //
  // Single-spaced regardless of the body setting: this is a block of
  // labels, not prose.
  set page(paper: "us-letter", margin: 1in, numbering: none, header: none)
  {
    set par(leading: leading-single, spacing: leading-single,
            first-line-indent: 0pt)
    // The manual asks for the title about a third of the way down the
    // page. 1in margin + 2.667in puts it at 3.667in on us-letter. This is
    // the one place turabian.typ deliberately departs from journal.typ,
    // whose 1.667in comes from the reference .docx.
    v(2.667in)
    // Turabian's title page rule: a subtitle goes on the line below the
    // title, and the title line ends with a colon. A title typed with
    // its own colon keeps just the one.
    if subtitle != "" {
      let lead = if title.ends-with(":") { title } else { title + ":" }
      align(center)[#lead#linebreak()#subtitle]
    } else {
      align(center)[#title]
    }
    v(2.5in)
    align(center)[
      #let info = (author, course, instructor, date).filter(x => x != "")
      #info.join(linebreak())
    ]
  }
  // Zero the counter *before* the break. A header reads the page counter
  // as the page is laid out, so an update placed after the break has
  // already been overtaken and the first body page comes out numbered 2.
  counter(page).update(0)
  pagebreak(weak: true)

  // ── Body pages ─────────────────────────────────────────────────────────
  //
  // A set rule on page takes effect from the next page, which is exactly
  // the boundary the pagebreak above just created.
  //
  // page.numbering is deliberately left unset: setting it makes Typst
  // generate its own centred footer, which lands a *second* page number
  // at the bottom of every page alongside the one in the header. The
  // counter is displayed explicitly instead, with its own pattern.
  // Turabian permits the page number either top right or bottom centre.
  // Bottom centre keeps the top margin clear, which reads better on a
  // paper with no running head.
  set page(
    footer-descent: 0.35in,
    footer: context {
      set text(font: body-fonts, size: 12pt)
      align(center)[#counter(page).display("1")]
    },
  )

  // Double-spaced body, no extra gap between paragraphs, half-inch indent
  // on every paragraph including the first after a heading.
  set par(
    leading: leading-double,
    spacing: leading-double,
    first-line-indent: (amount: 0.5in, all: true),
    justify: false,
  )

  show raw: set text(font: mono-fonts)
  // Operators the serif does not carry.
  show math.equation: set text(font: math-fonts)

  // Headings follow the CMOS presentation the Purdue OWL sample paper
  // demonstrates, not Turabian's prescribed five-level scheme.
  //
  // That is a deliberate choice and a legitimate one: Chicago does not
  // prescribe a heading format. It asks for a clear hierarchy applied
  // consistently -- "Chicago allows you to devise your own format but
  // prioritizes consistency" -- and leaves the rest to the writer.
  // Turabian, the student adaptation, is where the centred levels 1 and
  // 2 come from. A paper marked strictly against Turabian would expect
  // those; a paper marked against Chicago will not.
  //
  //   1  flush left, bold
  //   2  flush left, italic
  //   3  flush left, regular
  //   5  run in at the head of its paragraph, bold, ending in a period,
  //      with the text continuing on the same line
  //
  // Sizes stay at body size throughout: the hierarchy is carried by
  // placement and weight, which is what keeps it from reading as a
  // magazine layout.
  show heading: set text(font: body-fonts, size: 12pt, weight: "regular")

  // An extra line space before a heading, per the same guidance. The
  // body advances 2.30em baseline to baseline, so adding one to the
  // usual block gap puts a full blank line above the heading and
  // nothing extra below it.
  show heading: set block(
    above: leading-double + line-advance, below: leading-double)

  // set rules, not content rules: returning styled content yields inline
  // material, which drops the heading's block and runs it into the
  // paragraph beneath. Level 5 is the one place that is wanted.
  show heading.where(level: 1): set text(weight: "bold")
  show heading.where(level: 2): set text(style: "italic")
  show heading.where(level: 5): it => box(strong[#it.body.])

  // A bibliography starts its own page. Students type these by hand far
  // more often than they use a .bib, and the docx chain has always done
  // this via _lua_basic_filter -- the Typst path only ever handled the
  // <refs> block citeproc emits, so a typed heading got no break at all.
  show heading: it => {
    if _BIB_TITLES.any(t => _plain(it.body) == t) {
      // Styled here rather than left to the level rules, because the
      // level is the student's arbitrary choice of # or ## and should
      // not decide how a reference list is headed. Typed by hand is the
      // normal case, so it has to look right however it was typed.
      pagebreak(weak: true)
      block(above: leading-double, below: leading-double,
            text(weight: "bold", style: "normal", it.body))
    } else { it }
  }

  // The entries themselves: hanging indent, single-spaced within an
  // entry and a full line between them. _wrap_bibliography puts them in
  // a labelled block before pandoc sees the markdown, because a Typst
  // show rule cannot reach forward from a heading to the content after
  // it -- setting par inside a show par rule does not affect the
  // paragraph being shown, and rebuilding the paragraph recurses.
  show <bibentries>: it => {
    set par(leading: leading-single, spacing: leading-double,
            hanging-indent: 0.5in, first-line-indent: 0pt)
    it
  }

  // Block quotations are indented and set single-spaced, which is one of
  // the few places Turabian departs from double spacing throughout.
  show quote.where(block: true): set pad(left: 0.5in)
  show quote.where(block: true): set block(
    above: leading-double, below: leading-double)
  show quote.where(block: true): set par(
    leading: leading-single, spacing: leading-single, first-line-indent: 0pt)

  // ── Bibliography ───────────────────────────────────────────────────────
  //
  // Only reachable when a note sets bibliography:, which the student build
  // never does. Entries are single-spaced with a blank line between them
  // and a half-inch hanging indent.
  show <refs>: it => {
    pagebreak(weak: true)
    align(center)[#strong[Bibliography]]
    set par(leading: leading-single, spacing: leading-double,
            hanging-indent: 0.5in, first-line-indent: 0pt)
    it
  }

  doc
}
