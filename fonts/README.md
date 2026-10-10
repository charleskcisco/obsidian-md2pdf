# Bundled fonts

De Gruyter Serif and De Gruyter Sans Math, from
<https://gitlab.com/degruyter-public/font/de-gruyter-sans_serif>.

Licensed under the **SIL Open Font License, Version 1.1** — see `LICENSE`,
which the licence requires to travel with the font. OFL permits bundling,
redistribution and embedding in a PDF, which is what this directory is for.

## Why these are here

Export used to ask for Times New Roman and fall back through Tinos,
Liberation Serif and TeX Gyre Termes. On a minimal writerdeck **every one
of those misses**, so typst silently chose its own default and the output
was not the typography the templates describe:

    warning: unknown font family: times new roman
    warning: unknown font family: tinos
    warning: unknown font family: tex gyre termes

Bundling removes the guess. Every machine now sets the same page, whether
or not it has any system fonts at all.

De Gruyter Serif is derived from Noto SemiCondensed and covers Greek,
Greek Extended, Cyrillic, IPA and combining diacriticals — which matters
for a vault whose bibliography carries titles like
*Τὸ εὐαγγέλιον τὸ εὐαγγελισθὲν ὑπ' ἐμοῦ*.

Sans Math covers mathematical and logical operators the serif does not.

## What changed as a consequence

It is **not** Times-metric — narrower, being SemiCondensed. Vertical
spacing is unaffected, because the templates pin the line box with
explicit lengths rather than the font's own ascender and descender. Line
breaks and page counts do move, and `folio.typ` no longer matches the
LibreOffice render as a result. That parity was deliberately traded for
output that is identical everywhere.
