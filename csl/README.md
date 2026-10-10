# Bundled citation styles

Default CSL styles, chosen by a note's `style:` frontmatter. A note's own
`csl:` always wins; these are only the fallback, so that `style: chicago`
produces Chicago notes rather than pandoc's built-in author-date default.

| `style:` | File |
|---|---|
| `chicago`, `turabian` | `chicago-notes-bibliography.csl` |
| `mla` | `modern-language-association.csl` |

Both files are unmodified copies from the Citation Style Language project
(<https://github.com/citation-style-language/styles>) and are licensed
under Creative Commons Attribution-ShareAlike 3.0
(<http://creativecommons.org/licenses/by-sa/3.0/>), as recorded in each
file's own `<rights>` element. They are redistributed here unchanged.

- Chicago Manual of Style 18th edition (notes and bibliography) — the
  full-note variant (NB 13.18), not the shortened-note one.
- MLA Handbook 9th edition (in-text citations).
