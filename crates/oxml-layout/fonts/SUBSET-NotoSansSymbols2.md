# Noto Sans Symbols 2 bullet subset

Source: `ofl/notosanssymbols2/NotoSansSymbols2-Regular.ttf` from the Google
Fonts `main` branch at `51303ca9e8ac9dcea7b12d307ba568fd0e6fcfca`, retrieved
2026-10-09.

Source SHA-256:
`7d5fb73b7ca67a6798101741f5d280a3d016a56a197afcd4199dbb57b4b82a21`

Output SHA-256:
`7f5a31d932df09902c1097af020569b63d6296c067bc804ae6a597ef9b3ad21b`

The repertoire is the Unicode equivalents of the Symbol and Wingdings bullet
glyphs that no bundled Latin face draws: ◆ ◻ ☑ ☒ ✓ ✗ ❍ ❑ ❒ ❖ ➔ ➢ ⇨ ⧫ ⬥ ⬧.
Reproduce it with FontTools 4.66.1:

```text
pyftsubset NotoSansSymbols2-Regular.ttf --output-file=NotoSansSymbols2-bullets-subset.ttf --unicodes=U+21E8,U+25C6,U+25FB,U+2611,U+2612,U+2713,U+2717,U+274D,U+2751,U+2752,U+2756,U+2794,U+27A2,U+29EB,U+2B25,U+2B27 --glyph-names --symbol-cmap --legacy-cmap --notdef-glyph --notdef-outline --recommended-glyphs --name-IDs=* --name-legacy --name-languages=* --layout-features=* --no-hinting
```
