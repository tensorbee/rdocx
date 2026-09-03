# F-228, all aspects, pass 2

**Reviewed**: uncommitted worker diff after pass 1 remediation, 11 files and
approximately 2,800 added or changed lines including the untracked grammar
module
**Verdict**: 5 defects, 0 smells, 0 nitpicks

## Defects

### D1, three property writers violate OfficeMath schema order

`crates/rdocx-oxml/src/math.rs:1361`

The n-ary writer emits `subHide` and `supHide` before `limLoc` and `grow`, but
`CT_NaryPr` orders the supported children as `chr`, `limLoc`, `grow`,
`subHide`, and `supHide`. The delimiter writer at line 1457 emits `endChr`
before `sepChr`, and the run-property writer at line 1758 emits `sty` before
`lit`, `nor`, and `scr`. Authored combinations can therefore produce invalid
`xsd:sequence` order.

### D2, matrix base justification uses the column-justification domain

`crates/rdocx-oxml/src/math.rs:1020`

`m:baseJc` uses the vertical `top`, `center`, and `bottom` domain. The model
instead accepts and writes `left`, `center`, and `right`, which belongs to
matrix-column `m:mcJc`. A valid top or bottom matrix baseline is discarded and
the public authoring surface can emit invalid base-justification values.

### D3, math run breaks read and write the wrong attribute

`crates/rdocx-oxml/src/math.rs:1735`

`m:brk` carries its optional alignment point in `m:alnAt`, not `m:val`. The
parser reads `m:val` through `root_value`, and the generic property writer emits
`m:val`. It also cannot represent a present unaligned `m:brk`. Existing breaks
lose meaning and authored aligned breaks are invalid.

### D4, modeled math text loses leaf extensions and accepts excess text nodes

`crates/rdocx-oxml/src/math.rs:460`

The run writer always creates a fresh `m:t`, so unsupported attributes owned by
the modeled text leaf disappear on any save. The shape check at line 2114 also
accepts multiple `m:t` children and collapses them into one, even though the
run grammar has one required math text child. The parser must fail closed on
that malformed cardinality and preserve extensions on the one modeled leaf.

### D5, n-ary parsing invents hide defaults when properties are absent

`crates/rdocx-oxml/src/math.rs:1312`

The parser initializes from the authoring constructor, whose absent lower and
upper arguments set both hide flags. When source `m:naryPr` is absent, the
parsed flags therefore remain true even though omitted `subHide` and `supHide`
default to false. Saving the unchanged equation then authors both hide flags
and changes its semantics.

## Smells

None.

## Nitpicks

None.

## Not found

Pass 1 recursion, preservation, decoding, and production-shape defects are
resolved. No additional paragraph ordering, settings ownership, legacy
boundary, public-facade, dependency, or structural findings were found beyond
the defects above.
