# F-228, all aspects, pass 3

**Reviewed**: uncommitted worker diff after pass 2 remediation, 11 files and
approximately 2,900 added or changed lines including the untracked grammar
module
**Verdict**: 5 defects, 0 smells, 0 nitpicks

## Defects

### D1, duplicate and malformed property leaves can disappear

`crates/rdocx-oxml/src/math.rs:2003`

Property preservation claims every recognized local name as modeled. The
writer emits at most the first source leaf for each modeled field, so a second
`m:type`, `m:chr`, or similar leaf is dropped instead of retained raw. The
global and run-property parsers at lines 1616 and 1737 have the same duplicate
loss. They also claim recognized leaves whose value failed to parse, which can
replace malformed producer markup with a typed default. Only the first valid
occurrence may own the projection. Every duplicate or malformed occurrence
must remain raw.

### D2, an omitted n-ary character acquires summation semantics

`crates/rdocx-oxml/src/math.rs:1304`

The n-ary parser initializes the character to U+2211 SUMMATION. OfficeMath
defines U+222B INTEGRAL when `m:naryPr/m:chr` is omitted. Saving an equation
that relies on the schema default therefore changes both its XML and meaning.

### D3, malformed display equations become empty typed paragraphs

`crates/rdocx-oxml/src/math.rs:116`

`CT_OMathPara::from_raw` does not validate its required sequence. It accepts an
`m:oMathPara` with no `m:oMath` child, fills an empty vector, and serializes a
different invalid typed subtree. It must require optional `m:oMathParaPr`
followed by one or more `m:oMath` children and leave malformed paragraph-level
input raw at the paragraph boundary.

### D4, empty nested text elements and truncated text are accepted

`crates/rdocx-oxml/src/math.rs:2240`

`element_text` returns success for any nested empty element and for EOF before
the closing `m:t`. A producer payload such as `<m:t>x:empty/></m:t>` becomes
empty or partial typed text instead of failing closed, while a truncated
standalone parse can also succeed.

### D5, standalone canonical output can contain undeclared producer prefixes

`crates/rdocx-oxml/src/math.rs:2029`

The parser carries inherited namespace bindings while classifying raw content,
but root serialization declares only `m` plus attributes declared on the
OfficeMath root itself. If an unknown attribute or retained subtree uses a
prefix declared on the original paragraph, calling `OfficeMath::to_xml` or
moving that model to another document emits an undeclared prefix. Required
inherited bindings must be retained and replayed at the canonical standalone
root without allowing a source `m` binding to override the fixed namespace.

## Smells

None.

## Nitpicks

None.

## Not found

All pass 1 and pass 2 findings are otherwise resolved. No new facade ordering,
settings relationship, legacy-boundary, dependency, or structural defects were
found beyond the five items above.
