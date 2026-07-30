# 10, Bindings spec

Owners: `oxml-py-support`, `rdocx-py`, `rpptx-py`, `rdocx-wasm`, `rpptx-wasm`,
`oxml-cli-support`, `rdocx-cli`, `rpptx-cli`.

## The PyO3 lifetime problem

A `#[pyclass]` must be `'static`. The facade is built on borrow handles:
`Paragraph<'a> { inner: &'a mut CT_P }`, plus consuming builders and
`Document::add_paragraph(&mut self) -> Paragraph<'_>` which holds the document
mutably borrowed for the handle's whole life. Python additionally requires that
`p = doc.add_paragraph("x")` stay usable across arbitrary later mutations,
including ones that reallocate the content vector.

References are out, categorically. Four options were weighed:

| Option | Verdict |
|---|---|
| **Index and path handles** re-resolving on every call | **chosen** |
| `Rc<RefCell<_>>` or `Arc<Mutex<_>>` in the core | rejected: rewrites every crate, pollutes the Rust API with borrow noise for users who never touch Python, and `Rc` is not `Send` so `allow_threads` is lost |
| Arena with generational ids | correct long-term, but converts the content vectors across every crate. Deferred |
| A separate owned mirror API | rejected: doubles the API surface, and "attach" reintroduces the identity problem |

### The chosen design

```rust
pub enum PathSeg { Body(usize), Row(usize), Cell(usize),
                   Para(usize), Run(usize), Slide(usize), Shape(usize) }
pub struct ContentPath { pub segs: SmallVec<[PathSeg; 5]>, pub revision: u64 }

#[pyclass(name = "Document")]
struct PyDocument { inner: rdocx::Document, revision: u64 }

#[pyclass(name = "Paragraph")]
struct PyParagraph { doc: Py<PyDocument>, path: ContentPath }
```

Zero change to the Rust API. No interior mutability leaking into the core.
Aliasing is checked by PyO3's own `RefCell` on the pyclass, so a violation is a
clean `RuntimeError`, never undefined behaviour. Resolution is a handful of
vector index operations, negligible against FFI overhead.

`Shape(usize)` repeats for nesting into groups, so
`shape.text_frame.paragraphs[i].runs[j]` is one path.

### The invalidation problem, handled loudly

An index path addresses a **position**, not an object. After
`doc.remove_content(1)`, a handle to paragraph 3 would silently read what used
to be paragraph 4. python-docx does not have this problem because it holds an
lxml element pointer that follows the element.

v0.1 therefore carries a **document revision counter**, bumped by every
structural mutation and captured by every handle at construction. A mismatch
raises:

```
rdocx.StaleElementError: paragraph handle was created at document revision 4,
but the document is now at revision 5 (a structural change invalidated it).
Re-fetch it with doc.paragraphs[i].
```

**Loud failure beats silently wrong data.** There are no snapshot accessors that
keep working after invalidation.

v0.2 upgrades to lazily-assigned stable ids backed by `w14:paraId`, which OOXML
already defines for exactly this purpose, so they round-trip to disk and improve
DOCX fidelity as a side effect. Then a handle survives unrelated removals and
matches python-docx semantics, with no API change.

### Two supporting decisions

**Collections are lazy.** `doc.paragraphs` is a pyclass holding only
`Py<PyDocument>` and implementing `__len__`, `__getitem__` with negative and
slice support, and `__iter__`. `Document::paragraphs() -> Vec<ParagraphRef>` is
never called from the binding.

**Consuming builders are bypassed.** A `fn bold(mut self, val: bool) -> Self`
cannot back a Python property setter. The facade exposes 61 non-consuming
`set_*` twins: 24 on `Paragraph`, 19 on `Run`, and 18 across `Table`, `Row`, and
`Cell`. The existing builders delegate to them. The surface is additive, and a
borrowed nested handle can mutate without a rebind:
`doc.paragraph_mut(3).unwrap().add_run("text").set_bold(true)`.

**Threading.** `Document` remains `Send` and `Sync`. Its normal and
deterministic layouts live in separate `Mutex<Option<Arc<LayoutResult>>>`
caches, with a compile-time regression gate preserving that contract.
`to_pdf`, `render_all_pages` and `to_bytes` run inside `py.allow_threads`, so a
Python thread pool genuinely parallelises work across documents. Concurrent
rendering of one document shares the immutable cached result after the first
layout for that font mode. That is a capability python-docx has no equivalent
for.

## Python API shape

**Drop-in compatibility is an explicit non-goal. Source compatibility for the
documented API is an explicit goal.**

python-docx's real-world surface is inseparable from lxml, and a large fraction
of production code reaches through `._p`, `._r`, `doc.element.body`, `qn()` and
`OxmlElement`. Promising drop-in means promising an lxml-shaped shadow API that
can never be delivered, and every gap then reads as a bug.

What is promised: *if your code uses only the documented python-docx API, it
works unchanged.* Backed by a compatibility suite built from python-docx's own
documentation examples, and by making a touch of `._p` raise a clear
`NotImplementedError` naming the attribute and its equivalent rather than an
`AttributeError` five frames away.

```python
from rdocx import Document, Inches, Pt, RGBColor, WD_ALIGN_PARAGRAPH

doc = Document("in.docx")
p = doc.add_paragraph("Hello")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run(" world")
r.font.bold = True
r.font.size = Pt(18)
doc.add_picture("img.png", width=Inches(2))   # height inferred by oxml-media
doc.save("out.docx")
doc.save_pdf("out.pdf")                        # documented as an rdocx extension
```

- `font` and `paragraph_format` are themselves handles, so `r.font.bold = True`
  writes through the chain.
- **Tri-state properties return `None` for inherit**, `True` or `False` when
  explicit. rdocx's `Option<bool>` already matches. Never collapse `None` to
  `False`.
- `Length` subclasses `int` and returns EMU, matching `docx.shared.Length`, with
  `.inches`, `.cm`, `.mm`, `.pt`, `.emu` and `.twips`. This detail decides
  whether copy-pasted code works.
- Enums are pure-Python `IntEnum` shims so `WD_ALIGN_PARAGRAPH.CENTER == 1`
  holds and they carry docstrings.
- `RdocxError(Exception)` is the base, with `PackageError`, `XmlError`,
  `StaleElementError` and `LayoutError` beneath it.

`rpptx` mirrors python-pptx the same way.

## Packaging

**maturin, mixed Rust and Python layout**, so type stubs and enum shims have a
home. `python-source = "python"`, `module-name = "rdocx._rdocx"`,
`features = ["pyo3/extension-module"]`.

**abi3-py39.** One wheel per platform rather than one per interpreter version,
so roughly 6 wheels instead of 48. The cost is marginally slower attribute
access and no free-threaded build under abi3. Start abi3-only and revisit only
if profiling shows attribute overhead matters.

Matrix: `manylinux_2_28` x86_64 and aarch64, `musllinux_1_2` x86_64, macOS
x86_64 and arm64, Windows x86_64, plus an sdist.

Two traps specific to this workspace:

- **`fontdb`'s `fontconfig` feature is useless on musl and Windows.** Gate it
  per-target.
- **Build wheels with `bundled-fonts` on.** Otherwise `to_pdf()` produces blank
  or mangled text on a bare manylinux container with no system fonts, which
  would be the single most common support question. Roughly 4 MB per wheel is a
  fair trade.

Type stubs are hand-written with a `py.typed` marker, kept honest by a CI job
running `mypy --strict` and `stubtest`. Do not auto-generate them from PyO3.

**Distribution names `rdocx` and `rpptx`**, import names identical. The binding
crates are `publish = false`, because a cdylib has no business on crates.io.

## CI

`wheels.yml` on a **`py-v*` tag namespace**, separate from `publish.yml` on
`v*`, so a Rust patch release does not rebuild twelve wheels and a binding-only
fix does not force a crates.io release. Publishing uses PyPI trusted publishing
via OIDC, with no long-lived token in secrets.

**A PR-time job that builds the wheel and runs pytest is mandatory.** The
absence of exactly this job for wasm is why `rdocx-wasm` rotted.

The parity suite is worth more than any number of Rust-side assertions: write a
document with rdocx, open it with python-docx, assert text, styles and tables
survive, then the reverse. python-docx as a CI dev dependency is free.

## WASM

### The existing crate is a fork, not a binding

`rdocx-wasm` holds only `CT_Document` and `CT_Styles`. `from_bytes` stores
`package_bytes` and immediately marks it `#[allow(dead_code)]`. `to_docx_bytes`
discards it and calls `OpcPackage::new_docx()`.

Round-tripping any real document through it **silently destroys** every image
and its relationships, headers and footers, numbering, settings, the theme, the
font table, footnotes and endnotes, core and app properties, every content-type
override, and every relationship except the styles one it re-adds. It has no
tests, no CI job, and `publish = false`, so nothing has ever caught it.

### The fix

```rust
#[wasm_bindgen]
pub struct WasmDocument { inner: rdocx::Document }
```

Everything round-trips immediately, because `to_bytes` flushes into the
**original** package. Three blockers and their answers:

- `Document::open` uses `std::fs`, so expose only `fromBytes` and `toBytes`.
  `save()` is meaningless in a browser anyway.
- `FontManager::new()` loads system fonts and `fontconfig` will not build for
  `wasm32-unknown-unknown`. Add a `system-fonts` feature, default on, off for
  wasm, with `bundled-fonts` on instead. **Then `to_pdf()` works in the
  browser**, which is a genuinely compelling capability that is absent today.
- Watch `getrandom` creep. The workspace already trims `zip` features to avoid
  it.

Keep the existing JS method names so current users do not break. The semantics
only become correct.

**The actual fix is the CI job**: `cargo check --target wasm32-unknown-unknown`
plus `wasm-pack test --node`. The code drifted because nothing was watching.

`rpptx-wasm` wraps the real facade from day one, never a mini-model, in two
profiles: a default without rendering at roughly 600 KB gzipped, and a `render`
build with the rasteriser and bundled fonts at several MB.

## CLIs

`rpptx-cli` mirrors `rdocx-cli`: `inspect`, `text`, `convert`, `diff`,
`replace`, `validate`, `render`, using clap derive and `serde_json` for
`--json`, including the pattern of dispatching `validate` separately so its exit
code carries the verdict.

Two presentation-specific additions: **`thumbnail`**, slide one at a fixed size,
which is what every CMS wants, and **`outline`**, the title and bullet tree,
which is ideal for LLM ingestion and is a genuine differentiator.

`validate` is the highest-value command and pays for itself in the test suite by
running across the corpus in CI.

Shared plumbing, range parsing, output-path defaulting and the JSON envelope,
lives in `oxml-cli-support` rather than being copy-pasted. **Version the JSON
envelope from the first release**: `{"schema": 1, ...}`.
