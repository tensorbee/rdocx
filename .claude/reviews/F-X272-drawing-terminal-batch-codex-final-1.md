# F-X272, Drawing terminal family, Root final pass 1

Reviewed the complete1486line increment, two production files and progress,
patch `458dd01472869bd05c1a26cc9167de513e6afedcc9e070c458d259f70e55f612`.
This bounded Root pass does not reset exhausted implementation peers.
Verdict: one defect, zero smells and zero nitpicks. No product edit, test or
build occurred during this review. The counterexample below is source-derived,
not a claimed compiled RED.

## Defects

### D1, decoded namespace facts can certify an error the actual reader ignores

`crates/rdocx-oxml/src/document.rs:1472`

constructor_namespace_name resolves against CaptureDeclarations. Those bindings
come from raw_namespace_declarations, whose decoded_and_normalized_value call
at text.rs:26609 unescapes namespace URI character references. The actual Drawing
reader uses NsReader namespace resolution at drawing.rs:1899. Pinned quick-xml
0.41 NamespaceResolver::push passes the raw attribute value into Namespace, and
add retains those raw bytes. The actual reader does not use this decoded URI
view. The new source rule consequently treats some physical declarations as
qualified docPr or decorative owners when the real reader ignores them.

Use an otherwise admitted inline drawing with an Empty q:docPr whose own local
xmlns:q spells the WordprocessingDrawing URI with `200&#54;` instead of `2006` and
whose id is bad, followed by an ordinary wp:docPr id1 under the literal namespace.
The raw physical local declaration stays inside the captured drawing. NsReader
does not recognize the first URI and takes the valid later metadata owner.
The new decoded source lookup recognizes the first owner and reads its bad id,
certifying a false constructor failure. With an unrelated deep typed sibling,
that false failure can mark the enclosing control raw and suppress checked work.
The separate first-metadata postscan has the same raw-versus-decoded mismatch.

Preserve the actual reader's context and namespace-resolution policy, including
the distinction between physically retained local declarations and inherited
bindings emitted by capture closure. Do not change the real parser's namespace
policy globally, add a new refusal or weaken the existing resource equation.
Add actual reader/admission discriminators before calling this path qualified.
The remaining lost-binding/context checklist does not justify a false negative
fact in an enabled certification path.

## Not found

No additional actionable defect was established. The shared event/scalar actions
have real Inline/Anchor consumers. Consumed versus walked children, local-name
early ends, required ids, decorative depth and first metadata selection were
examined. The typed-prefix guard correctly retains its explicit unproved state
instead of erasing earlier resource demand. Source owner links avoid repeated
ancestor scans. The introduced277pass/1fail prefix regression and corrected
unchecked-parser test premise remain honestly retained.

The reported279passes and scoped gates do not prove the missing context case,
whole Drawing/C1 closure, provider completion or final native/release acceptance.
The review ends and returns to orchestration for separate remediation.
