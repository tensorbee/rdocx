# F-X280b, final union, Root pass 1

**Verdict**: 1 defect, 0 smells, 0 nitpicks. This final-gate finding reopens bounded remediation, not an independent peer-cap reset.

## Defects

### D1, Pre-index ZIP acceptance rejects the valid pinned trailing-newline PPTX

`crates/oxml-opc/src/package.rs:1073`

The pre-index end-record predicate requires the ZIP comment end to equal the physical file end. The pinned Maloney corpus PPTX has one trailing newline after a valid EOCD with no comment, so this new guard refuses it before the provider ZIP index runs. The approved compatibility scope limits new refusals to damaged input and invalid authoring, and does not authorize refusing this valid existing corpus file.

Actual full all-feature workspace tests at a78eba3c fail six chart corpus cases on this one file with valid end of central directory not found. The file is corpus/pptx/au.asn.aes.www_conferences_2011_presentations_Fri_20Room4Level4_20930_20Maloney.pptx,536086bytes, SHA256426582c9408e74cd061f2b4dc01fd6b8ddfbb55dda4fdf419e3dca313d107a7f. Root independently read its PK local header, valid EOCD suffix with one newline and Python zipfile acceptance. The remaining three failed chart cases concern absent local soffice and are separate environment failures, not this code finding. Complete stdout/stderr and failed child exit remain under final-union-gates/a78eba3c0242ed7f8e08b7acbe4a1882b0f06648.

Restore this existing provider-compatible acceptance while retaining physical entry-count limits, duplicate decoded part-name refusal, malformed/truncated rejection and ZIP64 controls. Use an actual generated trailing-newline positive and duplicate-name negative alongside this pinned corpus observation. Do not remove or rewrite the corpus, hide the failures, bypass the count guard or open a deeper security assessment. No public API or resource-limit choice is requested. The original F280b implementation peers remain exhausted, with affected Root final review after bounded repair.

## Smells

0.

## Nitpicks

0.

## Not found

No broader ZIP assurance finding established in this finite check. The audit ends here before remediation.
