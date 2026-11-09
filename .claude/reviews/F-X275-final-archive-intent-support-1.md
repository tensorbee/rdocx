# F-X275, final archive intent support 1

**Verdict**: 0 defects, 0 smells, 0 nitpicks. Accept the pre-freeze archive-record refresh. No full documentation, publication dry-run or release result follows.

Reviewed23tracked changed paths:22README archive rows and the two archive dictionaries in scripts/readme_doctests.py. Root independently compared the script AST with HEAD and found only ARCHIVE_MEASUREMENTS and ARCHIVE_REMEASUREMENT_DATES changed. Historical speed constants, dates, claims and other logic remain exact. Root compared all27README nonarchive content against HEAD and authenticated their recorded final file hashes.

Authenticated actual retained22archives and independently recalculated normalized member bytes, member counts and compressed sizes using the existing VCS normalization. Every applicable observation has exact member/count agreement, less than10MiB compressed and at most64bytes compression difference. Dates are the actual2026-11-06 recording date. The initial pass succeeded on stable source with the target override removed, matching the existing producer's actual target/package lookup.

The first post-adoption oxml-pdf observation failed exact member bytes because its own README row gained one byte. That failed observation remains in POST_ADOPTION_MEASUREMENTS.json. The subsequent tuple correction used actual measured bytes, and the affected package alone was repackaged successfully. OXML_PDF_FINAL.json and its retained archive authenticate102218compressed bytes,478926normalized member bytes and14members against recorded102219/478926/14. No measured value was calculated from a desired pass or selected from retries. The later package is the applicable successor, not a relabeled initial failure.

Evidence root is /private/tmp/rdocx-s96-archive-intent-20261106-h3jy9eml. SOURCE_BOUNDARY.json SHA256 is a334e6b056cc5320457a6264e422d4c9a146f82dd8000f53ca0fe68a4fdbd67c. Full README examples, clean publication dry-run and all final union gates remain pending at final source. Later material package changes earn affected remeasurement. No source product behavior, threshold, sampler or performance assertion changes in this diff.

## Defects

0.

## Smells

0.

## Nitpicks

0.

## Not found

No archive-recording, preservation, test-policy or scope defect established. No Cargo or measured performance operation was run by this review. Only this review file was written and the pass ends here.
