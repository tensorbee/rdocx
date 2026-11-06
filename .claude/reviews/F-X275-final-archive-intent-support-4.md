# F-X275 affected archive intent, support pass 4

Reviewed clean base d4325a12 and its two-line archive-intent successor.
Defects: 0. Smells: 0. Nitpicks: 0.

Only published rdocx comments.rs and document.rs change against producer15424ace.
Binding test sources are not published. Manifests and dependency lock are unchanged.
The actual affected package is rdocx alone.

Root inspected both retained archives, authenticated their hashes and independently
normalized the VCS member using the existing no-newline JSON convention at
`scripts/readme_doctests.py:1292`. The actual final archive is
4447904 compressed bytes, 56099504 normalized member bytes and 37 members.
The recorded initial compressed value is4447874, giving30bytes within the existing
64byte tolerance. The lock retains pulldown-cmark0.13.6 and the archive is below10MiB.

AST comparison proves only the rdocx tuple changes at
`scripts/readme_doctests.py:405`. Root reviewed the corresponding single README
row at `README.md:44`. The other21tuples, dates and speed claims remain unchanged.
The writer's pre-Cargo selector/import preparation errors and restored predecessor
are retained. They are not represented as product RED or successful checks.

Evidence: `/private/tmp/rdocx-s96-archive-intent-20261106-h3jy9eml/word-comment-cache-successor-1`,
including actual archives, child logs, SOURCE_BOUNDARY.json and ROOT_REPACKAGE_CHECK.json.

Accept the two-line intent update. Full clean-tree README/example and verified
publication dry-run applicability remain pending. No source changes in this review.
