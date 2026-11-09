# Support27, bounded archive metadata advisory

Zero findings in correctness, provenance, scope and claim wording. This is a tmp-only advisory on the four archive metadata hunks, not a formal implementation pass or cap change. I did not implement these hunks, run Cargo/tests/UI, modify tracked files or review the independent colour/F289/F290 product contracts.

The three actual archive hashes match ARCHIVE_MEASUREMENTS.json at :64, :133 and :202 under `.claude/scratch/S96-final-evidence/s96-release-freeze-20261108-1/native-refresh/colour-archive-inventory-1`. Each retained command exits0at :60, :129 and :198and explicitly uses locked/offline/allow-dirty/no-verify with22localpatch entries. Logs show Packaging/Packaged counts, not verification. This is inventory, NOT a publish dry-run or verified archive qualification.

Independent read-only tar inspection reproduces all three normalized triples: rdocx4501789/56396845/37, rdocx-layout463844/2543127/15and rdocx-oxml1006737/6177983/32. README.md:45, crates/rdocx-layout/README.md:24and crates/rdocx-oxml/README.md:23match those exact values and2026-11-09dates. The corresponding scripts/readme_doctests.py dates at :379/:382/:384and tuples at :405/:408/:410agree. No other script behavior changes appear in the four-file diff.

The normalized member-byte policy is unchanged at scripts/readme_doctests.py:1292-1310. It removes only git.dirty and replaces git.sha1with40zeros before counting formatted VCS JSON bytes. rdocx has no dirty field and no byte delta. The layout and OXML raw tar-member sums2543146and6178002include19extra VCS bytes each; normalized sums2543127and6177983are correct. This is the documented normalization, not omitted payload. Archive.member_count remains37/15/32.

All direct source/archive payload members match current worktree files, except the expected packaged README self-measurement rows, which exactly match their pre-update HEAD bytes. Generated normalized Cargo.toml/Cargo.lock/VCS owners are identified separately rather than falsely mapped to source files. Cargo.toml.orig maps to the actual source manifest. The per-member hashes/classification are in AUTHENTICATION.json. The archive README self-row update is later metadata and does not retroactively change the retained archive hash.

AST literal comparison against HEAD confirms exactly3of22measurement tuples and3dates changed. The other19PACKAGE entries remain unchanged, not19tar members. ARCHIVE_COMPRESSION_TOLERANCE_BYTES remains64at scripts/readme_doctests.py:394. No exact-member/count equality or compression guard was relaxed. Existing recorded performance rows and historical evidence remain unchanged.

The four current SHA256pins below bind this advisory. Full clean-source affected archives with actual verification remain due under the existing release workflow. Future self-row/clean-VCS compression must be measured against the unchanged guard, not assumed from these no-verify observations. No test, package verification, publication readiness or final union pass is claimed.

- `README.md`: `3113ffd61537cd144717045f3e67d88496282d6f4d57002cdd8d8f7a811a60ee`
- `crates/rdocx-layout/README.md`: `0e30a2361568813d242765d22c284a1cf8fb1fd94ad288f602b888f09be9e7b3`
- `crates/rdocx-oxml/README.md`: `60f39c4eba4848076abc74142580496103e1f735b602d8ace3c515739d305a38`
- `scripts/readme_doctests.py`: `00e253b011a988ba635bb7462e5a71496e34e0b5c9e7bbcd2d6bc5150baf5618`
