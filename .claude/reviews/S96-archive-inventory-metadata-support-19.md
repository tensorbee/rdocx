# S96 rdocx archive measurement metadata support19

Independent read-only support for the two current measurement metadata hunks. Zero defects, zero smells and zero nitpicks. Root remains sole writer. No Cargo, compilation, tests, package creation, UI, network or tracked-file edit was performed in this pass.

## Scope and exact producer custody

HEAD e430071d75ed33e8831752d644aec3c0a865d3a3. The reviewed diff changes only the rdocx archive numbers at README.md:44 and the corresponding tuple at scripts/readme_doctests.py:405. Version0.17.0, date2026-11-08, platform, command, member count and other measurement rows remain unchanged. The separate pending code/HLD/workflow edits are not re-reviewed here. BOUNDARY.json binds the exact current metadata/source hashes. REVIEWED_METADATA.patch retains the two hunks. ARCHIVE_INSPECTION.json records every archive member hash and its comparison with the current source.

The original measurement receipt RDOCX_CURRENT_MEASUREMENT.json records actual cargo package --locked --offline --allow-dirty --no-verify exit0 in a distinct rdocx-current-inventory target. Its retained archive independently hashes to dcf442afb8e6c8bc31c54f3287fb21de9164722c72559bb4c0f9dfed4291f811. Independent tar/gzip inspection reproduces(4501793,56396845,37).

The fresh post-row receipt RDOCX_POST_ROW_MEASUREMENT.json records the same measurement-only command in another distinct target, actual exit0 and no_verification_build_claim true. Its actual archive independently hashes to807d148893ed4222837be7388a1b60af015194395afb2df3756298204faecf23. Independent inspection reproduces exactly(4501793,56396845,37), matching both the receipt's recorded tuple and the current two metadata values.

## Source applicability and Cargo normalization

For both archives, all33 directly source-mapped members other than README match the current files exactly. They include every packaged Rust source, current repaired comparison.rs, integration_test.rs, the existing regression entrypoint, examples and Cargo.toml.orig. The original archive's only directly source-mapped mismatch is its README's prior numeric row. Its README and the current README are both16845bytes. In the post-row archive all34 directly source-mapped members, including README, are byte-identical to current source. No current source behavior is inferred from historical package proof.

The remaining three members are Cargo-generated Cargo.toml, Cargo.lock and .cargo_vcs_info.json. The normalized manifest identifies rdocx0.17.0, edition2024 and Rust1.93, points to packaged README.md, preserves the current feature table and has no dependency paths. Its shared dependencies use0.15.0 and Word dependencies0.17.0. Cargo.toml.orig equals the current package manifest. Generated Cargo.lock records rdocx0.17.0. These are normalized package observations, not a dependency compilation/compatibility pass.

The actual physical tar member-size sum is56396864. The measurement function at scripts/readme_doctests.py:1292 excludes dirty=true and substitutes40zeroes for the VCS SHA before counting that member's normalized JSON size. Its VCS member changes from125physical bytes to106normalized bytes, producing the independently reproduced56396845. The compressed count is the actual gzip file length and the member count is37. Normalization does not mask Rust source, README, manifest or lock changes.

## Self-measurement and unchanged tolerance

The README is itself included in the rdocx crate because crates/rdocx/Cargo.toml:13 names ../../README.md. Updating the row changes those bytes. Equal-length numeric substitutions keep README length, member size/count and tar layout stable but do not prove compressed size. The second actual post-row archive supplies that missing observation: its compressed length is also4501793. Its SHA differs from the original as expected, while the final normalized tuple exactly matches the new recorded value.

The existing ARCHIVE_COMPRESSION_TOLERANCE_BYTES remains64 at scripts/readme_doctests.py:394. Validation at scripts/readme_doctests.py:1349 continues to require exact normalized member bytes/count and compressed size below10MiB. Under the pinned compiler it requires an absolute compressed difference of at most64. Outside the pinned compiler it permits smaller outputs but limits upward difference to64. No tolerance or gate logic changed in these two hunks. The actual post-row difference from the recorded compressed value is0, so no allowance is needed for this measured artifact. A future regenerated archive remains subject to the existing gate.

validate_package_archive also requires the sole packaged README to equal current source at scripts/readme_doctests.py:1393. Independent read-only inspection confirms that condition for the actual post-row archive. It is deliberately not claimed for the original pre-row archive, whose historical bytes and receipt remain intact.

## Limits and pending gates

The two tuple changes accurately record the authenticated current-source measurement. They do not convert --no-verify into a verified source-package build. Root's source-package verification build and applicable final Rust, archive, installed, hosted and final-union gates remain separate. Root reported four affected metadata unit tests passed, but their raw receipt was not part of this bounded inspection, so this report takes no independent test credit from that statement. The earlier Support17/18/native/SDK evidence is neither altered nor expanded here.

An initial read-only inspection script attempted getmembers after its tar context closed and raised OSError before reporting results. The corrected read-only inspection retained the member list while open. No artifact or product file was changed by that inspection error. All conclusions above use the completed authenticated inspection.

End metadata support19 and return control to Root.
