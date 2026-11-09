# S96 archive inventory metadata support15

Zero defects, zero smells and zero nitpicks. The four-file metadata correction is suitable for the source-availability renewal checkpoint. This is inventory review, not verified publication packaging or a complete Docs/CI pass. Root remains sole writer. No Cargo, product runtime, network or source mutation occurred in this pass. Only this report was written.

The dirty diff changes exactly three archive tuples and three dates in scripts/readme_doctests.py, plus their three displayed README rows. No measurement logic,64byte tolerance, version, speed/memory row or product source changes. The actual CI Docs log identifies these same three stale tuples at `.claude/scratch/S96-final-evidence/s96-release-freeze-20261108-1/ci-docs-job-log.txt:466`.

I independently authenticated all three final .crate SHA-256 values against FINAL_METADATA_ARCHIVE_MEASUREMENTS.json, opened their tar members, recomputed the exact existing VCS-normalized byte totals/counts, and compared their embedded README bytes and all61 packaged src files with current source. Each measurement command records exit0 and the existing22local patches, with --locked/--offline/--allow-dirty/--no-verify. Source HEAD is2d24393cd02e353d5e6be535b073103215a0c33c. Reference: `.claude/scratch/S96-final-evidence/s96-release-freeze-20261108-1/FINAL_METADATA_ARCHIVE_MEASUREMENTS.json:2`.

| Package | Recorded tuple | Final actual tuple | Compressed delta |
|---|---|---|---:|
| rdocx |4500166,56385989,37|4500167,56385989,37|+1|
| rdocx-layout |463380,2540005,15|463379,2540005,15|-1|
| rdocx-oxml |1006331,6174933,32|1006331,6174933,32|0|

Normalized member bytes and member counts match exactly. All compressed sizes are below10MiB and within the unchanged64byte tolerance. The three dates are2026-11-08. The recorded tuples correspond to the retained post-metadata observations, so the final one-byte compression variation is neither a newly selected tolerance nor an unexplained source delta. References: `scripts/readme_doctests.py:394`, `scripts/readme_doctests.py:402`, `scripts/readme_doctests.py:1292`, `scripts/readme_doctests.py:1341`, `README.md:44`, `crates/rdocx-layout/README.md:23`, `crates/rdocx-oxml/README.md:22`.

The OXML self-row adjustment accounts for the earlier pre-metadata normalized6174931bytes becoming6174933after its embedded README tuple gained two bytes. Final archive and current README agree exactly. The first clean-tree packaging refusal remains an actual exit101 in ARCHIVE_DIRTY_REFUSAL.json and is not called a successful measurement. The later explicitly allow-dirty/no-verify inventory operation has its own actual success receipts.

These observations correct stale archive inventory only. They do not qualify cargo publish verification, final package compilation, remaining CI jobs, registry publication or release approval. No additional product gate or numerical rebenchmark is requested. This pass ends here and returns control to Root.
