# S96, Root archive metadata support, pass 1

Reviewed three changed metadata paths: `scripts/readme_doctests.py`, `README.md` and `crates/rpptx/README.md`. Two package tuples and their two displayed rows change. No measurement tolerance, speed claim, date, policy algorithm or other package row changes.

Defects: 0. Smells: 0. Nitpicks: 0.

The actual two-package producer completion is SHA `78a919ec5fc2ba7c6ade0ef542f7271e4c5462596e0adab0c84f0c6e69830a2f`. The reviewed adoption patch is `dfff76471160aacc16d979b7464c0ec03ab3582259e02018889a02e2b8a9d095`. Root inspected and applied its exact diff. Actual post-adoption completion is `2c99518d7f82243f852e8d83ff4ec1291044f2568f3c381acbfff9e54e0fee30`.

The rdocx tuple is 4,475,293 compressed bytes, 56,239,208 normalized member bytes and 37 members. The rpptx record is 799,413 compressed bytes, 4,246,845 member bytes and 16 members. Post-adoption rdocx matches exactly. Post-adoption rpptx differs by ten compressed bytes, within the existing 64-byte allowance, with exact normalized member bytes and count. Source, tests, README and manifest members were authenticated against the actual inputs. All 76 source inputs and tool identities were stable through the executions.

Five actual metadata/package/tool children and eight focused README policy cases passed. Local patch, measurement evidence and both archive validators passed. Expected mutation-test error output is retained rather than mistaken for a product failure. Other 20 package records are unchanged. The H10 failure is preserved, and no hosted successor result or final release qualification is claimed.
