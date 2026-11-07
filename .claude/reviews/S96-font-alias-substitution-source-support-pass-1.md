# S96 bounded font-alias substitution source review

Verdict: 0 defects, 0 smells, 0 nitpicks in the assigned diagnostic repair. No Cargo, product import, timing or source mutation was performed by this reviewer.

F278 at `.claude/plans/F-X278-design.md:165-167` limits the new Word diagnostic to absent requested families and keeps it out of PowerPoint's existing diagnostic list. The repair at `crates/oxml-layout/src/font.rs:1887-1903` follows the actual explicit alias target and suppresses a false missing-family fact only when the selected physical database ID is a member of that target's actual caller-label or embedded-label set. This is identity-based availability, not suppression based merely on a matching label. The physical-family comparison at `:1910` uses that same resolved target. Resolution order and selected FontId do not change.

The existing diagnostic tier classification and separate Word-only vector at `:1913-1946` remain intact. An unrelated fallback face does not satisfy the label-ID check. The existing caller-label test at `:4568` now covers a label-derived alias, an explicit alias and a missing-family diagnostic, while asserting the old shared/PowerPoint diagnostic list remains empty. These are source controls, not claimed executions in this pass.

Runtime regression15 and focused font controls remain for Root's actual source-pinned batch. This review does not qualify broader historical substitution behavior, appearance baselines or final font/performance acceptance. It supplies no new gate or review-cap reset.

Source SHA256: `ec37e321714d850852c278af262ab0ff2bfe74dca1bd8df30a140690d29a4469`.
