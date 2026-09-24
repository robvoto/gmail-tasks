# AGENTS.md

Minimal always-loaded repository instructions.

- Never guess or invent repository state.
- Challenge assumptions and proposals when evidence, logic, risk, or project constraints warrant it. Do not agree by default or optimise for validating the human; optimise for correctness and better decisions. Do not be contrarian when the evidence supports agreement.
- Inspect current files and `git status` before editing.
- Preserve unrelated or untracked work; never stash, reset, overwrite, or commit another agent's changes without explicit coordination.
- Before any branch/worktree, commit, push, PR, merge, or `main`-integration action, use `.skills/git-lifecycle/SKILL.md`.
- Keep changes small and scoped.
- For work spanning multiple files or likely to run for a while, work in bounded batches: state the current batch, complete and verify it, report progress, then continue.
- If a patch, exact-text replacement, or expected match fails, reread the current source and diagnose the mismatch before retrying. Do not retry stale input.
- Before declaring a required connector/tool/source unavailable, inspect the capabilities exposed by that required connector/tool first.
- Do not claim completion without validation evidence.
- Before any semantic/product/UX/business-rule/default/workflow/data-interpretation/classification/heuristic/fallback/persistent-data behaviour change: investigate, explain the current finding and exact proposed effect, then wait for Rob's explicit approval. Treat uncertain changes as semantic; mechanical no-behaviour changes may proceed.
- Never claim a preference, rule, memory, or instruction is persisted unless the authoritative persistent source was actually updated and verified.
