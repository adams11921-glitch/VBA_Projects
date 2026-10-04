---
description: Put back the previous version of the add-in
---
The last change didn't work out. Please put the add-in back the way it was:

1. Show me in one plain sentence what the last saved change was (`git log -1`) and whether there are also unsaved
   (uncommitted) changes.
2. Ask me to confirm. Then:
   - if there are uncommitted changes, discard only those (`git restore src ribbon` and `git checkout -- dist`), or
   - otherwise revert the last commit with `git revert --no-edit HEAD` (never rewrite history).
3. Ask me to close Word, then run `build\Update-AddIn.ps1` so Word gets the restored version.
4. Add a one-line note to docs/LESSONS.md under "Things that went wrong": what we tried and why it was undone.
   Commit that note.

Details from me (may be empty): $ARGUMENTS
