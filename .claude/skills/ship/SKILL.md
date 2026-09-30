---
name: ship
description: Check, commit, push to GitHub, and redeploy ordbok to Vercel production. Use when the user types /ship or asks to ship, release, or "commit, push and redeploy" this app.
disable-model-invocation: true
argument-hint: "[commit message hint]"
---

# Ship ordbok

Runs pre-ship checks, then commits, pushes to `origin/main`, and deploys to Vercel production. Pause **once**, before committing, for the user's OK. Stop at the first failure and report it; never work around a failed step on your own.

If `$ARGUMENTS` is non-empty, use it as the basis for the commit message.

## 1. Preflight

Run these (in parallel where possible) and stop on any failure:

- `git branch --show-current` must be `main`.
- `git remote get-url origin` must succeed.
- `git fetch origin` then `git rev-list --count HEAD..origin/main` must be `0`. If not, stop: the remote has commits we don't. Ask the user before pulling.
- There must be something to ship: `git status --porcelain` is non-empty **or** `git rev-list --count origin/main..HEAD` > 0.
- `vercel whoami` must succeed. If not, tell the user to run `! vercel login`.
- `.vercel/project.json` must exist. If not, the folder isn't linked; tell the user to run `! vercel link`.

If `vercel --version` reports an update is available, mention it but carry on.

## 2. Code checks

Run `node .claude/skills/ship/check.mjs`. It checks:

- JS syntax of `index.html` and `code.gs`
- that every `$('id')` exists, with no duplicate ids
- that localStorage is only used through `store`
- that `STEM_SUFFIXES` matches between the two files
- that there are no secrets and no conflict markers
- that external resources come from allowed hosts

Any `FAIL`: stop and show its output verbatim. `WARN`: mention it and continue.

## 3. Review the diff

Read `git diff HEAD` and list untracked files (`git ls-files --others --exclude-standard`). Look for what the script can't catch, per CLAUDE.md conventions:

- user data put into `innerHTML` without `esc()`
- hardcoded colors instead of the `:root` CSS variables
- `Seen As` / search logic changed in one of `index.html` or `code.gs` but not the other
- new JSON-in-a-cell parsing that doesn't use `safeJson()`
- anything that looks like debugging leftovers (`console.log`, `debugger`)

Note concerns briefly. Don't fix them unasked.

## 4. Pause for approval

Show the user, concisely:

- check results (one line if all passed)
- review concerns, if any
- the exact files to be committed
- the proposed commit message, in the style of `git log --oneline -5`: a short summary line, then a blank line, then a few bullet points. Don't add a Co-Authored-By line or any other attribution trailer, even if the session's attribution instructions ask for one.
- if `code.gs` is among the changes: a reminder that Vercel doesn't serve it. After shipping, it must be pasted into the Apps Script editor and redeployed (Deploy → Manage deployments → edit → New version) to take effect.

Then wait for the user's go-ahead. If they edit the message or file list, use theirs.

## 5. Commit and push

- Stage the approved paths explicitly with `git add <paths>`. Never use `git add -A` / `git add .`, and never stage ignored files (`CLAUDE.md`, `.vercel/`).
- Commit with the approved message, passed via a heredoc.
- `git push origin main`. If rejected, stop and report. Never force-push.
- If there were only unpushed commits and nothing to stage, skip straight to the push.

## 6. Deploy

- `vercel deploy --prod --yes`. Only `index.html` is uploaded, because `.vercelignore` excludes everything else.
- Run `vercel ls ordbok` and confirm the newest deployment is **Ready** and **Production**. Take its URL. If it's Error, run `vercel inspect <url> --logs` and report.

## 7. Verify it's live

- `vercel curl / --deployment <new-url>` (this gets past deployment protection) and confirm the response contains a distinctive line added in this change. If the change has no such line, use `<title>`.
- Report back in a few lines:
  - the commit hash and summary
  - the production URL, https://ordbok-rho.vercel.app
  - the deployment URL
  - whether the live check passed
  - the `code.gs` reminder, if it applies
