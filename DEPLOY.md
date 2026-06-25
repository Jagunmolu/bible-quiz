# Deploying the Bible Quiz (making it live)

The app is **static files only** (HTML/JS/PNG) - no server, no database. That
makes hosting free and trivial on several platforms. Recommended: **Netlify**
or **Vercel**. Both give you a public `https://` link in minutes.

You do NOT run Python on the host. The flow is:
1. Edit questions in the CSV on your computer.
2. Run `build_quiz_data.py` locally to regenerate `quiz-data.js`.
3. Push/upload the folder. The site serves the generated `quiz-data.js`.

---

## Option A - Netlify drag-and-drop (easiest, no Git needed)

1. Go to https://app.netlify.com/drop
2. Drag the whole `bible-quiz` folder onto the page.
3. Done - you get a live URL like `https://random-name.netlify.app`.
4. To update questions later: rebuild `quiz-data.js` locally, then drag the
   folder again (or set up Git below for auto-deploys).

`netlify.toml` is already included, so caching is configured correctly.

## Option B - Vercel or Netlify via GitHub (best for ongoing updates)

This gives you automatic redeploys every time you push a change.

1. Put the project in a GitHub repo:
   ```
   cd bible-quiz
   git init
   git add .
   git commit -m "Bible quiz app"
   git branch -M main
   git remote add origin https://github.com/<you>/bible-quiz.git
   git push -u origin main
   ```
2. **Vercel:** go to https://vercel.com/new, import the repo, framework
   preset = "Other", and deploy. `vercel.json` is already there.
   **Netlify:** https://app.netlify.com/start, pick the repo, deploy.
   `netlify.toml` is already there.
3. Every `git push` now redeploys automatically.

To update questions: edit the CSV, run the build script, commit, push.

## Option C - GitHub Pages (free, simple, slightly more manual)

1. Push the repo to GitHub (as above).
2. Repo Settings -> Pages -> Source = "Deploy from a branch" -> `main` / root.
3. Your site goes live at `https://<you>.github.io/bible-quiz/`.

Note: GitHub Pages ignores `vercel.json`/`netlify.toml`, so caching headers
won't apply - harmless for this use case. Add `cleanUrls` won't work either;
just visit `index.html` directly if the root doesn't load.

---

## Which should you pick?

For your lightweight use (a quiz screen at church, occasional question
updates): **Netlify drag-and-drop (Option A)** is the least friction - no Git,
no accounts to wire up, live in two minutes. If you expect to update questions
often and like version history, **Option B with GitHub** is worth the one-time
setup so updates are a single `git push`.

---

## Important: the password is not real security

The Host-view password lives in `quiz-data.js`, which is downloaded to every
visitor's browser. Anyone who opens the page can in principle read it in the
source. Once the site is **public on the internet**, treat the password as a
guard against *accidental* projection of answers, not as protection against a
determined person.

If you want the answers to be genuinely private on a public URL, that needs a
small backend (so answers are fetched only after a real login) - a different
build. Tell me if you ever want that; for an in-room church quiz the current
setup is normally fine.

## Tip: two screens from one URL

Once live, open the same URL twice:
- On the projector: leave it in **Projection** (answers hidden).
- On your laptop: switch to **Host** (enter the password) to see answers and
  drive the timer/scores.

They are independent - each browser keeps its own mode and used-question marks.
