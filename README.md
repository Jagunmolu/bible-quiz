# Bible Quiz — Host + Projection

A church Bible-quiz tool styled for **Okota Baptist Church**. One app, two modes,
one data file, and a script that builds the data file from a spreadsheet.

## Files

| File | What it is |
|---|---|
| `index.html` | The whole app. Open it in a browser. |
| `quiz-data.js` | The questions (generated — see below). |
| `build_quiz_data.py` | Converts a CSV/Excel question bank into `quiz-data.js`. |
| `hebrews_questions.csv` | Example question bank you can copy and edit. |
| `logo.png` | Church seal. |

## Running it

Open `index.html` in any browser. No build step, works offline.

For a projector setup: open the file on the projector and leave it in
**Projection**; open it again on your laptop and switch to **Host**.

---

## 1. Updating questions from a spreadsheet (no more editing JS by hand)

Edit `hebrews_questions.csv` (or your own Excel file), then run:

    python build_quiz_data.py hebrews_questions.csv --book "Hebrews"

That regenerates `quiz-data.js`. Reload the app — done.

### Spreadsheet format

One row per question. Columns (case-insensitive, any order):

| Column | Required | Notes |
|---|---|---|
| `category` | yes | short code or label — see list below |
| `question` | yes | the question text |
| `answer` | yes | the answer text |
| `sn` | no | optional number; use `--sort-by-sn` to order by it |

**Category codes** (map to on-screen labels automatically):
`general`, `fill` (Fill in the Gap), `tof` (True or False),
`spell` / `golden` (both map to Spelling), `recitation`. Anything else
becomes its own category,
title-cased. To change or add codes, edit `CATEGORY_MAP` near the top of
`build_quiz_data.py`.

**Line breaks inside a cell:** type a literal \n where you want a break
(handy for the Recitation rules). The script turns it into a real line break.

**Instruction rounds (like Recitation):** any category whose code contains
`recit` is treated as an instructions card — it shows one card with the text
visible immediately, instead of a numbered grid with hidden answers. Put the
round title in `question` and the rules in `answer`.

### Shuffling the question order

By default questions appear in the same order as the CSV. To randomise the
on-screen order **within each category** (so the tiles aren't predictable):

    python build_quiz_data.py acts_quiz_1_14.csv --shuffle

The CSV file itself stays tidy and ordered - only the generated `quiz-data.js`
(what the app shows) is shuffled. To get the *same* shuffle again (e.g. so the
projector and your laptop match, or to reproduce a set you liked), pin a seed:

    python build_quiz_data.py acts_quiz_1_14.csv --shuffle --seed 42

Same seed = same order, every time. Different seed = different order.

### Switching books

    python build_quiz_data.py acts_questions.csv --book "Acts" --version NIV

Nothing in `index.html` changes — it's fully driven by `quiz-data.js`.

### Options

    --book "Name"        Book name shown in the header (else guessed from filename)
    --version NIV        Bible version (default NIV)
    --church "..."       Church name (default: Okota Baptist Church)
    --password "..."     Host-view password (default: okota)
    --out quiz-data.js   Output file
    --sort-by-sn         Order questions within a category by the sn column
    --shuffle            Randomise the on-screen order of questions within
                         each category (the CSV file stays in tidy order)
    --seed N             Seed for --shuffle so the same order can be
                         reproduced, e.g.  --shuffle --seed 42

Excel input needs openpyxl: `pip install openpyxl --break-system-packages`
(or just save the sheet as `.csv`).

---

## 2. Host-view password

To stop answers being projected by accident:

- The app **always opens in Projection** (answers hidden — safe to show).
- Switching **to Host view asks for a password**. Default is `okota`.
- Switching back to Projection is always free.
- Once you've entered the password, it won't ask again for that browser
  session (so flipping between views mid-quiz is frictionless).

Change the password with `--password` when you build, or edit `hostPassword`
in `quiz-data.js`.

**Honest limitation:** this is a static web page, so the password lives in the
file and a determined person could read it in the source. It's designed to
prevent the *accidental* "oops, the answers are on the projector" mistake — not
to defend against someone deliberately digging through the code.

## 3. The two modes

| | Host view | Projection |
|---|---|---|
| For | You, the moderator | The audience screen |
| Answer | Shown with the question | Hidden until you reveal it |
| Scoreboard | Visible | Hidden |
| Password | Required to enter | None (default view) |

Used-question marks and scores survive a refresh within the same browser session.

## Keyboard shortcuts (when a question is open)

- **Space** — reveal the answer (projection)
- **Enter** — mark the question used and return to the board
- **Esc** — return without marking it used
- **T** — start the 5-second timer (press again for +5s)

## Timer

A large countdown ring sits at the top-right of the question view. Press **T**
or click **Start timer**. At zero it turns red, pulses, shows **TIME**, and a
full-screen **TIME!** flashes briefly so the whole room sees it. Click
**Restart timer** to reset.

## Scoring

Host view has a scoreboard. Add teams, tap +/− (steps of 5 to match the round
rules), click a name to rename.

---

## One answer to verify

In **Spelling** (Hebrews set), the "shortest chapter in Hebrews" answer is flagged
`(verify against your NIV)`. Hebrews has 13 chapters / 303 verses total
(confirmed), but please confirm the shortest-chapter claim against your own
Bible before using it on screen.
