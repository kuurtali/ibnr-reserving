# IBNR Reserving

![Tests](https://img.shields.io/badge/tests-21-brightgreen)
![Python](https://img.shields.io/badge/python-3.10+-3776AB?logo=python&logoColor=white)
![Method](https://img.shields.io/badge/method-chain%20ladder-8957e5)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A reserving tool for non-life insurance. It takes a claims table, builds
**development triangles per line of business**, estimates ultimate losses
and the IBNR reserve, and produces a reasoned Excel report.

> **Data policy.** This repository contains no real portfolio data and
> cannot. Everything ships with synthetic data generated from a known
> development pattern; real inputs are supplied by environment variable
> and never committed.

---

## The problem it solves

An insurer knows what it has paid. It does not know what it will
eventually pay — some claims are reported late, and reported claims
develop upward for years. The gap between the two is the reserve, and
getting it wrong is expensive in both directions: too low and the balance
sheet is a fiction, too high and capital sits idle.

**Chain ladder** answers this by reading history. If claims from 2020 grew
45% between their first and second quarter of development, and so did
2021, and 2022, then 2025's claims will probably do something similar.
Those growth ratios — *link ratios* — project every immature accident
period forward to its ultimate value.

```
                    development period →
accident   ┌─────────────────────────────┐
 period    │ 1,000  1,452  1,714  1,893  │   fully developed
   ↓       │ 1,100  1,597  1,886        ?│   ← projected
           │ 1,250  1,815      ?        ?│   ← from link ratios
           │ 1,400      ?      ?        ?│
           └─────────────────────────────┘
                                    ultimate − paid to date = IBNR
```

---

## Per line of business, and per claim type

Reserving is not one calculation. Different lines develop at completely
different speeds, and averaging them together destroys the signal.

The tool splits every calculation into **segments** and reports them
separately as well as combined:

| | |
|---|---|
| **Branch code** | Selected at run time from whatever codes exist in the data — motor liability, general liability, and so on. Each branch is reserved on its own triangle. |
| **Pool** | Separate pool arrangements are reserved separately. |
| **Material vs bodily injury** | The single most important split. Material damage closes in months; bodily injury develops for years and carries most of the tail. Reserving them together under-reserves the long tail and over-reserves the short one. |

The report shows each segment's triangle, factors, ultimate and IBNR, then
a combined total — so the number can be defended one segment at a time
rather than as a single figure with no working shown.

---

## What the calculation actually does

```
claims table
  → filter to branch / pool, split material vs bodily injury
  → cumulative paid + outstanding  =  incurred triangle
  → mask cells beyond the valuation diagonal
  → volume-weighted link ratios  (all history / last 5y / last 3y)
  → tail factor  (manual, or bounded log-linear extrapolation)
  → project → ultimate → IBNR
  → factor-window scenarios + backtest
  → Excel report with triangles, factor tables and charts
```

### Choices worth naming

**Incurred, not paid.** The triangle is cumulative paid *plus* outstanding
case reserves, so the tool uses what the claims handlers already know
rather than waiting for payment.

**Volume-weighted link ratios**, `Σ C(i,j+1) / Σ C(i,j)` — each accident
period is weighted by its own size. A simple average lets a small, noisy
period move the factor as much as a large one.

**The valuation diagonal is masked before projection.** Cells beyond the
valuation date are zeroed first. Skip this and a few stray forward-dated
records silently corrupt every factor downstream.

**The tail is bounded.** Automatic extrapolation fits log-linear decay to
the trailing factors, projects 20 steps forward, and **caps at 1.5**; if
factors aren't decaying it refuses and returns 1.0. Unbounded
extrapolation is the most common way a tail factor goes wrong.

**Salvage and subrogation sign is configurable**, because recovery records
arrive with different conventions depending on the source system, and
hard-coding the sign is a classic way to be quietly wrong.

**It backtests itself.** The tool re-runs the valuation 4, 8 and 12
quarters in the past and compares what it *would* have estimated then
against how those accident periods actually developed. A reserving tool
without a backtest is an opinion.

**Scenarios, not a single number.** All three factor windows are computed
side by side with percentage deviation from the base. How much the answer
moves when the method moves is part of the answer.

---

## Is it right?

Two different questions, answered two different ways.

### Does the arithmetic work?

`data/uret.py` builds claims from a **known** development pattern, so the
factors the tool recovers can be checked against the factors used to
create the data:

```
  d   expected    computed        diff
  0    1.45238    1.45238     1.3e-09
  1    1.18033    1.18033     6.2e-10
  2    1.10417    1.10417     5.0e-11
  ...
```

`tests/test_chain_ladder.py` asserts this, plus: a tail of 1 leaves the
ultimate equal to the last column, a larger tail raises the ultimate,
cells outside the diagonal stay zero, the tail estimator refuses
non-decaying factors and respects its cap, and every input validator
rejects bad input.

### How accurate is the estimate?

Synthetic data cannot answer that — for accuracy you need triangles whose
future is already known.

The **CAS Loss Reserve Database** provides exactly that: Schedule P
triangles for several hundred US insurers across six lines of business,
**including the lower triangle**, so each accident year's actual
development is on record.

```bash
python3 data/cas_indir.py --brans ppauto     # download + convert
python3 data/cas_indir.py --brans wkcomp --sirket 43
```

The converter maps CAS's cumulative paid and incurred columns onto the
tool's record schema — paid becomes incremental rows, `incurred − paid`
becomes the outstanding — and holds back the lower triangle as the answer
key. `tests/test_cas_donusum.py` verifies the round trip: rebuilding
`cumulative paid + outstanding` from the converted records reproduces the
original incurred triangle exactly.

One caveat, stated because it matters: CAS data is **annual** while this
tool is quarterly. Accident years are mapped onto consecutive period
codes, so the arithmetic is correct but the report labels one year as one
"quarter".

*Data source: [CAS Loss Reserving Data Pulled from NAIC Schedule
P](https://www.casact.org/publications-research/research/research-resources/loss-reserving-data-pulled-naic-schedule-p),
compiled for the CAS by S&P Global Market Intelligence; earlier sets
compiled by Glenn G. Meyers and Peng Shi. Downloaded by the script above,
not redistributed here.*

---

## Running it

```bash
pip install -r requirements.txt

python3 data/uret.py     # build the synthetic sample
python3 ibnr.py          # run — GUI if available, otherwise CLI
pytest -q                # 21 tests
```

Point it at your own data:

```bash
set IBNR_GIRIS_DOSYASI=D:\path\claims.xlsx      # Windows
export IBNR_GIRIS_DOSYASI=/path/claims.xlsx     # Linux / macOS
python3 ibnr.py
```

Settings — valuation period, branch, pool, triangle length, factor
window, tail, salvage sign — persist between runs in
`~/.ibnr_ayarlar.json`, and every one of them is validated before it
reaches the calculation.

### Input columns

| Column | Meaning |
|---|---|
| `HasarDonem` | accident quarter, `YYYYQn` |
| `GelDonemHT` | development period, integer quarters from the accident period |
| `HazAltBrnKod` | branch code |
| `Bedeni` | 0 = material damage, 1 = bodily injury |
| `Havuz` | pool code |
| `HasarDurum` | `Odenen` (paid) · `Muallak` (outstanding) · `SovtajRucu` (salvage/subrogation) |
| `HasarTutar` | amount |

`Odenen` rows are **incremental** paid amounts; `Muallak` is the
outstanding balance at that development point.

---

## Output

A single Excel workbook: summary, one sheet per segment with the incurred
triangle and the projected triangle, link-ratio tables with the selected
window highlighted, an assumptions sheet recording every input that went
into the run, a data-quality sheet counting what was excluded and why,
scenario comparison, backtest results, and charts.

The assumptions and data-quality sheets exist for one reason: a reserve
figure that can't be traced back to its inputs is not reviewable.

---

## Layout

```
ibnr.py                       calculation · Excel report · UI
data/uret.py                  synthetic claims generator
data/ornek_hasar.xlsx         generated sample (committed)
data/cas_indir.py             CAS database downloader + converter
tests/test_chain_ladder.py    core                    · 14 tests
tests/test_cas_donusum.py     CAS conversion          ·  7 tests
.github/workflows/tests.yml   CI: generate data, run tests
```

Code and comments are in Turkish; this README is in English.

---

## What's missing, and what comes next

This is one method producing one number. The two things that would make it
a complete reserving toolkit are known, and they're the same two things
the professional syllabus asks for:

**More than one method.** Chain ladder is least stable exactly where
alternatives do best — immature accident periods, where one large late
claim distorts the whole row.

| | |
|---|---|
| **Bornhuetter-Ferguson** | blends chain ladder with an expected loss ratio; far steadier for young periods. Needs earned premium, which the CAS dataset already carries |
| **Cape Cod** | derives that expected ratio from the data itself, so it needs no outside assumption |
| **Munich Chain Ladder** | reconciles the paid and incurred triangles instead of picking one and ignoring the other |

**A measure of uncertainty.** "IBNR is 12.4m" invites the question "how
volatile?", and a point estimate can't answer it. The **Mack (1993)**
standard error is the direct next step — distribution-free, and it builds
on the link ratios already computed here. **Bootstrap ODP** after that,
for the whole reserve distribution rather than a bound.

Smaller, honest items:

- **Single module.** Calculation, reporting and UI share one file.
- **`veri_oku` raises `SystemExit`** on a data error, which is wrong for
  library code — the caller can't catch it. A test documents the current
  behaviour rather than hiding it.

---

## License

MIT — see [LICENSE](LICENSE).
