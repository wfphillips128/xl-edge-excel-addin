# Changelog

Every XL Edge release, newest first. Each version is a tagged release with the
`.xlam` attached: see [Releases](https://github.com/wfphillips128/xl-edge-excel-addin/releases).

## 2.31 — 2026-10-04 · Time series, seasonal forecasting and diagnostics

**Two new fly-outs on the Tools menu**, after Insert Monte Carlo Chart:

- **Insert Time Series ▸**, in three groups:
  - *Simulate paths* (one row per trial, one column per period): geometric
    Brownian motion, jump-diffusion (Merton), mean reversion
    (Ornstein-Uhlenbeck), square-root mean reversion (CIR), ARMA,
    ARIMA (p, 1, q) and GARCH / GJR returns.
  - *Fit to a history*: a Parameter / Estimate block for each model.
  - *Seasonal ARIMA*: fit, forecast table with a P10–P90 band, simulated
    paths, backtest, and rank orders by AIC.
- **Insert Statistical Diagnostics ▸**: Correlation Matrix, Covariance Matrix,
  Normality Tests (Shapiro-Wilk, Anderson-Darling, Jarque-Bera),
  Multicollinearity (VIF), Eigenvalues and Condition Number, Outliers and
  Leverage. Each block's size is known in advance, so the target area is checked
  before writing.

**Also new**

- **Fan Chart (Time Series)** on Insert Monte Carlo Chart: the P10–P90 band,
  the median and an optional plan line, from any block of paths.
- **Metalog** and **SPT metalog** in the distribution selector (41 in all):
  fitted to data or quantiles, or to a P10 / P50 / P90 estimate.
- **Utility and certainty equivalent** in the library: `fx.RiskUtilλ`,
  `fx.RiskUtilInvλ`, `fx.RiskCertEquivλ` (no ribbon item).
- The install item is renamed **Install or Update Monte Carlo and Statistical
  Tools Library**.

**Library v0.7.0, 136 functions**, bundled in the add-in and installed on
demand, offline. It is published as two gists: the
[Monte Carlo gist](https://gist.github.com/wfphillips128/f91bff77212ab2c3d8f55a4f0a51b8b6)
(90 functions) and the new
[Time Series & Statistics gist](https://gist.github.com/wfphillips128/96171c91caba82aa708b12b9f8f575ce)
(54). Validation: 557 checks pass, and both live gists pass a 32-check install,
save and reopen round trip.

**Breaking changes (library v0.7)**

- Models **recalculate only when their inputs change**: no distribution,
  statistics block or chart block is volatile any more.
- `fx.RiskUλ`'s `"RAND"` method is retired and returns `#VALUE!`. Pass
  `fx.RiskRandλ(Trials)` as `Trials` for live resampling.
- Statistics and chart blocks no longer read their name from the cell above the
  trials (that needed `OFFSET`, which is volatile). Without `Name` the header
  shows "Value" / "Variable 1". The ribbon already passes `Name`.

Ribbon: 152 buttons (was 127), 77 images (was 57). VBA: 23 components; the
real changes are in `modMonteCarlo`, `modMonteCarloRibbon`,
`modMonteCarloCharts` and `modAbout`.

## 2.21 — 2026-10-01 · Copulas, charts, and a bundled library

**New on the Tools menu.** The Monte Carlo items are now five, three of them
fly-outs:

- **Insert Monte Carlo Copula ▸**: Gaussian, Student t, Clayton, Gumbel and
  Frank, plus Clayton and Frank with **negative dependence** for two inputs.
  Each writes a block of correlated uniforms, one column per input, to feed into
  the distributions' `Trials` argument, so inputs move together.
- **Insert Monte Carlo Statistics ▸**: Statistics and **Detailed Statistics**
  for one result; **Variables Table** and **Variables Table (Detailed)** for
  several at once; Histogram Data; Risk Measures. (Histogram Data and Risk
  Measures moved here from the top level of the menu.)
- **Insert Monte Carlo Chart ▸**: **Histogram + S-Curve**, **Outcome
  Histogram (P10 / P50 / P90)** and **Tornado (Sensitivity)**. Each writes its
  data and draws a native Excel chart, with the title linked to a cell.

**The library ships inside the add-in.** Monte Carlo LAMBDA library **v0.6.0**
(79 functions) is bundled on a hidden sheet. Every Monte Carlo command
installs **only the functions it needs**, plus their dependencies, into the
workbook. No internet connection is needed. **Install or Update Monte Carlo
Library** now reads the bundled copy: *Yes* installs what the workbook's
formulas use, *No* installs the full library. The public gist carries the same
v0.6.0 for anyone without the add-in.

**LAMBDA Studio starts lean.** The stored library now ships **5** LAMBDAs
(`fxZScore`, `fxZOutlier`, `fxZOutlierHigh`, `fxZOutlierLow`,
`fxSlicerSelection`) instead of 61. The 56 `fx.Risk…λ` rows moved to the
bundled Monte Carlo library above.

**Breaking changes in the statistics blocks** (library v0.6). Check any
workbook built with an earlier version:

- `fx.RiskStatsλ` now starts with a header row holding the variable's name, so
  it is **12 rows** (was 11). Its default percentiles are now **P10 / P50 / P90**
  (were P5 / P50 / P95). `fx.RiskStatsDetailλ` is 33 rows.
- `fx.RiskHistλ` also gains a header row; the chart data functions start with a
  title row.
- `fx.RiskChartMarkersλ` is **removed**; its rows are now inside
  `fx.RiskChartHistλ`.
- A block that grew by a row shows **#SPILL!** if anything sits directly below
  it; move what is underneath. A formula that picks a row by position, such as
  `INDEX(E1#, 3, 2)`, needs updating.

**Source.** 23 VBA components (was 21): new `modMonteCarloCharts` and the
library sheet's module. 18 new ribbon icons.

## 2.1 — 2026-09-26 · 39 distributions and risk measures

- **Insert Monte Carlo Distribution** offers **39 distributions** (was 15),
  each with its own icon: 19 new continuous (Cauchy, Chi-squared, F, Gumbel,
  Half-Cauchy, Half-normal, Half-Student t, Inverse chi-squared, Inverse gamma,
  Inverse Gaussian, Laplace, Logistic, Noncentral beta, Noncentral F,
  Noncentral t, Pareto, Skew normal, Student t, Truncated normal) and 5 new
  discrete (Benford, Geometric, Hypergeometric, Negative Binomial, Poisson).
- New **Insert Monte Carlo Risk Measures**: VaR, CVaR and Expected Shortfall at
  a confidence level you choose.
- Optional distribution parameters can be skipped by typing `none`.
- Monte Carlo LAMBDA library v0.3.0 (56 functions).

## 2.03 — 2026-09-23 · Monte Carlo distributions

- Four new **Tools** items under Create Panel Chart: **Insert Monte Carlo
  Distribution** (15 distributions), **Insert Monte Carlo Statistics**,
  **Insert Monte Carlo Histogram Data** and **Install or Update Monte Carlo
  Library** (from the public gist).
- **Removed:** Formula Tools → *Fill Range with Random Uniform / Normal Values*,
  and their `LOWER_BOUND` / `UPPER_BOUND` settings. The distribution LAMBDAs
  replace them.
- **Shortcuts:** Ctrl-Shift-R fills right and Ctrl-Shift-D fills down;
  Ctrl-Shift-8 (toggle indents) is removed.
- 21 VBA modules, including the new `modMonteCarlo` and `modMonteCarloRibbon`.

## 2.02 — 2026-09-09 · LAMBDA Studio fixes

- The **Active File (AF)** dropdown no longer loses its selection, including on
  alt-tabbing back into Excel.
- Clicking an AF function no longer resets the filter box.
- Lower-case `=lambda(` definitions are now found.
- "[None]" can no longer be acted on as though it were a function.
- The two *Fill Range with Random …* screentips now read "Between X to Y".

## 2.01 — 2026-08-22 · First public release

- **LAMBDA Studio**: a searchable, shareable library of custom LAMBDA functions
  in the ribbon, with import and export as an Excel workbook or Advanced Formula
  Environment gist text.
- Productivity macros for formatting, formulas and sheet management.
- *Updated 2026-08-30:* **Tools → Create Panel Chart** builds a grid of small
  charts on one shared scale as a single native chart. The asset was refreshed
  in place; the version string stayed at 2.01.
