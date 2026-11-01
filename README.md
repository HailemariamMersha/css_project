# Generative AI Exposure and Full-Time → Involuntary Part-Time Transitions

Computational Social Science (CS-UH 2219E), Fall 2026.

This folder reproduces the CPS data-processing and preliminary descriptive results used for Milestone 2. The AIOE merge will be completed once the Language-Model AIOE workbook and the final occupation crosswalk are added.

## Research question

> Did workers in U.S. occupations with greater exposure to language-model capabilities become more likely to transition from full-time wage employment to part-time work for economic reasons after 2022, relative to workers in lower-exposure occupations?

## Data

Place the cleaned IPUMS Basic Monthly CPS extract at:

```text
data/cps_00003.dat.gz
```

The raw IPUMS file should **not** be committed to GitHub. It is large and is governed by the IPUMS data-use terms. Commit the scripts, documentation, and derived aggregate tables instead.

The extract used here contains 2017-2025 Basic Monthly CPS records and has fixed record width 137 characters. The parser in `01_build_cps_transitions.py` validates that width before proceeding.

## Baseline analytical sample

A record is eligible as a baseline when:

- age is 18-64;
- `EMPSTAT == 10` (at work);
- `CLASSWKR` is one of 22, 23, 25, 27, 28 (private for-profit/nonprofit or federal/state/local wage/salary employment);
- `WKSTAT == 11` (35+ hours and usually full-time);
- `MISH` is 1-4;
- `LNKFW1YWT > 0`;
- baseline year is 2017-2024.

The same respondent is matched in the same calendar month one year later using the validated longitudinal identifier `CPSIDV`.

## Outcomes

Primary:

```text
pter_transition = 1 if follow-up WKSTAT ∈ {21, 22}
```

Complementary:

```text
unemp_transition = 1 if follow-up EMPSTAT ∈ {21, 22}
pt_nonecon_transition = 1 if follow-up WKSTAT ∈ {12, 41}
ft_followup = 1 if follow-up WKSTAT == 11
```

## Reproduce the transition file

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python scripts/01_build_cps_transitions.py \
  --input data/cps_00003.dat.gz \
  --output outputs/cps_one_year_transitions.csv.gz \
  --sample-summary outputs/sample_construction.csv
```

Expected headline counts for the current extract:

| Stage | N |
|---|---:|
| Basic Monthly CPS person-records, 2017-2025 | 11,635,165 |
| Age 18-64 | 6,849,416 |
| At-work wage/salary records | 4,344,203 |
| Full-time wage/salary records | 3,456,082 |
| Eligible one-year baselines, 2017-2024 | 1,112,384 |
| Validated one-year transitions | 1,062,438 |
| FT → PTER transitions | 13,103 |

## Reproduce descriptive results

```bash
python scripts/02_summarize_cps_transitions.py \
  --input outputs/cps_one_year_transitions.csv.gz \
  --output-dir outputs \
  --figure-dir figures
```

This creates:

- `transition_summary_by_year.csv`
- `overall_followup_outcomes.csv`
- `pter_reasons.csv`
- `pter_by_age_band.csv`
- `period_summary.csv`
- `figures/pter_rate_by_year.png`
- `figures/pter_by_age_band.png`

## AIOE merge

`03_merge_aioe_template.py` is a placeholder for the next stage. Once the Language-Model AIOE workbook is available, we will inspect its exact SOC coding and build/validate the occupation crosswalk before implementing the merge. We will report both occupation-level and worker-level match coverage rather than silently dropping unmatched occupations.

## Important interpretation note

The CPS results in Milestone 2 are **descriptive validation of the outcome and panel construction**. They do not test the AI hypothesis until occupation-level LM-AIOE exposure has been merged. In particular, the aggregate post-2022 PTER rate can remain stable even if high-exposure occupations diverge from low-exposure occupations.
