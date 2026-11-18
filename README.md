# Queensland crime trends <img src="assets/readme-hex.svg" align="right" width="150" alt="Queensland crime trends hex logo" />

An interactive dashboard of reported offences, offenders and victims in Queensland, per 100,000
population, by crime type and location, with area disadvantage (ABS SEIFA) alongside. Built from publicly available Queensland Police Service (QPS) and
Australian Bureau of Statistics (ABS) data sets.

Views: Trends (offence rate; adult and juvenile offender and victim rates, by sex), Biggest changes, Map
(rates or SEIFA disadvantage), Funnel plot (areas against the Queensland rate), Offence mix, Disadvantage
(rates against SEIFA), Victims (QPS monthly by place; ABS annual detail), and About the data (sources and method).

## How this repository works

- `index.html` and `data/` are the whole site. There is no server-side code.
- `.github/workflows/refresh-and-deploy.yml` publishes the site to GitHub Pages on every push to `main`.
  On the 20th of each month it also downloads the latest QPS files (eight offence and offender CSVs and two
  victim workbooks), rebuilds `data/`, commits the result
  and publishes. It can be run by hand from the Actions tab (Run workflow).
- `build_data.py` turns the source files into `data/`. `source-files/` holds the inputs that change rarely:
  two ABS population workbooks and the two boundary files.
- `refresh_report.md` summarises the latest month after each refresh.

One-time setup after the first push: Settings > Pages > Build and deployment > Source: GitHub Actions.

Each monthly refresh adds a commit of the rebuilt data (a few megabytes compressed), so the repository
grows slowly over time.

## Updating the ABS files (about once a year)

When ABS releases new population estimates, replace the two workbooks in `source-files/`:

- `3101053.xlsx`: National, state and territory population, Table 53 (ERP by single year of age, Queensland)
- `32350DS0006_*.xlsx`: Regional population by age and sex (ERP by age and sex, LGAs, 2001 onwards)

Until then, offender-rate denominators for months after the latest 30 June estimate are extrapolated.

## The Victims tab (about once a year)

The Victims tab reads `data/victims.json`, which is built from the ABS release Recorded Crime – Victims.
It is not part of the monthly refresh, because ABS publishes it once a year and does not offer it through an API.

When ABS publishes a new release (https://www.abs.gov.au/statistics/people/crime-and-justice/recorded-crime-victims/latest-release):

1. Move `victims-source/current/2. Victims of crime, states and territories (Tables 9 to 16).xlsx`
   to `victims-source/earlier/victims_<its year>.xlsx` (for example `victims_2025.xlsx`). The age series from 2010 is assembled from one such file per year.
2. Download the new data cubes 2, 3 and 5 into `victims-source/current/`, keeping the ABS file names.
3. Make sure `source-files/3101053.xlsx` (Queensland population by age and sex) covers 30 June of the new year.
4. Run `python build_victims.py --victims victims-source --erp source-files/3101053.xlsx --out data/victims.json`
5. Commit and push. The build prints how closely its age-specific rates match those ABS publishes; it should be within about half a per cent.

## Area disadvantage (SEIFA, every five years)

`data/seifa.json` holds ABS SEIFA 2021 (all four indexes) for LGAs and police divisions. It is built by
`build_seifa.py` and is not part of the monthly refresh. Police division scores are population-weighted
averages of SA1 scores; the build checks the method by comparing LGA averages with the published LGA scores.
Its inputs are kept outside the repository because the SA1 boundary file is larger than GitHub allows:

    pip install pandas openpyxl shapely pyshp
    python build_seifa.py --raw <folder with the inputs> --meta data/meta.json --out data/seifa.json

Inputs: `SEIFA2021_LGA_Indexes.xlsx` and `SEIFA2021_SA1_Indexes.xlsx` (ABS SEIFA 2021 data cubes),
`SA1_2021_AUST_SHP_GDA2020.zip` (ABS SA1 boundaries), `LGA_2025_AUST_GDA2020.zip` and `QPS_DIVISIONS.zip`.
Rebuild when ABS publishes SEIFA from the 2026 Census, or if QPS division boundaries change.

## Running it yourself

    pip install pandas numpy openpyxl
    mkdir -p raw && cp source-files/* raw/
    python build_data.py --raw raw --out data --download-qps --report refresh_report.md
    python -m http.server 8000      # then open http://localhost:8000

Add `--tidy tidy` to also write long-format analysis files (gzipped CSV): offences by location, offence
and month for Queensland, LGAs and police divisions, offenders by age group and sex for Queensland
and LGAs, and victims (all victims, and by age group and sex for offences against the person) for
Queensland and LGAs. They are too large to keep in the repository.

## Sources

- Queensland Police Service, Maps and statistics: https://www.police.qld.gov.au/maps-and-statistics
- ABS, National, state and territory population: https://www.abs.gov.au/statistics/people/population/national-state-and-territory-population/latest-release
- ABS, Regional population by age and sex: https://www.abs.gov.au/statistics/people/population/regional-population-age-and-sex/latest-release
- ABS, ASGS Edition 3 digital boundary files (LGA 2025, SA1 2021)
- ABS, Socio-Economic Indexes for Areas (SEIFA), 2021: https://www.abs.gov.au/statistics/people/people-and-communities/socio-economic-indexes-areas-seifa-australia/latest-release
- ABS, Recorded Crime – Victims: https://www.abs.gov.au/statistics/people/crime-and-justice/recorded-crime-victims/latest-release
- Spearman confidence intervals: Bonett DG, Wright TA (2000). Sample size requirements for estimating
  Pearson, Kendall and Spearman correlations. Psychometrika. https://doi.org/10.1007/BF02294183
- Funnel plot method: Spiegelhalter DJ (2005). Funnel plots for comparing institutional performance.
  Statistics in Medicine. https://doi.org/10.1002/sim.1970

The source data remain subject to the QPS and ABS terms of use; check those before reusing the data.
This dashboard is an independent analysis and is not a QPS or ABS product.
