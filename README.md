# Queensland crime trends

An interactive dashboard of reported offences and offenders in Queensland, per 100,000 population,
by crime type and location. Built from publicly available Queensland Police Service (QPS) and
Australian Bureau of Statistics (ABS) data sets.

Views: Trends (offence rate; adult and juvenile offender rates, by sex), Biggest changes, Map,
Funnel plot (areas against the Queensland rate), Offence mix, and About the data (sources and method).

## How this repository works

- `index.html` and `data/` are the whole site. There is no server-side code.
- `.github/workflows/refresh-and-deploy.yml` publishes the site to GitHub Pages on every push to `main`.
  On the 20th of each month it also downloads the latest QPS files, rebuilds `data/`, commits the result
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

## Running it yourself

    pip install pandas numpy openpyxl
    mkdir -p raw && cp source-files/* raw/
    python build_data.py --raw raw --out data --download-qps --report refresh_report.md
    python -m http.server 8000      # then open http://localhost:8000

Add `--tidy tidy` to also write long-format analysis files (gzipped CSV): offences by location, offence
and month for Queensland, LGAs and police divisions, and offenders by age group and sex for Queensland
and LGAs. They are too large to keep in the repository.

## Sources

- Queensland Police Service, Maps and statistics: https://www.police.qld.gov.au/maps-and-statistics
- ABS, National, state and territory population: https://www.abs.gov.au/statistics/people/population/national-state-and-territory-population/latest-release
- ABS, Regional population by age and sex: https://www.abs.gov.au/statistics/people/population/regional-population-age-and-sex/latest-release
- ABS, ASGS Edition 3 digital boundary files (LGA 2025)
- Funnel plot method: Spiegelhalter DJ (2005). Funnel plots for comparing institutional performance.
  Statistics in Medicine. https://doi.org/10.1002/sim.1970

The source data remain subject to the QPS and ABS terms of use; check those before reusing the data.
This dashboard is an independent analysis and is not a QPS or ABS product.
