# sot-dq-questions

Shiny app to automate the process of filling out the SoT data quality questions template

## Setup

1. Clone this repository into somewhere on the stats drive and open it in posit
2. Once opened, run the command renv::restore() in the posit console.
This sets up your version of the project with the same packages used to develop it.
This will take a couple of hours but can just be left to run.


### analysis/

**analysis.R** - Does the bulk of the analysis \
**large_changes.R** - Calculates the changes between large time bands \
**sankey_charts.R** - Generates Sankey diagrams which can be screenshotted \