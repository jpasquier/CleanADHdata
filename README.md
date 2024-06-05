# CleanADHdata.R

## Overview

**CleanADHdata.R** is an open-source R script designed to clean and enrich
electronic monitoring (EM) adherence data. This tool is specifically developed
for researchers to preprocess raw EM data collected from studies on medication
adherence, particularly focusing on oral anticancer treatments. The script aids
in preparing the data for subsequent implementation and persistence analysis.

## Features

- **Data Cleaning**: Automatically cleans raw EM adherence data to correct
  errors and inconsistencies.
- **Data Enrichment**: Enhances the dataset with additional relevant
  covariables, such as demographic and clinical data.
- **Standardization**: Ensures data is standardized for accurate adherence
  analysis.
- **Ease of Use**: Requires minimal user input with predefined formats for
  input files.

## Installation

To use the CleanADHdata.R script, you need to have R installed on your system.
You can download and install R from [CRAN](https://cran.r-project.org/).

## Usage

1. **Prepare Your Data**: Extract your EM adherence data into a CSV or Excel
   file with the following columns:
   - `PatientCode`: Alphanumeric patient identifier
   - `Monitor`: Alphanumeric monitor identifier
   - `Date`: Date-time format of the EM event
   - Additional columns such as `RecordedOpenings` for daily adherence data

2. **Gather auxiliary data**: The script requires additional data to be
   provided in Excel file to enrich the dataset. This Excel file must contain
   the following sheets:
   - `EMInfo`
   - `Regimen` The following sheets are optional:
   - `PatientCovariables`
   - `EMCovariables`
   - `AddedOpenings`
   - `NonMonitoredPeriods`
   - `AdverseEvents`

3. **Run the Script**: Execute the CleanADHdata.R script in R. The script will
   prompt for the necessary input files and parameters.

   ```R
   source("CleanADHdata.R")
   ```

4. **Review the Output**: The script generates a cleaned and enriched dataset
   ready for analysis, saved as `implementation.xlsx`.

## Example

Sample raw EM data and a tutorial are provided to help users practice using the
script. The sample raw EM can be accessed on our [GitHub
repository](https://github.com/jpasquier/CleanADHdata), whereas the tutorial is
available on [JMIR Formative
Research](https://formative.jmir.org/2024/1/e51013).

## Citation

If you use CleanADHdata.R in your research, please cite the following
publication:

Bandiera C, Pasquier J, Locatelli I, Schneider MP. Using a Semiautomated
Procedure (CleanADHdata.R Script) to Clean Electronic Adherence Monitoring
Data: Tutorial. JMIR Form Res 2024;8:e51013. doi: 10.2196/51013.
