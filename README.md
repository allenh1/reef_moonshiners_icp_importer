# `populate_icp.py`

This script is used to populate the Reef Moonshiner's excel sheet automatically
provided input from the ATI ICP test results.


## Usage

User must provide the tank size + the (public) analysis ID as arguments.

Additionally, the user must download the [Reef Moonshiner's ICP Assessment & Dosing Calculation Tool](https://347c7676-ebf9-4b0b-9a50-7f836334db30.filesusr.com/ugd/e30d85_7897efb15db94a058818bef04d45ace0.xlsx?dn=Reef%20Moonshiners%20Tools%20R2.xlsx), and rename it `reef_moonshiners.xlsx` (hopefully we can omit this step in the future).

```
usage: populate_icp.py [-h] --tank-size TANK_SIZE analysis-id

Utility to quickly read ATI test results and output an excel sheet

positional arguments:
  analysis-id           Public Analysis ID (number from the url following 'https://lab.atiaquaristik.com/publicAnalysis/')

options:
  -h, --help            show this help message and exit
  --tank-size TANK_SIZE
                        Total water volume (gallons)
```

The output will be a new exel sheet named `reef_moonshiners_[your analysis ID].xlsx`.
