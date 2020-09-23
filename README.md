# cbook_stat

## Overview
cbook_stat is a Stata package that creates a codebook with up to three features:

1. A descriptive list of variables with metadata about variable information and summary statistics. 
2. A list of encoded categorical variables with response rates.
3. A comparison of variable missingness by any variable in the dataset.

## Installation

```Stata
* ipacheck may be installed directly from GitHub
net install cbook_stat, from("https://raw.githubusercontent.com/PovertyAction/cbook_stat/master") replace 
```
