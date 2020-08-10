{smcl}
{* *! version 0.9.2 2020aug08}{...}
{cmd:help cbook_stat}
{hline}

{title:Title}

    {hi:cbook_stat} {c -} exports a codebook for a dataset.

{title:Syntax}

{p 8 17 2}
{cmd:cbook_stat} [{it:varlist}] {cmd:using} [{cmd:if}] [{cmd:in}]
[{cmd:,} {it:options}] 

{synoptset 22 tabbed}{...}
{synopthdr}
{synoptline}
{synopt:{opt rep:lace}} replaces the saved codebook. {p_end}
{synopt:{opth comp:arison(varname)}} compares whch variables are missing on an additional sheet. {p_end}
{synoptline}
{p2colreset}{...}

{title:Description}

{pstd}
{cmd:cbook_stat} takes variables of a dataset and exports summary statistics and variable information to an excel codebook. This provides
summary statistics with the variables as well as information on the labels for encoded variables. Optionally, information on if a varlist is specified, those variables only will be included in the codebook. {p_end}

{pstd} The codebook produces two to three sheets:
(1) {it:Variables}, which lists the variables in the dataset with summary information about their storage type and labeling as well as a set of summary stats;
(2) {it:Labels}, which describes the categories that encoded variables take as well as the rate which they take each level; and optionally,
(3) {it:Comparison}, which indicates which variables are non-missing for a single categorical variable. This is used primarily for describing datasets that combine data from multiple sources or survey waves.
{p_end}

{pstd}
Statistics include the following: count of non-missing values, count of missing values, mean, standard deviation, minimum, 5th percentile, 25th percentile, 50th percentile, 75th percentile, 95th percentile, and maximum. {p_end}

{marker options}
{title:Options}
{dlgtab:Options}
{marker options}{...}
{phang}
{opt rep:lace} replace the codebook file specified in {cmd:using}.{p_end}
{phang}
{opt comp:arison(varname)} produces an additional sheet that shows which variables are non-missing for the values of specified variabled. This variable will also be included in the first two sheets of the codebook, if applicable. 
{p_end}

{title:Author}

{phang}
Michael Rosenbaum, Innovations for Poverty Action{p_end}
{phang}
researchsupport@poverty-action.org
{p_end}

