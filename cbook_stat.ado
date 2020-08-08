*! version 0.9.1 Michael Rosenbaum - 07aug2020

program define cbook_stat

	*Syntax and version set-up
	version 15.1
	syntax [varlist] using/ [if] [in], [replace] [COMParison(varname)]

	*marksample
	marksample touse, novarlist

	cap putexcel close

	***********************************
	* Error Codes
	***********************************
	if "`replace'"=="" {
		confirm new file `"`using'"'
	}

	*If no varlist specified
	if "`varlist'" == "" {
		qui ds
		loc varlist `r(varlist)'
	}


	***********************************
	* Produce data
	***********************************
	*Produce data
	cbook_2 `varlist' `if' `in'
	mat A_c = e(sumstats)
	cbook_lab `varlist' `if' `in'
	mat A_l = e(sumstats2)

	export_cbook 	`varlist' using "`using'", replace
	export_cbooklab `varlist' using "`using'" 
	if "`comparison'" != "" export_cbookcomp `varlist' using "`using'", comparison(`comparison')

end
// end program



*******************************************************************************
* Subroutines
*******************************************************************************
*A. cbook (save summary stats)
*B. cbook_lab (save labeling file)
*C. cbook_comparison (save longitudinal variable check)
*D. get_qtile (gets quintile for cbook)
*E. export_cbook (export sheets)


**A. cbook
program cbook_2, eclass

	syntax varlist [if] [in]

	*Mar sample and names
	marksample touse, novarlist 
	tempname A i vallab

	*Save matrix
	mat `A' = J(`: word count `varlist'', 10, .)

	* Loop through varlist ot summarize
	loc `i' = 1 // init at start of matrix
	foreach var of local varlist {
		
		*only summarize unlabeled numeric vars
		cap confirm numeric variable `var'
		if _rc {
			loc ++`i'
			continue // skip string vars after describe
		}
		loc `vallab' : value label `var'
		if "``vallab''" != "" & "``vallab''" != "noyes" {
			loc ++`i'
			continue 
		}

		* Summarize variable
		qui su `var' if `touse', d
		
		mat `A'[``i'',1] = `r(mean)'
		mat `A'[``i'',2] = `r(sd)'
		mat `A'[``i'',4] = `r(min)'
		mat `A'[``i'',5] = `r(p5)'
		mat `A'[``i'',6] = `r(p25)'
		mat `A'[``i'',7] = `r(p50)'
		mat `A'[``i'',8] = `r(p75)'
		mat `A'[``i'',9] = `r(p95)'
		mat `A'[``i'',10] = `r(max)'

		qui count if mi(`var') & `touse' 
		mat `A'[``i'',3] = `r(N)'
		
		*Advance counter 
		loc ++`i'
	}
	// end foreach var of local varlist


	ereturn matrix sumstats = `A'

end


**B. cbook_lab
program cbook_lab, eclass

	syntax varlist [if] [in]

	*Mark sample and tempnames
	marksample touse, novarlist
	tempname A i lablist vallab values nlabs denom

	*Loop through labels
	loc `lablist' 																// init empty
	loc `nlabs' = 0 															// init at 0 so always has value
	foreach var of local varlist {
		loc `vallab' : value label `var'
		if "``vallab''" == "" continue 
		if "``vallab''" == "noyes" continue 									// no need to confirm yes no vars

		*Save information if labeled
		loc `lablist' ``lablist'' `var'
		qui lab list ``vallab''
		loc `nlabs' = ``nlabs'' + `r(k)' + 1 									// add in each column
	} 
	
	*Start matrix in n+1 matrices
	mat `A' = J(``nlabs'', 5, .)

	*Save values
	loc `i' = 1
	foreach var of local `lablist' {
			
		* Count missing 
		qui count if mi(`var') & `touse'
		mat `A'[``i'', 1] = `r(N)'
		loc ++`i' 																// advance past name
		
		*Set denom for valued 
		qui count if !mi(`var') & `touse'
		loc denom `r(N)'

		* Collect labels
		qui levelsof `var', loc(`values') 										// collect unique non-missing values
		foreach j of local `values' {
			
			*Create matrix
			mat `A'[``i'', 2] = `j'
			
			* Coount values 
			qui count if `var' == `j' & `touse'

			* Set matrix summaries
			mat `A'[``i'', 4] = `r(N)' 
			mat `A'[``i'', 5] = `r(N)'/`denom'

			loc ++`i' 															// advance rows
		}
		// foreach j of local `vallab'

	}
	// foreach var of local `lablist'

	ereturn matrix sumstats2 = `A'

end


**C. Export cbook
program export_cbooklab 

	syntax varlist using/ 
 
 	*Reserve names
 	tempname xlcmd i lablist vallab vlab values

	*Putexcel set 
	qui putexcel set "`using'", sheet("Labels") modify open 
	putexcel D2 = matrix(A_l), nformat("#0")

	*Place titles
	putexcel A1 = "Variable Name" 												///
			 B1 = "Variable Label" 												///
			 C1 = "Label Name" 													///
			 D1 = "Missing" 													///
			 E1 = "Value" 														/// 
			 F1 = "Value Name" 													///
			 G1 = "Count" 														///
			 H1 = "Percentage", 												///
	bold hcenter border(bottom)
 

	*Save values
	loc `lablist' // init empty
	foreach var of local varlist {
		loc `vlab' : value label `var'
		if "``vlab''" == "" continue
		if "``vlab''" == "noyes" continue 

		*Save information if labeled
		loc `lablist' ``lablist'' `var'
	} 
	// end foreach var of local varlist

	loc `xlcmd' putexcel
	loc `vlab' // init empty
	loc `i' = 2
	foreach var of local `lablist' {

		loc `vlab' 		: variable label `var'
		loc `vallab' 	: value label `var'	

		loc `xlcmd' ``xlcmd'' A``i'' = `"`var'"'
		loc `xlcmd' ``xlcmd'' B``i'' = `"``vlab''"'
		loc `xlcmd' ``xlcmd'' C``i'' = `"``vallab''"'
		loc ++`i'
	
		qui levelsof `var', loc(`values')
		foreach j of local `values' {
			loc `xlcmd' ``xlcmd'' F``i'' = `"`: label ``vallab'' `j''"'
			loc ++`i'
		}
		// end forval j = `r(min)'(1)`r(max)'

	}
	// end foreach var of local `lablist'

	putexcel H3:H``i'', nformat("#0.0%")

	*Run command
	``xlcmd''
	putexcel close

	*Formatting
	mata:	b=xl()
	mata:	b.load_book("`using'")
	mata: 	b.set_sheet("Labels")
	mata:   b.set_column_width(1, 1, 29)
	mata: 	b.set_column_width(2, 2, 60)
	mata:   b.set_column_width(3, 3, 20)
	mata:	b.set_column_width(4, 5, 10)
	mata: 	b.set_column_width(6, 6, 30)
	mata:	b.set_column_width(7, 8, 10)
	mata:	b.close_book()

	*Close putexcel environment
	putexcel clear

end 


**D. Export cbook labels
program export_cbook

	syntax varlist using/, replace 

	*Putexcel set 
	qui putexcel set "`using'", sheet("Variables") `replace' open
	putexcel E3 = matrix(A_c), nformat("#0.00")

	*Place titles
	putexcel A1 = "Variable Name" ///
			 B1 = "Variable Label" ///
			 C1 = "Type" ///
			 D1 = "Encoded" /// 
			 E1 = "Mean" ///
			 F1 = "SD" ///
			 G1 = "# Missing" ///
			 H1 = "Minimum" ///
			 I1 = "Percentiles" ///
			 N1 = "Maximum", ///
	bold hcenter
	putexcel I1:M1, merge
		
	*Place subtitles
	putexcel A2 = "" ///
			B2 = "" ///
			C2 = "" /// 
			D2 = "" ///
			E2 = "" /// 
			F2 = "" ///
			G2 = "" ///
			H2 = "" ///
			I2 = "5"  /// 
			J2 = "25" /// 
			K2 = "50" ///
			L2 = "75" ///
			M2 = "95" ///
			N2 = "", ///
	border(bottom) 

	*Save metadata
	tempname i xlcmd vlab type vallab
	loc `xlcmd' putexcel
	loc `i' = 3 // starts at 3
	foreach var of local varlist {

		* Collect names 
		loc `vlab' 		: variable label `var'
		loc `vallab' 	: value label `var'


		loc `type' // init empty
		cap confirm numeric variable `var'
		if !_rc loc `type' "integer"
		foreach vtype in string float double {
			cap confirm `vtype' variable `var' 
			if !_rc loc `type' "`vtype'"
		}

		loc `xlcmd' ``xlcmd'' A``i'' = `"`var'"'
		loc `xlcmd' ``xlcmd'' B``i'' = `"``vlab''"'
		loc `xlcmd' ``xlcmd'' C``i'' = `"``type''"'
		loc `xlcmd' ``xlcmd'' D``i'' = `"``vallab''"'

		loc ++`i'
	}

	*Format missings correctly
	putexcel G3:G``i'', nformat("0")

	*Run command
	``xlcmd''
	putexcel close

	*Formatting
	mata:	b=xl()
	mata:	b.load_book("`using'")
	mata: 	b.set_sheet("Variables")
	mata:   b.set_column_width(1, 1, 29)
	mata: 	b.set_column_width(2, 2, 40)
	mata:	b.set_column_width(3, 14, 10)
	mata:	b.close_book()

	*Close putexcel environment
	putexcel clear

end


**E. Export codebook comparison
program export_cbookcomp

	syntax varlist using/ [if] [in], comparison(varname)

	*Mark sample and name space
	marksample touse, novarlist
	tempname i j xlcmd levelslist vlab di_lvl

	*Putexcel set 
	qui putexcel set "`using'", sheet("Comparison") modify open

	qui levelsof `comparison'
	loc `levelslist' `r(levels)'

	loc `i' = 1 // init at start
	loc `j' = 3 // init at col start 
	loc `xlcmd' putexcel A1 = `"Variable"' B1 = `"Value Label"'
	forval level = 1(1)`: word count `"``levelslist''"'' {
		
		*Save lcoal
	/* " // subime parsing fix */
		// quotes don't get preserved this way so it doesn't break putexcel with an extraneous double quote on the first element of the list
		loc `xlcmd' ``xlcmd'' `: word ``j'' of `c(ALPHA)''1 = `"`: word `level' of `"``levelslist''"''"' 
	/* " // subime parsing fix */

		*Advance counters
		loc ++`i'
		loc ++`j'
	}
	// end foreach level of local `levelslist'

	``xlcmd'', hcenter txtwrap bold border(bottom)


	*Add Xs
	loc `i' = 2 // init at start
	loc `xlcmd' putexcel // init empty
	foreach var of local varlist { 

		loc `vlab' : variable label `var'
		loc `xlcmd' ``xlcmd'' A``i'' = `"`var'"'
		loc `xlcmd' ``xlcmd'' B``i'' = `"``vlab''"'

		loc `j' = 3 // init at col start
		forval level = 1(1)`: word count `"``levelslist''"'' {

	/* " // subime parsing fix */
			qui count if !mi(`var') & `touse' 									///
				& (`comparison' == `"`: word `level' of `"``levelslist''"''"')
	/* " // subime parsing fix */
			if `r(N)' != 0 {
				loc `xlcmd' ``xlcmd'' `: word ``j'' of `c(ALPHA)''``i'' = "X"
			}
			loc ++`j'
		}
		// end foreach level of local `levelslist'

		loc ++`i' 
	}
	// end foreach var f varlist

	``xlcmd'', hcenter
	putexcel A1:B``i'', left
	putexcel close

	loc --`j'

	*Formatting
	mata:	b=xl()
	mata:	b.load_book("`using'")
	mata: 	b.set_sheet("Comparison")
	mata:   b.set_column_width(1, 1, 29)
	mata: 	b.set_column_width(2, 2, 70)
	mata:	b.set_column_width(3, ``j'', 10)
	mata:	b.close_book()

end


**EOF**
