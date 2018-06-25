*! version 1.0.0 Rosemarie Sandino 07jun2018

program tsheets
	syntax varlist [if] [in], ///
		SORTvars(varlist max=2) 	 /// list of variables to sort by
		[TITLEvars(varlist max=2)] 	 /// list of title variables
		[FILEname(str)] 			 /// name of tracking sheet file
		[Number(integer 25)]		 /// number on each page, default is 25
		[VARiable]					 /// default is to use variable labels
		[NOLabel]
	marksample touse, strok novarlist

version 13	
	
if mi("`filename'") {
	local filename = "Tracking_Sheet"
}

if regexm("`filename'", ".xls") {
	local filename = substr("`filename'", 1, strpos("`filename'", ".xl")-1) 
}

	
if mi("`variable'") {
	local variable = "varl"
}

qui {
	preserve
**************************** Create Tracking Sheets ****************************

	keep if `touse'

	sort `sortvars' `varlist'
	order `varlist'
	
	*If want value labels, encode variable so those are used as colwidth
	if "`nolabel'" == "" {
		ds `varlist', has(vallab)
		foreach var in `r(varlist)' {
			decode `var', gen(`var'_new)
			drop `var'
			ren `var'_new `var'
		}
	}

	ds `varlist' `sortvars' `titlevars', has(type string)
	foreach var in `r(varlist)' {
		replace `var' = stritrim(strtrim(strproper(subinstr(`var', ".", "", .))))
	}

	* Everything must be a string to maintain formatting
	ds `sortvars' `titlevars', has(type numeric)
	foreach var in `r(varlist)' {
		tostring `var', replace force
	}
	
	if mi("`titlevars'") {
		local titlevars `sortvars'
	}

	* If they want variable labels instead of variables as titles
	foreach var in `varlist' {	
		if "`variable'" == "varl" {
			local lab : variable label `var'
			local len = strlen("`lab'")
		}
		if "`variable'" == "variable" | "`len'" == "0" {
			local len = strlen("`var'")
		}
		local varname_widths `varname_widths' `len' 
	}
	
	gettoken unitvar subunitvar : sortvars
	gettoken titlevar subtitlevar : titlevars

	if mi("`subunitvar'") {
		local subunitvar = "`unitvar'"
	}
	
	levelsof `unitvar', local(units)

	foreach unit in `units' {

		levelsof `subunitvar' if `unitvar' == "`unit'", local(subunits)
		noi dis "Generating tracking sheet for `unitvar' `unit'..." _continue
		foreach subunit in `subunits' {
		
			* Calculate the number of pages needed
			local pages = ceil(`r(N)'/`number')
			
			* Place holders
			local start = 1
			local end = `start' + `number'
		
			forval i = 1/`pages' {
			
				levelsof `titlevar' if `unitvar' == "`unit'", local(t)
				local title : word 1 of `t'
				
				if !mi("`subtitlevar'") {
					levelsof `subtitlevar' if `subunitvar' == "`subunit'", local(sub)
					local subtitle : word 1 of `sub'
				}
				
				local pgnum = "`i' of `pages'"
				tempvar n
				bysort `subunitvar' : gen `n' = _n
				count if `n' >= `start' & `n' < `end' & `subunitvar' == "`subunit'"
				local test = `r(N)'
				if `test' < `number' {
					local endcount = `test'
				}
				else local endcount = `number'
				local endrow = 4 + `endcount'
				
				if `i'==1 {
					cap confirm file "`filename'_`unit'.xlsx"
					if _rc==0 {
						erase "`filename'_`unit'.xlsx"
					}
				}
				
				* Export
				export excel `varlist' ///
					if `unitvar' == "`unit'" & `subunitvar' == "`subunit'" & ///
					`n' >= `start'  & `n' < `end' ///
					using "`filename'_`unit'.xlsx", ///
					sheet("`subunit'_`i'") ///
					sheetmodify ///
					firstrow(`variable')  ///
					cell(A4) miss("") `nolabel'
				
				*Add formatting
				mata: adjust_column_widths("`filename'_`unit'", "`subunit'_`i'", tokens("`varlist'"))
				
				local start = `start' + `number'
				local end = `end' + `number'	
			}
		}
		noi dis "Done!" _col(55)
	}
	restore
}
end 


mata:
mata clear

void adjust_column_widths(string scalar filename, string scalar sheetname, string matrix varlist) 
{
	class xl scalar b
	real scalar i
	real vector column_widths, varname_widths
	string scalar varlabel

	b.load_book(filename)
	b.set_sheet(sheetname)
	b.set_mode("open")
	
	varname_widths = strlen(varlist)
	column_widths = colmax(strlen(st_sdata(., varlist)))
	
	for (i=1; i<=cols(column_widths); i++) {
		if (st_local("variable") == "varl") {
			varlabel = st_varlabel(varlist[i])
			if (varname_widths[i] < strlen(varlabel)) {
				varname_widths[i] = strlen(varlabel)
			}
		}	
		if	(column_widths[i] < varname_widths[i]) {
			column_widths[i] = varname_widths[i] 
		}
		b.set_column_width(i, i, column_widths[i]+2)
	}	
	
	b.set_horizontal_align(1, (1, cols(column_widths)), "merge")
	b.set_horizontal_align(2, (1, cols(column_widths)), "merge")
	b.put_string(1, 1, st_local("title"))
	b.put_string(2, 1, st_local("subtitle"))
	b.put_string(strtoreal(st_local("endrow"))+2, 1, st_local("pgnum"))
	b.set_font_bold((1,4), (1, cols(column_widths)), "on")
	b.set_border((4, strtoreal(st_local("endrow"))), (1, cols(column_widths)), "thin")
	b.set_font((1, 2), (1, cols(column_widths)), "Calibri", 16)
	b.set_horizontal_align((4, strtoreal(st_local("endrow"))), (1, cols(column_widths)), "center")
	
	b.close_book()
}
	
end

