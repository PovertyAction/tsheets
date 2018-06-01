*trackingsheets_v2


program tsheets
	syntax varlist [if] [in], ///
		SORTvars(varlist max=2) 	 /// list of variables to sort by
		[TITLEvars(varlist max=2)] 	 /// list of title variables
		[FILEname(str)] 			 /// name of tracking sheet file
		[Number(integer 25)]		 /// number on each page, default is 20
		[VARiable]					 /// default is to use variable labels
		[NOLabel]
	marksample touse, strok novarlist

if mi("`filename'") {
	local filename = "Tracking_Sheet"
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
				mata: adjust_column_widths("`filename'_`unit'", "`subunit'_`i'", tokens("`varlist'"), tokens("`varname_widths'"))
				
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

void adjust_column_widths(string scalar filename, string scalar sheetname, string matrix varlist, string vector varname_widths) 
{
	class xl scalar b
	real scalar i, j, k
	real vector column_widths
	string scalar title, subtitle, pgnum

	b.load_book(filename)
	b.set_sheet(sheetname)
	b.set_mode("open")
	
	column_widths = colmax(strlen(st_sdata(., varlist)))
	for (i=1; i<=cols(column_widths); i++) {
		
		if	(column_widths[i] < strtoreal(varname_widths[i])) {
			column_widths[i] = strtoreal(varname_widths[i]) 
		}
		
		b.set_column_width(i, i, column_widths[i]+2)
	}	
	
	title = st_local("title")
	subtitle = st_local("subtitle")
	pgnum = st_local("pgnum")
	j = cols(column_widths)
	k = strtoreal(st_local("endrow"))
	l = k+2
	b.set_horizontal_align(1, (1,j), "merge")
	b.set_horizontal_align(2, (1,j), "merge")
	b.put_string(1, 1, title)
	b.put_string(2, 1, subtitle)
	b.put_string(l, 1, pgnum)
	b.set_font_bold((1,4), (1,j), "on")
	b.set_border((4,k), (1,j), "thin")
	b.set_font((1, 2), (1,j), "Calibri", 16)
	b.set_horizontal_align((4, k),(1,j), "center")
	
	b.close_book()
}
	
end

