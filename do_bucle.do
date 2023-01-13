clear all
set more off

cd "ruta"

cap erase "BD_FT_EXCEL.xlsx"
sleep 2
local files: dir . files "*.xlsx"
set obs 0
gen XXX=""
save "BD_FT_EXCEL.dta", replace
foreach file in `files'{
	clear
	quietly: import excel using "`file'", describe
	forvalues x=1/`r(N_worksheet)' {
		quietly: import excel using "`file'", describe
		local _nombrehoja=r(worksheet_`x')
		di "`file'"
		di "`_nombrehoja'"
		di "---"
 		import excel using "`file'", clear sheet(`_nombrehoja') allstring
		quietly: import excel using "`file'", describe
		gen name_sheet="`_nombrehoja'"
		gen name_file="`file'"
		append using "BD_FT_EXCEL.dta"
		save "BD_FT_EXCEL.dta", replace
	}
}

use "BD_FT_EXCEL.dta", clear
cap drop XXX
cap drop G-Z
cap drop AA-AH
drop if (A=="" & B=="" & C=="" & D=="" & E=="" & F=="")
save "BD_FT_EXCEL.dta", replace
export excel using "BD_FT_EXCEL.xlsx", replace sheet("BD")
