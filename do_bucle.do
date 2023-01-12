clear all
set more off

cd "ruta"

set obs 0
gen XXX=""
save "BD_FT_EXCEL.dta", replace

local files: dir . files "*.xlsx"
foreach file in `files'{
	clear
	import excel using "`file'", describe
	forvalues x=2/`r(N_worksheet)' {
		local nombrehoja = "`" + "r(worksheet_" + "`x'" + ")" + "'"
		di "`nombrehoja'"
		import excel using "`file'", clear sheet(`nombrehoja') allstring
		append using "BD_FT_EXCEL.dta"
		save "BD_FT_EXCEL.dta", replace
	}
}

use "BD_FT_EXCEL.dta", clear
drop XXX
save "BD_FT_EXCEL.dta", replace
