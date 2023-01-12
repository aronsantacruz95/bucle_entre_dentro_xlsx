clear all
set more off

cd "ruta"

cap erase "BD_FT_EXCEL.xlsx"
local files: dir . files "*.xlsx"
set obs 0
gen XXX=""
save "BD_FT_EXCEL.dta", replace
foreach file in `files'{
	clear
	import excel using "`file'", describe
	forvalues x=1/`r(N_worksheet)' {
		local nombrehoja = "`" + "r(worksheet_" + "`x'" + ")" + "'"
		di "`nombrehoja'"
		import excel using "`file'", clear sheet(`nombrehoja') allstring
		append using "BD_FT_EXCEL.dta"
		save "BD_FT_EXCEL.dta", replace
	}
}

use "BD_FT_EXCEL.dta", clear
drop XXX
drop G-Z
save "BD_FT_EXCEL.dta", replace
export excel using "BD_FT_EXCEL.xlsx", replace sheet("BD")
