* ======================================================================
* THE COMPLETE WORKFLOW: DATA CLEANING AND MASTER MERGE
* ======================================================================

* ----------------------------------------------------------------------
* STEP 1: FOLDER SETUP
* Tell Stata exactly where all your Excel files are saved.
* ----------------------------------------------------------------------
clear all
cd "D:\SAVER\Research\Data\SET"


* ----------------------------------------------------------------------
* STEP 2: CLEANING TEMPLATE FOR PANEL DATA (Repeat for all 15 files)
* ----------------------------------------------------------------------

* -----------------------------------------------------------
* 0. Total Assets (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "total asset 20FY SET&MAI filter.xlsx", sheet("Sheet1")

* Rename IDs
rename A ticker
rename B company_name
rename C country

* Dynamically rename all data columns to assets_0, assets_1, etc.
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' assets_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}

* Drop the 3 Excel header rows
drop in 1/3

* Convert to numbers (Use this for Revenue, Debt, Assets, etc.)
destring assets_*, replace force

* Handle duplicate tickers
duplicates report ticker
duplicates drop ticker, force

* Reshape to Panel Format and drop duplicated background years
reshape long assets_, i(ticker) j(lag_years)
drop if lag_years >= 20

* Save the clean panel dataset
save "assets_panel_clean.dta", replace

* -----------------------------------------------------------
* 1. ESG Grade (TEXT DATA)
* -----------------------------------------------------------
clear
import excel "RAW_ESG_GRADE_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' esg_grade_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
* Text cleaning: replace "NULL" with blank
foreach var of varlist esg_grade_* {
    replace `var' = "" if `var' == "NULL"
}
duplicates drop ticker, force
reshape long esg_grade_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "esg_grade_clean.dta", replace


* -----------------------------------------------------------
* 2. ESG Score (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_ESG_SCORE_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' esg_score_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring esg_score_*, replace force
duplicates drop ticker, force
reshape long esg_score_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "esg_score_clean.dta", replace


* -----------------------------------------------------------
* 3. Board Size (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_BOARDSIZE_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' boardsize_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring boardsize_*, replace force
duplicates drop ticker, force
reshape long boardsize_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "boardsize_clean.dta", replace


* -----------------------------------------------------------
* 4. Independent Member Percentage (NUMERIC DATA)
* NOTE: Renamed variable to _pct because Stata hates the % symbol
* -----------------------------------------------------------
clear
import excel "RAW_IND_MEMBER_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' ind_member_pct_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring ind_member_pct_*, replace force
duplicates drop ticker, force
reshape long ind_member_pct_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "ind_member%_clean.dta", replace


* -----------------------------------------------------------
* 5. Number of Meetings Per Year (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_NUM_MEETING_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' num_meeting_py_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring num_meeting_py_*, replace force
duplicates drop ticker, force
reshape long num_meeting_py_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "num_meeting_py_clean.dta", replace


* -----------------------------------------------------------
* 6. Average Attendance Meeting Percentage (NUMERIC DATA)
* NOTE: Renamed variable to _pct
* -----------------------------------------------------------
clear
import excel "RAW_AVG_ATTEND_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' avg_attent_meeting_pct_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring avg_attent_meeting_pct_*, replace force
duplicates drop ticker, force
reshape long avg_attent_meeting_pct_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "avg_attent_meeting%_clean.dta", replace


* -----------------------------------------------------------
* 7. Leverage (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_LEVERAGE_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' leverage_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring leverage_*, replace force
duplicates drop ticker, force
reshape long leverage_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "leverage_clean.dta", replace


* -----------------------------------------------------------
* 8. PBV (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_PBV_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' pbv_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring pbv_*, replace force
duplicates drop ticker, force
reshape long pbv_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "pbv_clean.dta", replace


* -----------------------------------------------------------
* 9. Market Cap Q (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_MKTCAP_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' mktcapQ_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring mktcapQ_*, replace force
duplicates drop ticker, force
reshape long mktcapQ_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "mktcapQ_clean.dta", replace


* -----------------------------------------------------------
* 10. Liabilities for Q (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_LIBFORQ_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' libforQ_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring libforQ_*, replace force
duplicates drop ticker, force
reshape long libforQ_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "libforQ_clean.dta", replace


* -----------------------------------------------------------
* 11. Assets for Q (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "RAW_ASSETSFORQ_FILE.xlsx", sheet("Sheet1")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' assetsforQ_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring assetsforQ_*, replace force
duplicates drop ticker, force
reshape long assetsforQ_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "assetsforQ_clean.dta", replace


* -----------------------------------------------------------
* 12. CEO Duality (NUMERIC DATA)
* -----------------------------------------------------------
clear
import excel "CEO Duality-copy.xlsx", sheet("CEO Duality-clean")
rename A ticker
rename B company_name
rename C country
local i = 0
local col_num = 1
foreach var of varlist _all {
    if `col_num' >= 4 {
        rename `var' dua_`i'
        local i = `i' + 1
    }
    local col_num = `col_num' + 1
}
drop in 1/3
destring dua_*, replace force
duplicates drop ticker, force
reshape long dua_, i(ticker) j(lag_years)
drop if lag_years >= 20
save "dua_clean.dta", replace

* ----------------------------------------------------------------------
* STEP 3: CLEANING THE INCORPORATION DATE (Cross-Sectional Data)
* ----------------------------------------------------------------------
clear
import excel "YOUR_AGE_FILE_NAME.xlsx", sheet("Sheet1")

rename A ticker
rename B company_name
rename C country
rename D incorp_date_str

drop in 1

* Convert text date to Stata date and extract the year
generate incorp_date = date(incorp_date_str, "DMY")
format incorp_date %td
generate incorp_year = year(incorp_date)
drop incorp_date_str

* Drop duplicates and save (NO RESHAPE NEEDED HERE)
duplicates drop ticker, force
save "incorp_clean.dta", replace


* ----------------------------------------------------------------------
* STEP 4: THE MEGA-MERGE
* Bring all the clean .dta files together into one master dataset.
* ----------------------------------------------------------------------

* Open your first main file (Master)
use "assets_panel_clean.dta", clear

* 1:1 Merges (Panel to Panel)
* (List all your other clean files here)
merge 1:1 ticker lag_years using "esg_grade_clean.dta"
drop _merge

merge 1:1 ticker lag_years using "leverage_clean.dta"
drop _merge

merge 1:1 ticker lag_years using "dua_clean.dta"
drop _merge

* m:1 Merge (Cross-Sectional to Panel)
* Merge the age data because it does not have the 'lag_years' column
merge m:1 ticker using "incorp_clean.dta"
drop _merge


* ----------------------------------------------------------------------
* STEP 5: CALCULATE VARIABLES & FINALIZE PANEL
* ----------------------------------------------------------------------

* Calculate Age (Assuming FY0 is the year 2023)
* Change 2023 to whatever actual calendar year FY0 represents for your data!
generate current_year = 2023 - lag_years
generate incorp_age = current_year - incorp_year

* Safely clear any old IDs, then create the official numeric panel ID
capture drop company_id
encode ticker, generate(company_id)

* Declare the dataset as a panel using the new ID and lag_years
xtset company_id lag_years

* Save your final masterpiece!
save "ULTIMATE_MASTER_PANEL.dta", replace

* ======================================================================
* WORKFLOW COMPLETE!
* ======================================================================
