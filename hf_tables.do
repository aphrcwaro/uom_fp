
clear all
set maxvar 20000
set matsize 10000
set more off,permanently

cap cd "D:\Dropbox\Dropbox\FP-Project-APHRC-UM\Worshop_FP_UM\Health_facility_tables"
cap cd "F:\Dropbox_APHRC\Dropbox\FP-Project-APHRC-UM\Worshop_FP_UM\Health_facility_tables"

use health_facility_s4_dropped.dta, clear

*8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
*///////////////////////////////////////////////////////////////////////////////
***Table 5.1.1.1.  % facilities with available and functioning infrastructure items
*///////////////////////////////////////////////////////////////////////////////

putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") replace
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"

putexcel A2 = "Table 5.1.1.1. % facilities with available and functioning infrastructure items", bold

putexcel B3 = "A waiting room with seating", bold
putexcel C3 = "Men's toilet with running water in the waiting room", bold
putexcel D3 = "Women's toilet with running water in the waiting room", bold
putexcel E3 = "Handwashing device", bold
putexcel F3 = "Drinking water", bold
putexcel G3 = "Power supply", bold
putexcel H3 = "Laboratory", bold
putexcel I3 = "Imaging Service", bold
putexcel J3 = "Operating room", bold
putexcel K3 = "Pharmacy", bold
putexcel L3 = "Orientation Signs", bold
putexcel M3 = "Ramps for people with disabilities", bold
putexcel N3 = "Biomedical waste collection room", bold
putexcel O3 = "Car parking", bold
putexcel P3 = "Number of facilities ", bold






*///////////////////////////////////////////////////////////////////////////////
*** A waiting room with seating
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1a, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  total_responses
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel P5 = total_responses  , nformat(number_d0)
putexcel B5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1a s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
 putexcel P7 = total_dakar, nformat(number_d0)
 putexcel P8= total_diourbel, nformat(number_d0) 
 putexcel P9= total_fatick , nformat(number_d0) 
 putexcel P10= total_kaffrine , nformat(number_d0) 
 putexcel P11= total_kaolack , nformat(number_d0) 
 putexcel P12= total_kedougou , nformat(number_d0) 
 putexcel P13= total_kolda , nformat(number_d0) 
 putexcel P14= total_louga , nformat(number_d0) 
 putexcel P15= total_matam , nformat(number_d0) 
 putexcel P16= total_stlouis , nformat(number_d0) 
 putexcel P17= total_sedhiou , nformat(number_d0) 
 putexcel P18= total_tambacounda , nformat(number_d0) 
 putexcel P19= total_thies , nformat(number_d0) 
 putexcel P20= total_ziguinchor , nformat(number_d0) 


// Percentages "Oui"
 putexcel B7= pct_dakar, nformat(number_d2)
 putexcel B8= pct_diourbel, nformat(number_d2)
 putexcel B9= pct_fatick, nformat(number_d2)
 putexcel B10= pct_kaffrine, nformat(number_d2)
 putexcel B11= pct_kaolack, nformat(number_d2)
 putexcel B12= pct_kedougou, nformat(number_d2)
 putexcel B13= pct_kolda, nformat(number_d2)
 putexcel B14= pct_louga, nformat(number_d2)
 putexcel B15= pct_matam, nformat(number_d2)
 putexcel B16= pct_stlouis, nformat(number_d2)
 putexcel B17= pct_sedhiou, nformat(number_d2)
 putexcel B18= pct_tambacounda, nformat(number_d2)
 putexcel B19= pct_thies, nformat(number_d2)
 putexcel B20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1a s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel B22 = pct_eps, nformat(number_d2)
putexcel B23 = pct_cs , nformat(number_d2)
putexcel B24 = pct_ps , nformat(number_d2)
putexcel B25 = pct_case , nformat(number_d2)
putexcel B26 = pct_clinic , nformat(number_d2)


putexcel P22 = total_eps , nformat(number_d0)
putexcel P23 = total_cs , nformat(number_d0)
putexcel P24 = total_ps , nformat(number_d0)
putexcel P25 = total_case , nformat(number_d0)
putexcel P26 = total_clinic, nformat(number_d0)
putexcel close




*///////////////////////////////////////////////////////////////////////////////
*** Men's toilet with running water in the waiting room
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1b, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel C33 = count_oui  , nformat(number_d0)
putexcel C5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1b s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel C7= pct_dakar, nformat(number_d2)
 putexcel C8= pct_diourbel, nformat(number_d2)
 putexcel C9= pct_fatick, nformat(number_d2)
 putexcel C10= pct_kaffrine, nformat(number_d2)
 putexcel C11= pct_kaolack, nformat(number_d2)
 putexcel C12= pct_kedougou, nformat(number_d2)
 putexcel C13= pct_kolda, nformat(number_d2)
 putexcel C14= pct_louga, nformat(number_d2)
 putexcel C15= pct_matam, nformat(number_d2)
 putexcel C16= pct_stlouis, nformat(number_d2)
 putexcel C17= pct_sedhiou, nformat(number_d2)
 putexcel C18= pct_tambacounda, nformat(number_d2)
 putexcel C19= pct_thies, nformat(number_d2)
 putexcel C20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1b s1_7, matcell(freqs) matrow(names)
scalar total_eps   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_ps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_clinic= freqs[1,5] + freqs[2,5]


scalar pct_eps    = (freqs[1,1] / total_eps) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_ps     = (freqs[1,3] / total_ps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_clinic = (freqs[1,5] / total_clinic) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel C22 = pct_eps, nformat(number_d2)
putexcel C23 = pct_cs , nformat(number_d2)
putexcel C24 = pct_ps , nformat(number_d2)
putexcel C25 = pct_case , nformat(number_d2)
putexcel C26 = pct_clinic , nformat(number_d2)
putexcel close





*///////////////////////////////////////////////////////////////////////////////
*** Women's toilet with running water in the waiting room
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1c, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel D33 = count_oui  , nformat(number_d0)
putexcel D5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1c s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel D7= pct_dakar, nformat(number_d2)
 putexcel D8= pct_diourbel, nformat(number_d2)
 putexcel D9= pct_fatick, nformat(number_d2)
 putexcel D10= pct_kaffrine, nformat(number_d2)
 putexcel D11= pct_kaolack, nformat(number_d2)
 putexcel D12= pct_kedougou, nformat(number_d2)
 putexcel D13= pct_kolda, nformat(number_d2)
 putexcel D14= pct_louga, nformat(number_d2)
 putexcel D15= pct_matam, nformat(number_d2)
 putexcel D16= pct_stlouis, nformat(number_d2)
 putexcel D17= pct_sedhiou, nformat(number_d2)
 putexcel D18= pct_tambacounda, nformat(number_d2)
 putexcel D19= pct_thies, nformat(number_d2)
 putexcel D20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1c s1_7, matcell(freqs) matrow(names)
scalar total_eps   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_ps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_clinic= freqs[1,5] + freqs[2,5]


scalar pct_eps    = (freqs[1,1] / total_eps) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_ps     = (freqs[1,3] / total_ps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_clinic = (freqs[1,5] / total_clinic) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel D22 = pct_eps, nformat(number_d2)
putexcel D23 = pct_cs , nformat(number_d2)
putexcel D24 = pct_ps , nformat(number_d2)
putexcel D25 = pct_case , nformat(number_d2)
putexcel D26 = pct_clinic , nformat(number_d2)
putexcel close





*///////////////////////////////////////////////////////////////////////////////
*** Handwashing device
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1d, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel E33 = count_oui  , nformat(number_d0)
putexcel E5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1d s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel E7= pct_dakar, nformat(number_d2)
 putexcel E8= pct_diourbel, nformat(number_d2)
 putexcel E9= pct_fatick, nformat(number_d2)
 putexcel E10= pct_kaffrine, nformat(number_d2)
 putexcel E11= pct_kaolack, nformat(number_d2)
 putexcel E12= pct_kedougou, nformat(number_d2)
 putexcel E13= pct_kolda, nformat(number_d2)
 putexcel E14= pct_louga, nformat(number_d2)
 putexcel E15= pct_matam, nformat(number_d2)
 putexcel E16= pct_stlouis, nformat(number_d2)
 putexcel E17= pct_sedhiou, nformat(number_d2)
 putexcel E18= pct_tambacounda, nformat(number_d2)
 putexcel E19= pct_thies, nformat(number_d2)
 putexcel E20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1d s1_7, matcell(freqs) matrow(names)
scalar total_eps   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_ps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_clinic= freqs[1,5] + freqs[2,5]


scalar pct_eps    = (freqs[1,1] / total_eps) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_ps     = (freqs[1,3] / total_ps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_clinic = (freqs[1,5] / total_clinic) * 100


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel E22 = pct_eps, nformat(number_d2)
putexcel E23 = pct_cs , nformat(number_d2)
putexcel E24 = pct_ps , nformat(number_d2)
putexcel E25 = pct_case , nformat(number_d2)
putexcel E26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
***Drinking water
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1e, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel F33 = count_oui  , nformat(number_d0)
putexcel F5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1e s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel F7= pct_dakar, nformat(number_d2)
 putexcel F8= pct_diourbel, nformat(number_d2)
 putexcel F9= pct_fatick, nformat(number_d2)
 putexcel F10= pct_kaffrine, nformat(number_d2)
 putexcel F11= pct_kaolack, nformat(number_d2)
 putexcel F12= pct_kedougou, nformat(number_d2)
 putexcel F13= pct_kolda, nformat(number_d2)
 putexcel F14= pct_louga, nformat(number_d2)
 putexcel F15= pct_matam, nformat(number_d2)
 putexcel F16= pct_stlouis, nformat(number_d2)
 putexcel F17= pct_sedhiou, nformat(number_d2)
 putexcel F18= pct_tambacounda, nformat(number_d2)
 putexcel F19= pct_thies, nformat(number_d2)
 putexcel F20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1e s1_7, matcell(freqs) matrow(names)
scalar total_eps   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_ps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_clinic= freqs[1,5] + freqs[2,5]


scalar pct_eps    = (freqs[1,1] / total_eps) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_ps     = (freqs[1,3] / total_ps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_clinic = (freqs[1,5] / total_clinic) * 100


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel F22 = pct_eps, nformat(number_d2)
putexcel F23 = pct_cs , nformat(number_d2)
putexcel F24 = pct_ps , nformat(number_d2)
putexcel F25 = pct_case , nformat(number_d2)
putexcel F26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
*** Power supply
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1f, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel G33 = count_oui  , nformat(number_d0)
putexcel G5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1f s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel G7= pct_dakar, nformat(number_d2)
 putexcel G8= pct_diourbel, nformat(number_d2)
 putexcel G9= pct_fatick, nformat(number_d2)
 putexcel G10= pct_kaffrine, nformat(number_d2)
 putexcel G11= pct_kaolack, nformat(number_d2)
 putexcel G12= pct_kedougou, nformat(number_d2)
 putexcel G13= pct_kolda, nformat(number_d2)
 putexcel G14= pct_louga, nformat(number_d2)
 putexcel G15= pct_matam, nformat(number_d2)
 putexcel G16= pct_stlouis, nformat(number_d2)
 putexcel G17= pct_sedhiou, nformat(number_d2)
 putexcel G18= pct_tambacounda, nformat(number_d2)
 putexcel G19= pct_thies, nformat(number_d2)
 putexcel G20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1f s1_7, matcell(freqs) matrow(names)
scalar total_case = freqs[1,1] + freqs[2,1]
scalar total_clinic = freqs[1,2] + freqs[2,2]
scalar total_cs = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]



scalar pct_case = (freqs[1,1] / total_case) * 100
scalar pct_clinic = (freqs[1,2] / total_clinic) * 100
scalar pct_cs = (freqs[1,3] / total_cs) * 100
scalar pct_eps = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100

  


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel G22 = pct_eps, nformat(number_d2)
putexcel G23 = pct_cs , nformat(number_d2)
putexcel G24 = pct_ps , nformat(number_d2)
putexcel G25 = pct_case , nformat(number_d2)
putexcel G26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
*** Laboratory
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1g, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel H33 = count_oui  , nformat(number_d0)
putexcel H5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1g s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel H7= pct_dakar, nformat(number_d2)
 putexcel H8= pct_diourbel, nformat(number_d2)
 putexcel H9= pct_fatick, nformat(number_d2)
 putexcel H10= pct_kaffrine, nformat(number_d2)
 putexcel H11= pct_kaolack, nformat(number_d2)
 putexcel H12= pct_kedougou, nformat(number_d2)
 putexcel H13= pct_kolda, nformat(number_d2)
 putexcel H14= pct_louga, nformat(number_d2)
 putexcel H15= pct_matam, nformat(number_d2)
 putexcel H16= pct_stlouis, nformat(number_d2)
 putexcel H17= pct_sedhiou, nformat(number_d2)
 putexcel H18= pct_tambacounda, nformat(number_d2)
 putexcel H19= pct_thies, nformat(number_d2)
 putexcel H20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1g s1_7, matcell(freqs) matrow(names)
scalar total_clinic   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_eps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_clinic    = (freqs[1,1] / total_clinic) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_eps     = (freqs[1,3] / total_eps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel H22 = pct_eps, nformat(number_d2)
putexcel H23 = pct_cs , nformat(number_d2)
putexcel H24 = pct_ps , nformat(number_d2)
putexcel H25 = pct_case , nformat(number_d2)
putexcel H26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
*** Imaging Service
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1h, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel I33 = count_oui  , nformat(number_d0)
putexcel I5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1h s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel I7= pct_dakar, nformat(number_d2)
 putexcel I8= pct_diourbel, nformat(number_d2)
 putexcel I9= pct_fatick, nformat(number_d2)
 putexcel I10= pct_kaffrine, nformat(number_d2)
 putexcel I11= pct_kaolack, nformat(number_d2)
 putexcel I12= pct_kedougou, nformat(number_d2)
 putexcel I13= pct_kolda, nformat(number_d2)
 putexcel I14= pct_louga, nformat(number_d2)
 putexcel I15= pct_matam, nformat(number_d2)
 putexcel I16= pct_stlouis, nformat(number_d2)
 putexcel I17= pct_sedhiou, nformat(number_d2)
 putexcel I18= pct_tambacounda, nformat(number_d2)
 putexcel I19= pct_thies, nformat(number_d2)
 putexcel I20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1h s1_7, matcell(freqs) matrow(names)
scalar total_clinic   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_eps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_clinic    = (freqs[1,1] / total_clinic) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_eps     = (freqs[1,3] / total_eps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel I22 = pct_eps, nformat(number_d2)
putexcel I23 = pct_cs , nformat(number_d2)
putexcel I24 = pct_ps , nformat(number_d2)
putexcel I25 = pct_case , nformat(number_d2)
putexcel I26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
*** Operating room
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1i, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel J33 = count_oui  , nformat(number_d0)
putexcel J5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1i s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel J7= pct_dakar, nformat(number_d2)
 putexcel J8= pct_diourbel, nformat(number_d2)
 putexcel J9= pct_fatick, nformat(number_d2)
 putexcel J10= pct_kaffrine, nformat(number_d2)
 putexcel J11= pct_kaolack, nformat(number_d2)
 putexcel J12= pct_kedougou, nformat(number_d2)
 putexcel J13= pct_kolda, nformat(number_d2)
 putexcel J14= pct_louga, nformat(number_d2)
 putexcel J15= pct_matam, nformat(number_d2)
 putexcel J16= pct_stlouis, nformat(number_d2)
 putexcel J17= pct_sedhiou, nformat(number_d2)
 putexcel J18= pct_tambacounda, nformat(number_d2)
 putexcel J19= pct_thies, nformat(number_d2)
 putexcel J20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1i s1_7, matcell(freqs) matrow(names)
scalar total_clinic   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_eps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_clinic    = (freqs[1,1] / total_clinic) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_eps     = (freqs[1,3] / total_eps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel J22 = pct_eps, nformat(number_d2)
putexcel J23 = pct_cs , nformat(number_d2)
putexcel J24 = pct_ps , nformat(number_d2)
putexcel J25 = pct_case , nformat(number_d2)
putexcel J26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
*** Pharmacy
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1j, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel K33 = count_oui  , nformat(number_d0)
putexcel K5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1j s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel K7= pct_dakar, nformat(number_d2)
 putexcel K8= pct_diourbel, nformat(number_d2)
 putexcel K9= pct_fatick, nformat(number_d2)
 putexcel K10= pct_kaffrine, nformat(number_d2)
 putexcel K11= pct_kaolack, nformat(number_d2)
 putexcel K12= pct_kedougou, nformat(number_d2)
 putexcel K13= pct_kolda, nformat(number_d2)
 putexcel K14= pct_louga, nformat(number_d2)
 putexcel K15= pct_matam, nformat(number_d2)
 putexcel K16= pct_stlouis, nformat(number_d2)
 putexcel K17= pct_sedhiou, nformat(number_d2)
 putexcel K18= pct_tambacounda, nformat(number_d2)
 putexcel K19= pct_thies, nformat(number_d2)
 putexcel K20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1j s1_7, matcell(freqs) matrow(names)
scalar total_clinic   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_eps    = freqs[1,3] + freqs[2,3]
scalar total_ps  = freqs[1,4] + freqs[2,4]
scalar total_case= freqs[1,5] + freqs[2,5]


scalar pct_clinic    = (freqs[1,1] / total_clinic) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_eps     = (freqs[1,3] / total_eps) * 100
scalar pct_ps   = (freqs[1,4] / total_ps) * 100
scalar pct_case = (freqs[1,5] / total_case) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel K22 = pct_eps, nformat(number_d2)
putexcel K23 = pct_cs , nformat(number_d2)
putexcel K24 = pct_ps , nformat(number_d2)
putexcel K25 = pct_case , nformat(number_d2)
putexcel K26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
***Orientation Signs
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1k, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel L33 = count_oui  , nformat(number_d0)
putexcel L5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1k s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel L7= pct_dakar, nformat(number_d2)
 putexcel L8= pct_diourbel, nformat(number_d2)
 putexcel L9= pct_fatick, nformat(number_d2)
 putexcel L10= pct_kaffrine, nformat(number_d2)
 putexcel L11= pct_kaolack, nformat(number_d2)
 putexcel L12= pct_kedougou, nformat(number_d2)
 putexcel L13= pct_kolda, nformat(number_d2)
 putexcel L14= pct_louga, nformat(number_d2)
 putexcel L15= pct_matam, nformat(number_d2)
 putexcel L16= pct_stlouis, nformat(number_d2)
 putexcel L17= pct_sedhiou, nformat(number_d2)
 putexcel L18= pct_tambacounda, nformat(number_d2)
 putexcel L19= pct_thies, nformat(number_d2)
 putexcel L20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1k s1_7, matcell(freqs) matrow(names)
scalar total_clinic   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_eps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_clinic= freqs[1,5] + freqs[2,5]


scalar pct_clinic    = (freqs[1,1] / total_clinic) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_eps     = (freqs[1,3] / total_eps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_cs = (freqs[1,5] / total_cs) * 100


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel L22 = pct_eps, nformat(number_d2)
putexcel L23 = pct_cs , nformat(number_d2)
putexcel L24 = pct_ps , nformat(number_d2)
putexcel L25 = pct_case , nformat(number_d2)
putexcel L26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
*** Ramps for people with disabilities
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1l, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel M33 = count_oui  , nformat(number_d0)
putexcel M5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1l s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel M7= pct_dakar, nformat(number_d2)
 putexcel M8= pct_diourbel, nformat(number_d2)
 putexcel M9= pct_fatick, nformat(number_d2)
 putexcel M10= pct_kaffrine, nformat(number_d2)
 putexcel M11= pct_kaolack, nformat(number_d2)
 putexcel M12= pct_kedougou, nformat(number_d2)
 putexcel M13= pct_kolda, nformat(number_d2)
 putexcel M14= pct_louga, nformat(number_d2)
 putexcel M15= pct_matam, nformat(number_d2)
 putexcel M16= pct_stlouis, nformat(number_d2)
 putexcel M17= pct_sedhiou, nformat(number_d2)
 putexcel M18= pct_tambacounda, nformat(number_d2)
 putexcel M19= pct_thies, nformat(number_d2)
 putexcel M20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1l s1_7, matcell(freqs) matrow(names)
scalar total_clinic   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_eps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_clinic    = (freqs[1,1] / total_clinic) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_eps     = (freqs[1,3] / total_eps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel M22 = pct_eps, nformat(number_d2)
putexcel M23 = pct_cs , nformat(number_d2)
putexcel M24 = pct_ps , nformat(number_d2)
putexcel M25 = pct_case , nformat(number_d2)
putexcel M26 = pct_clinic , nformat(number_d2)
putexcel close



*///////////////////////////////////////////////////////////////////////////////
*** Biomedical waste collection room
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1m, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel N33 = count_oui  , nformat(number_d0)
putexcel N5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1m s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel N7= pct_dakar, nformat(number_d2)
 putexcel N8= pct_diourbel, nformat(number_d2)
 putexcel N9= pct_fatick, nformat(number_d2)
 putexcel N10= pct_kaffrine, nformat(number_d2)
 putexcel N11= pct_kaolack, nformat(number_d2)
 putexcel N12= pct_kedougou, nformat(number_d2)
 putexcel N13= pct_kolda, nformat(number_d2)
 putexcel N14= pct_louga, nformat(number_d2)
 putexcel N15= pct_matam, nformat(number_d2)
 putexcel N16= pct_stlouis, nformat(number_d2)
 putexcel N17= pct_sedhiou, nformat(number_d2)
 putexcel N18= pct_tambacounda, nformat(number_d2)
 putexcel N19= pct_thies, nformat(number_d2)
 putexcel N20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1m s1_7, matcell(freqs) matrow(names)
scalar total_clinic   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_eps    = freqs[1,3] + freqs[2,3]
scalar total_ps  = freqs[1,4] + freqs[2,4]
scalar total_case= freqs[1,5] + freqs[2,5]


scalar pct_clinic = (freqs[1,1]/ total_clinic) * 100
scalar pct_cs = (freqs[1,2]/ total_cs) * 100
scalar total_eps = (freqs[1,3]/ total_eps) * 100
scalar pct_case = (freqs[1,4]/ total_case) * 100
scalar pct_ps = (freqs[1,5]/ total_ps) * 100


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel N22 = pct_eps, nformat(number_d2)
putexcel N23 = pct_cs , nformat(number_d2)
putexcel N24 = pct_ps , nformat(number_d2)
putexcel N25 = pct_case , nformat(number_d2)
putexcel N26 = pct_clinic , nformat(number_d2)
putexcel close






*///////////////////////////////////////////////////////////////////////////////
*** Car parking
*///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab s3_1n, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
//putexcel O33 = count_oui  , nformat(number_d0)
putexcel O5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab s3_1n s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify

// Percentages "Oui"
 putexcel O7= pct_dakar, nformat(number_d2)
 putexcel O8= pct_diourbel, nformat(number_d2)
 putexcel O9= pct_fatick, nformat(number_d2)
 putexcel O10= pct_kaffrine, nformat(number_d2)
 putexcel O11= pct_kaolack, nformat(number_d2)
 putexcel O12= pct_kedougou, nformat(number_d2)
 putexcel O13= pct_kolda, nformat(number_d2)
 putexcel O14= pct_louga, nformat(number_d2)
 putexcel O15= pct_matam, nformat(number_d2)
 putexcel O16= pct_stlouis, nformat(number_d2)
 putexcel O17= pct_sedhiou, nformat(number_d2)
 putexcel O18= pct_tambacounda, nformat(number_d2)
 putexcel O19= pct_thies, nformat(number_d2)
 putexcel O20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab s3_1n s1_7, matcell(freqs) matrow(names)
scalar total_eps   = freqs[1,1] + freqs[2,1]
scalar total_cs    = freqs[1,2] + freqs[2,2]
scalar total_ps    = freqs[1,3] + freqs[2,3]
scalar total_case  = freqs[1,4] + freqs[2,4]
scalar total_clinic= freqs[1,5] + freqs[2,5]


scalar pct_eps    = (freqs[1,1] / total_eps) * 100
scalar pct_cs     = (freqs[1,2] / total_cs) * 100
scalar pct_ps     = (freqs[1,3] / total_ps) * 100
scalar pct_case   = (freqs[1,4] / total_case) * 100
scalar pct_clinic = (freqs[1,5] / total_clinic) * 100


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.1") modify
putexcel O22 = pct_eps, nformat(number_d2)
putexcel O23 = pct_cs , nformat(number_d2)
putexcel O24 = pct_ps , nformat(number_d2)
putexcel O25 = pct_case , nformat(number_d2)
putexcel O26 = pct_clinic , nformat(number_d2)
putexcel close









*8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.1.1.2.1.   % of facilities that regularly provide FP method 											
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
use tabulation_data_withouts4.dta, clear


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"

putexcel A2 = "Table 5.1.1.2.1  % of facilities that regularly provide FP method ", bold


putexcel B3 = "DUI", bold
putexcel C3 = "Injectables", bold
putexcel D3 = "Implants", bold
putexcel E3 = "Male condoms", bold
putexcel F3 = "Femaleale condoms", bold
putexcel G3 = "Pills", bold
putexcel H3 = "Emergency pills", bold
putexcel I3 = "Male seterilisation", bold
putexcel J3 = "Female seterilisation", bold








///////////////////////////////////////////////////////////////////////////////
*** DUI
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab dui_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel K5 = total_responses  , nformat(number_d0)
putexcel B5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab dui_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
 putexcel K7 = total_dakar, nformat(number_d0)
 putexcel K8= total_diourbel, nformat(number_d0) 
 putexcel K9= total_fatick , nformat(number_d0) 
 putexcel K10= total_kaffrine , nformat(number_d0) 
 putexcel K11= total_kaolack , nformat(number_d0) 
 putexcel K12= total_kedougou , nformat(number_d0) 
 putexcel K13= total_kolda , nformat(number_d0) 
 putexcel K14= total_louga , nformat(number_d0) 
 putexcel K15= total_matam , nformat(number_d0) 
 putexcel K16= total_stlouis , nformat(number_d0) 
 putexcel K17= total_sedhiou , nformat(number_d0) 
 putexcel K18= total_tambacounda , nformat(number_d0) 
 putexcel K19= total_thies , nformat(number_d0) 
 putexcel K20= total_ziguinchor , nformat(number_d0) 


// Percentages "Oui"
 putexcel B7= pct_dakar, nformat(number_d2)
 putexcel B8= pct_diourbel, nformat(number_d2)
 putexcel B9= pct_fatick, nformat(number_d2)
 putexcel B10= pct_kaffrine, nformat(number_d2)
 putexcel B11= pct_kaolack, nformat(number_d2)
 putexcel B12= pct_kedougou, nformat(number_d2)
 putexcel B13= pct_kolda, nformat(number_d2)
 putexcel B14= pct_louga, nformat(number_d2)
 putexcel B15= pct_matam, nformat(number_d2)
 putexcel B16= pct_stlouis, nformat(number_d2)
 putexcel B17= pct_sedhiou, nformat(number_d2)
 putexcel B18= pct_tambacounda, nformat(number_d2)
 putexcel B19= pct_thies, nformat(number_d2)
 putexcel B20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab dui_daily s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel B22 = pct_eps, nformat(number_d2)
putexcel B23 = pct_cs , nformat(number_d2)
putexcel B24 = pct_ps , nformat(number_d2)
putexcel B25 = pct_case , nformat(number_d2)
putexcel B26 = pct_clinic , nformat(number_d2)


putexcel K22 = total_eps , nformat(number_d0)
putexcel K23 = total_cs , nformat(number_d0)
putexcel K24 = total_ps , nformat(number_d0)
putexcel K25 = total_case , nformat(number_d0)
putexcel K26 = total_clinic, nformat(number_d0)
putexcel close



///////////////////////////////////////////////////////////////////////////////
***Injectables
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab inject_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel C33 = count_oui  , nformat(number_d0)
putexcel C5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab inject_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel C7= pct_dakar, nformat(number_d2)
 putexcel C8= pct_diourbel, nformat(number_d2)
 putexcel C9= pct_fatick, nformat(number_d2)
 putexcel C10= pct_kaffrine, nformat(number_d2)
 putexcel C11= pct_kaolack, nformat(number_d2)
 putexcel C12= pct_kedougou, nformat(number_d2)
 putexcel C13= pct_kolda, nformat(number_d2)
 putexcel C14= pct_louga, nformat(number_d2)
 putexcel C15= pct_matam, nformat(number_d2)
 putexcel C16= pct_stlouis, nformat(number_d2)
 putexcel C17= pct_sedhiou, nformat(number_d2)
 putexcel C18= pct_tambacounda, nformat(number_d2)
 putexcel C19= pct_thies, nformat(number_d2)
 putexcel C20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab inject_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100





putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel C22 = pct_eps, nformat(number_d2)
putexcel C23 = pct_cs , nformat(number_d2)
putexcel C24 = pct_ps , nformat(number_d2)
putexcel C25 = pct_case , nformat(number_d2)
putexcel C26 = pct_clinic , nformat(number_d2)
putexcel close




///////////////////////////////////////////////////////////////////////////////
*** Implants
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab implant_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel D33 = count_oui  , nformat(number_d0)
putexcel D5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab implant_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel D7= pct_dakar, nformat(number_d2)
 putexcel D8= pct_diourbel, nformat(number_d2)
 putexcel D9= pct_fatick, nformat(number_d2)
 putexcel D10= pct_kaffrine, nformat(number_d2)
 putexcel D11= pct_kaolack, nformat(number_d2)
 putexcel D12= pct_kedougou, nformat(number_d2)
 putexcel D13= pct_kolda, nformat(number_d2)
 putexcel D14= pct_louga, nformat(number_d2)
 putexcel D15= pct_matam, nformat(number_d2)
 putexcel D16= pct_stlouis, nformat(number_d2)
 putexcel D17= pct_sedhiou, nformat(number_d2)
 putexcel D18= pct_tambacounda, nformat(number_d2)
 putexcel D19= pct_thies, nformat(number_d2)
 putexcel D20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab implant_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel D22 = pct_eps, nformat(number_d2)
putexcel D23 = pct_cs , nformat(number_d2)
putexcel D24 = pct_ps , nformat(number_d2)
putexcel D25 = pct_case , nformat(number_d2)
putexcel D26 = pct_clinic , nformat(number_d2)
putexcel close




///////////////////////////////////////////////////////////////////////////////
*** Male condoms
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab male_condoms_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel E33 = count_oui  , nformat(number_d0)
putexcel E5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab male_condoms_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel E7= pct_dakar, nformat(number_d2)
 putexcel E8= pct_diourbel, nformat(number_d2)
 putexcel E9= pct_fatick, nformat(number_d2)
 putexcel E10= pct_kaffrine, nformat(number_d2)
 putexcel E11= pct_kaolack, nformat(number_d2)
 putexcel E12= pct_kedougou, nformat(number_d2)
 putexcel E13= pct_kolda, nformat(number_d2)
 putexcel E14= pct_louga, nformat(number_d2)
 putexcel E15= pct_matam, nformat(number_d2)
 putexcel E16= pct_stlouis, nformat(number_d2)
 putexcel E17= pct_sedhiou, nformat(number_d2)
 putexcel E18= pct_tambacounda, nformat(number_d2)
 putexcel E19= pct_thies, nformat(number_d2)
 putexcel E20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab male_condoms_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel E22 = pct_eps, nformat(number_d2)
putexcel E23 = pct_cs , nformat(number_d2)
putexcel E24 = pct_ps , nformat(number_d2)
putexcel E25 = pct_case , nformat(number_d2)
putexcel E26 = pct_clinic , nformat(number_d2)
putexcel close





///////////////////////////////////////////////////////////////////////////////
*** Femaleale condoms
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab female_condoms_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel F33 = count_oui  , nformat(number_d0)
putexcel F5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab female_condoms_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel F7= pct_dakar, nformat(number_d2)
 putexcel F8= pct_diourbel, nformat(number_d2)
 putexcel F9= pct_fatick, nformat(number_d2)
 putexcel F10= pct_kaffrine, nformat(number_d2)
 putexcel F11= pct_kaolack, nformat(number_d2)
 putexcel F12= pct_kedougou, nformat(number_d2)
 putexcel F13= pct_kolda, nformat(number_d2)
 putexcel F14= pct_louga, nformat(number_d2)
 putexcel F15= pct_matam, nformat(number_d2)
 putexcel F16= pct_stlouis, nformat(number_d2)
 putexcel F17= pct_sedhiou, nformat(number_d2)
 putexcel F18= pct_tambacounda, nformat(number_d2)
 putexcel F19= pct_thies, nformat(number_d2)
 putexcel F20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab female_condoms_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel F22 = pct_eps, nformat(number_d2)
putexcel F23 = pct_cs , nformat(number_d2)
putexcel F24 = pct_ps , nformat(number_d2)
putexcel F25 = pct_case , nformat(number_d2)
putexcel F26 = pct_clinic , nformat(number_d2)
putexcel close



///////////////////////////////////////////////////////////////////////////////
*** Pills
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab pills_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel G33 = count_oui  , nformat(number_d0)
putexcel G5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab pills_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel G7= pct_dakar, nformat(number_d2)
 putexcel G8= pct_diourbel, nformat(number_d2)
 putexcel G9= pct_fatick, nformat(number_d2)
 putexcel G10= pct_kaffrine, nformat(number_d2)
 putexcel G11= pct_kaolack, nformat(number_d2)
 putexcel G12= pct_kedougou, nformat(number_d2)
 putexcel G13= pct_kolda, nformat(number_d2)
 putexcel G14= pct_louga, nformat(number_d2)
 putexcel G15= pct_matam, nformat(number_d2)
 putexcel G16= pct_stlouis, nformat(number_d2)
 putexcel G17= pct_sedhiou, nformat(number_d2)
 putexcel G18= pct_tambacounda, nformat(number_d2)
 putexcel G19= pct_thies, nformat(number_d2)
 putexcel G20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab pills_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel G22 = pct_eps, nformat(number_d2)
putexcel G23 = pct_cs , nformat(number_d2)
putexcel G24 = pct_ps , nformat(number_d2)
putexcel G25 = pct_case , nformat(number_d2)
putexcel G26 = pct_clinic , nformat(number_d2)
putexcel close



///////////////////////////////////////////////////////////////////////////////
*** Emergency pills
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab pcu_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel H33 = count_oui  , nformat(number_d0)
putexcel H5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab pcu_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel H7= pct_dakar, nformat(number_d2)
 putexcel H8= pct_diourbel, nformat(number_d2)
 putexcel H9= pct_fatick, nformat(number_d2)
 putexcel H10= pct_kaffrine, nformat(number_d2)
 putexcel H11= pct_kaolack, nformat(number_d2)
 putexcel H12= pct_kedougou, nformat(number_d2)
 putexcel H13= pct_kolda, nformat(number_d2)
 putexcel H14= pct_louga, nformat(number_d2)
 putexcel H15= pct_matam, nformat(number_d2)
 putexcel H16= pct_stlouis, nformat(number_d2)
 putexcel H17= pct_sedhiou, nformat(number_d2)
 putexcel H18= pct_tambacounda, nformat(number_d2)
 putexcel H19= pct_thies, nformat(number_d2)
 putexcel H20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab pcu_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel H22 = pct_eps, nformat(number_d2)
putexcel H23 = pct_cs , nformat(number_d2)
putexcel H24 = pct_ps , nformat(number_d2)
putexcel H25 = pct_case , nformat(number_d2)
putexcel H26 = pct_clinic , nformat(number_d2)
putexcel close


///////////////////////////////////////////////////////////////////////////////
*** Male sterilization
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab male_ster_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel I33 = count_oui  , nformat(number_d0)
putexcel I5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab male_ster_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel I7= pct_dakar, nformat(number_d2)
 putexcel I8= pct_diourbel, nformat(number_d2)
 putexcel I9= pct_fatick, nformat(number_d2)
 putexcel I10= pct_kaffrine, nformat(number_d2)
 putexcel I11= pct_kaolack, nformat(number_d2)
 putexcel I12= pct_kedougou, nformat(number_d2)
 putexcel I13= pct_kolda, nformat(number_d2)
 putexcel I14= pct_louga, nformat(number_d2)
 putexcel I15= pct_matam, nformat(number_d2)
 putexcel I16= pct_stlouis, nformat(number_d2)
 putexcel I17= pct_sedhiou, nformat(number_d2)
 putexcel I18= pct_tambacounda, nformat(number_d2)
 putexcel I19= pct_thies, nformat(number_d2)
 putexcel I20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab male_ster_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel I22 = pct_eps, nformat(number_d2)
putexcel I23 = pct_cs , nformat(number_d2)
putexcel I24 = pct_ps , nformat(number_d2)
putexcel I25 = pct_case , nformat(number_d2)
putexcel I26 = pct_clinic , nformat(number_d2)
putexcel close



///////////////////////////////////////////////////////////////////////////////
*** Female sterilisation
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab female_ster_daily, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
//putexcel J33 = count_oui  , nformat(number_d0)
putexcel J5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab female_ster_daily s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify

// Percentages "Oui"
 putexcel J7= pct_dakar, nformat(number_d2)
 putexcel J8= pct_diourbel, nformat(number_d2)
 putexcel J9= pct_fatick, nformat(number_d2)
 putexcel J10= pct_kaffrine, nformat(number_d2)
 putexcel J11= pct_kaolack, nformat(number_d2)
 putexcel J12= pct_kedougou, nformat(number_d2)
 putexcel J13= pct_kolda, nformat(number_d2)
 putexcel J14= pct_louga, nformat(number_d2)
 putexcel J15= pct_matam, nformat(number_d2)
 putexcel J16= pct_stlouis, nformat(number_d2)
 putexcel J17= pct_sedhiou, nformat(number_d2)
 putexcel J18= pct_tambacounda, nformat(number_d2)
 putexcel J19= pct_thies, nformat(number_d2)
 putexcel J20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab female_ster_daily s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100





putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.1") modify
putexcel J22 = pct_eps, nformat(number_d2)
putexcel J23 = pct_cs , nformat(number_d2)
putexcel J24 = pct_ps , nformat(number_d2)
putexcel J25 = pct_case , nformat(number_d2)
putexcel J26 = pct_clinic , nformat(number_d2)
putexcel close





*8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table Table 5.1.1.2.2.  % of facilities that regularly provide each MNCH service 										
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
use tabulation_data_withouts4.dta, clear


putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"


putexcel A2 = "Table 5.1.1.2.2.  % of facilities that regularly provide each MNCH service", bold


putexcel B3 = "Antenatal care", bold
putexcel C3 = "Delivery care", bold
putexcel D3 = "Postnatal care", bold
putexcel E3 = "Essential services for newborns", bold
putexcel F3 = "Post-abortion services", bold
putexcel G3 = "Child and Newborn Health Services", bold


**Generating the ANC variable
capture drop ANC
gen ANC =.
replace ANC = 1 if s5_1A_a == 1 | s5_1A_a == 0 
replace ANC = 0 if s5_1A_a == -9 | s5_1A_a == . 
label val ANC yesno


***Delivery care
capture drop delivery_care
gen delivery_care =.
replace delivery_care = 1 if s5_1B_a == 1 | s5_1B_a == 0 
replace delivery_care = 0 if s5_1B_a == -9 | s5_1B_a == . 
label val delivery_care yesno



***Postnatal care
capture drop PNC
gen PNC =.
replace PNC = 1 if s5_1C_a == 1 | s5_1C_a == 0 
replace PNC = 0 if s5_1C_a == -9 | s5_1C_a == . 
label val PNC yesno



***Essential services for newborns 
capture drop essential_newborn
gen essential_newborn =.
replace essential_newborn = 1 if s5_1D_a == 1 | s5_1D_a == 0 
replace essential_newborn = 0 if s5_1D_a == -9 | s5_1D_a == . 
label val essential_newborn yesno




***Post-abortion services 
capture drop post_abortion
gen post_abortion =.
replace post_abortion = 1 if s5_1E_a == 1 | s5_1E_a == 0 
replace post_abortion = 0 if s5_1E_a == -9 | s5_1E_a == . 
label val post_abortion yesno



***Child and Newborn Health Services 
capture drop child_newborn_services
gen child_newborn_services =.
replace child_newborn_services = 1 if s5_1F_a == 1 | s5_1F_a == 0 
replace child_newborn_services = 0 if s5_1F_a == -9 | s5_1F_a == . 
label val child_newborn_services yesno



///////////////////////////////////////////////////////////////////////////////
** Antenatal care
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab ANC, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel H5 = total_responses  , nformat(number_d0)
putexcel B5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab ANC s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
 putexcel H7 = total_dakar, nformat(number_d0)
 putexcel H8= total_diourbel, nformat(number_d0) 
 putexcel H9= total_fatick , nformat(number_d0) 
 putexcel H10= total_kaffrine , nformat(number_d0) 
 putexcel H11= total_kaolack , nformat(number_d0) 
 putexcel H12= total_kedougou , nformat(number_d0) 
 putexcel H13= total_kolda , nformat(number_d0) 
 putexcel H14= total_louga , nformat(number_d0) 
 putexcel H15= total_matam , nformat(number_d0) 
 putexcel H16= total_stlouis , nformat(number_d0) 
 putexcel H17= total_sedhiou , nformat(number_d0) 
 putexcel H18= total_tambacounda , nformat(number_d0) 
 putexcel H19= total_thies , nformat(number_d0) 
 putexcel H20= total_ziguinchor , nformat(number_d0) 


// Percentages "Oui"
 putexcel B7= pct_dakar, nformat(number_d2)
 putexcel B8= pct_diourbel, nformat(number_d2)
 putexcel B9= pct_fatick, nformat(number_d2)
 putexcel B10= pct_kaffrine, nformat(number_d2)
 putexcel B11= pct_kaolack, nformat(number_d2)
 putexcel B12= pct_kedougou, nformat(number_d2)
 putexcel B13= pct_kolda, nformat(number_d2)
 putexcel B14= pct_louga, nformat(number_d2)
 putexcel B15= pct_matam, nformat(number_d2)
 putexcel B16= pct_stlouis, nformat(number_d2)
 putexcel B17= pct_sedhiou, nformat(number_d2)
 putexcel B18= pct_tambacounda, nformat(number_d2)
 putexcel B19= pct_thies, nformat(number_d2)
 putexcel B20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab ANC s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel B22 = pct_eps, nformat(number_d2)
putexcel B23 = pct_cs , nformat(number_d2)
putexcel B24 = pct_ps , nformat(number_d2)
putexcel B25 = pct_case , nformat(number_d2)
putexcel B26 = pct_clinic , nformat(number_d2)


putexcel H22 = total_eps , nformat(number_d0)
putexcel H23 = total_cs , nformat(number_d0)
putexcel H24 = total_ps , nformat(number_d0)
putexcel H25 = total_case , nformat(number_d0)
putexcel H26 = total_clinic, nformat(number_d0)
putexcel close







///////////////////////////////////////////////////////////////////////////////
*** Delivery care
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab delivery_care, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
//putexcel D33 = count_oui  , nformat(number_d0)
putexcel C5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab delivery_care s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify

// Percentages "Oui"
 putexcel C7= pct_dakar, nformat(number_d2)
 putexcel C8= pct_diourbel, nformat(number_d2)
 putexcel C9= pct_fatick, nformat(number_d2)
 putexcel C10= pct_kaffrine, nformat(number_d2)
 putexcel C11= pct_kaolack, nformat(number_d2)
 putexcel C12= pct_kedougou, nformat(number_d2)
 putexcel C13= pct_kolda, nformat(number_d2)
 putexcel C14= pct_louga, nformat(number_d2)
 putexcel C15= pct_matam, nformat(number_d2)
 putexcel C16= pct_stlouis, nformat(number_d2)
 putexcel C17= pct_sedhiou, nformat(number_d2)
 putexcel C18= pct_tambacounda, nformat(number_d2)
 putexcel C19= pct_thies, nformat(number_d2)
 putexcel C20= pct_ziguinchor, nformat(number_d2)
putexcel close





***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab delivery_care s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel C22 = pct_eps, nformat(number_d2)
putexcel C23 = pct_cs , nformat(number_d2)
putexcel C24 = pct_ps , nformat(number_d2)
putexcel C25 = pct_case , nformat(number_d2)
putexcel C26 = pct_clinic , nformat(number_d2)
putexcel close





///////////////////////////////////////////////////////////////////////////////
***Postnatal care
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab PNC, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel D5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab PNC s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




 

// Percentages "Oui"
 putexcel D7= pct_dakar, nformat(number_d2)
 putexcel D8= pct_diourbel, nformat(number_d2)
 putexcel D9= pct_fatick, nformat(number_d2)
 putexcel D10= pct_kaffrine, nformat(number_d2)
 putexcel D11= pct_kaolack, nformat(number_d2)
 putexcel D12= pct_kedougou, nformat(number_d2)
 putexcel D13= pct_kolda, nformat(number_d2)
 putexcel D14= pct_louga, nformat(number_d2)
 putexcel D15= pct_matam, nformat(number_d2)
 putexcel D16= pct_stlouis, nformat(number_d2)
 putexcel D17= pct_sedhiou, nformat(number_d2)
 putexcel D18= pct_tambacounda, nformat(number_d2)
 putexcel D19= pct_thies, nformat(number_d2)
 putexcel D20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab PNC s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel D22 = pct_eps, nformat(number_d2)
putexcel D23 = pct_cs , nformat(number_d2)
putexcel D24 = pct_ps , nformat(number_d2)
putexcel D25 = pct_case , nformat(number_d2)
putexcel D26 = pct_clinic , nformat(number_d2)
putexcel close




///////////////////////////////////////////////////////////////////////////////
***Essential services for newborns
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab essential_newborn, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel E5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab essential_newborn s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify

// Percentages "Oui"
 putexcel E7= pct_dakar, nformat(number_d2)
 putexcel E8= pct_diourbel, nformat(number_d2)
 putexcel E9= pct_fatick, nformat(number_d2)
 putexcel E10= pct_kaffrine, nformat(number_d2)
 putexcel E11= pct_kaolack, nformat(number_d2)
 putexcel E12= pct_kedougou, nformat(number_d2)
 putexcel E13= pct_kolda, nformat(number_d2)
 putexcel E14= pct_louga, nformat(number_d2)
 putexcel E15= pct_matam, nformat(number_d2)
 putexcel E16= pct_stlouis, nformat(number_d2)
 putexcel E17= pct_sedhiou, nformat(number_d2)
 putexcel E18= pct_tambacounda, nformat(number_d2)
 putexcel E19= pct_thies, nformat(number_d2)
 putexcel E20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab essential_newborn s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel E22 = pct_eps, nformat(number_d2)
putexcel E23 = pct_cs , nformat(number_d2)
putexcel E24 = pct_ps , nformat(number_d2)
putexcel E25 = pct_case , nformat(number_d2)
putexcel E26 = pct_clinic , nformat(number_d2)

putexcel close







///////////////////////////////////////////////////////////////////////////////
***Post-abortion services
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab PNC, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel F5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab PNC s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
// Percentages "Oui"
 putexcel F7= pct_dakar, nformat(number_d2)
 putexcel F8= pct_diourbel, nformat(number_d2)
 putexcel F9= pct_fatick, nformat(number_d2)
 putexcel F10= pct_kaffrine, nformat(number_d2)
 putexcel F11= pct_kaolack, nformat(number_d2)
 putexcel F12= pct_kedougou, nformat(number_d2)
 putexcel F13= pct_kolda, nformat(number_d2)
 putexcel F14= pct_louga, nformat(number_d2)
 putexcel F15= pct_matam, nformat(number_d2)
 putexcel F16= pct_stlouis, nformat(number_d2)
 putexcel F17= pct_sedhiou, nformat(number_d2)
 putexcel F18= pct_tambacounda, nformat(number_d2)
 putexcel F19= pct_thies, nformat(number_d2)
 putexcel F20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab PNC s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel F22 = pct_eps, nformat(number_d2)
putexcel F23 = pct_cs , nformat(number_d2)
putexcel F24 = pct_ps , nformat(number_d2)
putexcel F25 = pct_case , nformat(number_d2)
putexcel F26 = pct_clinic , nformat(number_d2)
putexcel close


///////////////////////////////////////////////////////////////////////////////
***Child and Newborn Health Services
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab child_newborn_services, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel G5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab child_newborn_services s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify

// Percentages "Oui"
 putexcel G7= pct_dakar, nformat(number_d2)
 putexcel G8= pct_diourbel, nformat(number_d2)
 putexcel G9= pct_fatick, nformat(number_d2)
 putexcel G10= pct_kaffrine, nformat(number_d2)
 putexcel G11= pct_kaolack, nformat(number_d2)
 putexcel G12= pct_kedougou, nformat(number_d2)
 putexcel G13= pct_kolda, nformat(number_d2)
 putexcel G14= pct_louga, nformat(number_d2)
 putexcel G15= pct_matam, nformat(number_d2)
 putexcel G16= pct_stlouis, nformat(number_d2)
 putexcel G17= pct_sedhiou, nformat(number_d2)
 putexcel G18= pct_tambacounda, nformat(number_d2)
 putexcel G19= pct_thies, nformat(number_d2)
 putexcel G20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab child_newborn_services s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.2.2") modify
putexcel G22 = pct_eps, nformat(number_d2)
putexcel G23 = pct_cs , nformat(number_d2)
putexcel G24 = pct_ps , nformat(number_d2)
putexcel G25 = pct_case , nformat(number_d2)
putexcel G26 = pct_clinic , nformat(number_d2)

putexcel close
















*8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.1.1.3.1.  % of facilities that have trained personnel available (specialists, doctors and nurses, as applicable) for providing each FP method
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
use tabulation_data_withouts4.dta, clear

putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.1") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"

putexcel A2 = "Table 5.1.1.3.1.  % of facilities that have trained personnel available (specialists, doctors and nurses, as applicable) for providing each FP method", bold


putexcel B3 = "DUI", bold
putexcel C3 = "Injectables", bold
putexcel D3 = "Implants", bold
putexcel E3 = "Male condoms", bold
putexcel F3 = "Femaleale condoms", bold
putexcel G3 = "Pills", bold
putexcel H3 = "Emergency pills", bold
putexcel I3 = "Male seterilisation", bold
putexcel J3 = "Female seterilisation", bold








///////////////////////////////////////////////////////////////////////////////
*** DUI
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab dui_daily_trained, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.1") modify
putexcel L5 = total_responses  , nformat(number_d0)
putexcel B5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab dui_daily_trained s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.1") modify
 putexcel L7 = total_dakar, nformat(number_d0)
 putexcel L8= total_diourbel, nformat(number_d0) 
 putexcel L9= total_fatick , nformat(number_d0) 
 putexcel L10= total_kaffrine , nformat(number_d0) 
 putexcel L11= total_kaolack , nformat(number_d0) 
 putexcel L12= total_kedougou , nformat(number_d0) 
 putexcel L13= total_kolda , nformat(number_d0) 
 putexcel L14= total_louga , nformat(number_d0) 
 putexcel L15= total_matam , nformat(number_d0) 
 putexcel L16= total_stlouis , nformat(number_d0) 
 putexcel L17= total_sedhiou , nformat(number_d0) 
 putexcel L18= total_tambacounda , nformat(number_d0) 
 putexcel L19= total_thies , nformat(number_d0) 
 putexcel L20= total_ziguinchor , nformat(number_d0) 


// Percentages "Oui"
 putexcel B7= pct_dakar, nformat(number_d2)
 putexcel B8= pct_diourbel, nformat(number_d2)
 putexcel B9= pct_fatick, nformat(number_d2)
 putexcel B10= pct_kaffrine, nformat(number_d2)
 putexcel B11= pct_kaolack, nformat(number_d2)
 putexcel B12= pct_kedougou, nformat(number_d2)
 putexcel B13= pct_kolda, nformat(number_d2)
 putexcel B14= pct_louga, nformat(number_d2)
 putexcel B15= pct_matam, nformat(number_d2)
 putexcel B16= pct_stlouis, nformat(number_d2)
 putexcel B17= pct_sedhiou, nformat(number_d2)
 putexcel B18= pct_tambacounda, nformat(number_d2)
 putexcel B19= pct_thies, nformat(number_d2)
 putexcel B20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab dui_daily_trained s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.1") modify
putexcel B22 = pct_eps, nformat(number_d2)
putexcel B23 = pct_cs , nformat(number_d2)
putexcel B24 = pct_ps , nformat(number_d2)
putexcel B25 = pct_case , nformat(number_d2)
putexcel B26 = pct_clinic , nformat(number_d2)


putexcel L22 = total_eps , nformat(number_d0)
putexcel L23 = total_cs , nformat(number_d0)
putexcel L24 = total_ps , nformat(number_d0)
putexcel L25 = total_case , nformat(number_d0)
putexcel L26 = total_clinic, nformat(number_d0)
putexcel close







///////////////////////////////////////////////////////////////////////////////
*** Implants
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab implant_daily_trained, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.1") modify
//putexcel D33 = count_oui  , nformat(number_d0)
putexcel D5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab implant_daily_trained s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.1") modify

// Percentages "Oui"
 putexcel D7= pct_dakar, nformat(number_d2)
 putexcel D8= pct_diourbel, nformat(number_d2)
 putexcel D9= pct_fatick, nformat(number_d2)
 putexcel D10= pct_kaffrine, nformat(number_d2)
 putexcel D11= pct_kaolack, nformat(number_d2)
 putexcel D12= pct_kedougou, nformat(number_d2)
 putexcel D13= pct_kolda, nformat(number_d2)
 putexcel D14= pct_louga, nformat(number_d2)
 putexcel D15= pct_matam, nformat(number_d2)
 putexcel D16= pct_stlouis, nformat(number_d2)
 putexcel D17= pct_sedhiou, nformat(number_d2)
 putexcel D18= pct_tambacounda, nformat(number_d2)
 putexcel D19= pct_thies, nformat(number_d2)
 putexcel D20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab implant_daily_trained s1_7, matcell(freqs) matrow(names)
scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100





putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.1") modify
putexcel D22 = pct_eps, nformat(number_d2)
putexcel D23 = pct_cs , nformat(number_d2)
putexcel D24 = pct_ps , nformat(number_d2)
putexcel D25 = pct_case , nformat(number_d2)
putexcel D26 = pct_clinic , nformat(number_d2)
putexcel close



*8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.1.1.3.2 % of facilities that have trained personnel available (specialists, doctors, and nurses, as applicable) for providing each MNCH service						
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
use tabulation_data_withs4, clear

**Generating the ANC variable
capture drop ANC
gen ANC =.
replace ANC = 1 if s5_1A_a == 1 | s5_1A_a == 0 
replace ANC = 0 if s5_1A_a == -9 | s5_1A_a == . 
label val ANC yesno


***Delivery care
capture drop delivery_care
gen delivery_care =.
replace delivery_care = 1 if s5_1B_a == 1 | s5_1B_a == 0 
replace delivery_care = 0 if s5_1B_a == -9 | s5_1B_a == . 
label val delivery_care yesno



***Postnatal care
capture drop PNC
gen PNC =.
replace PNC = 1 if s5_1C_a == 1 | s5_1C_a == 0 
replace PNC = 0 if s5_1C_a == -9 | s5_1C_a == . 
label val PNC yesno



***Essential services for newborns 
capture drop essential_newborn
gen essential_newborn =.
replace essential_newborn = 1 if s5_1D_a == 1 | s5_1D_a == 0 
replace essential_newborn = 0 if s5_1D_a == -9 | s5_1D_a == . 
label val essential_newborn yesno




***Post-abortion services 
capture drop post_abortion
gen post_abortion =.
replace post_abortion = 1 if s5_1E_a == 1 | s5_1E_a == 0 
replace post_abortion = 0 if s5_1E_a == -9 | s5_1E_a == . 
label val post_abortion yesno



***Child and Newborn Health Services 
capture drop child_newborn_services
gen child_newborn_services =.
replace child_newborn_services = 1 if s5_1F_a == 1 | s5_1F_a == 0 
replace child_newborn_services = 0 if s5_1F_a == -9 | s5_1F_a == . 
label val child_newborn_services yesno
tab s4_1, m
tab s4_0, m

*2 Gynecologist Gynécologue
*4 Doctor  médecin généraliste
*8 Nurse  infirmier 
*9 Midwives  Sage-femmes

putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.3.2") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"

putexcel A2 = "Table 5.1.1.3.2 % of facilities that have trained personnel available (specialists, doctors, and nurses, as applicable) for providing each MNCH service", bold


putexcel B3 = "Antenatal care", bold
putexcel B4 = "Gynecologist", bold
putexcel C4 = "Doctor", bold
putexcel D4 = "Nurse", bold
putexcel E4 = "Midwives", bold


putexcel F3 = "Delivery care", bold
putexcel F4 = "Gynecologist", bold
putexcel G4 = "Doctor", bold
putexcel H4 = "Nurse", bold
putexcel I4 = "Midwives", bold

putexcel J3 = "Postnatal care", bold
putexcel J4 = "Gynecologist", bold
putexcel K4 = "Doctor", bold
putexcel L4 = "Nurse", bold
putexcel M4 = "Midwives", bold


putexcel N3 = "Essential services for newborns", bold
putexcel N4 = "Gynecologist", bold
putexcel O4 = "Doctor", bold
putexcel P4 = "Nurse", bold
putexcel Q4 = "Midwives", bold

putexcel R3 = "Post-abortion services", bold
putexcel R4 = "Gynecologist", bold
putexcel S4 = "Doctor", bold
putexcel T4 = "Nurse", bold
putexcel U4 = "Midwives", bold

putexcel V3 = "Child and Newborn Health Services", bold
putexcel V4 = "Gynecologist", bold
putexcel W4 = "Doctor", bold
putexcel X4 = "Nurse", bold
putexcel Y4 = "Midwives", bold















































*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&88888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.1.1.4.1.  % of facilities that have adequate infrastructure/equipment/drug/supplies for providing each FP method
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
use tabulation_data_withouts4, clear

*DUI (56)
use tabulation_data_withouts4, clear

foreach var in s6_3trav_a s6_3trav_b s6_3trav_c s6_3trav_d s6_3trav_e s6_3trav_f s6_3trav_g s6_3trav_h s6_3trav_i s6_3trav_j s6_3trav_k s6_3trav_l s6_3trav_m s6_3trav_n s6_3coin_a s6_3coin_b s6_3coin_c s6_3coin_d s6_3coin_e s6_3coin_f s6_3coin_g s6_3coin_h s6_3coin_i s6_3coin_j s6_3coin_k s6_3coin_l s6_3coin_m s6_3coin_n s6_3both_a s6_3both_b s6_3both_c s6_3both_d s6_3both_e s6_3both_f s6_3both_g s6_3both_h s6_3both_i s6_3both_j s6_3both_k s6_3both_l s6_3both_m s6_3both_n s6_3other_a s6_3other_b s6_3other_c s6_3other_d s6_3other_e s6_3other_f s6_3other_g s6_3other_h s6_3other_i s6_3other_j s6_3other_k s6_3other_l s6_3other_m s6_3other_n {
    replace `var' = . if `var' != 1
}

capture drop IUD_equipments
egen IUD_equipments  = rowtotal(s6_3trav_a s6_3trav_b s6_3trav_c s6_3trav_d s6_3trav_e s6_3trav_f s6_3trav_g s6_3trav_h s6_3trav_i s6_3trav_j s6_3trav_k s6_3trav_l s6_3trav_m s6_3trav_n s6_3coin_a s6_3coin_b s6_3coin_c s6_3coin_d s6_3coin_e s6_3coin_f s6_3coin_g s6_3coin_h s6_3coin_i s6_3coin_j s6_3coin_k s6_3coin_l s6_3coin_m s6_3coin_n s6_3both_a s6_3both_b s6_3both_c s6_3both_d s6_3both_e s6_3both_f s6_3both_g s6_3both_h s6_3both_i s6_3both_j s6_3both_k s6_3both_l s6_3both_m s6_3both_n s6_3other_a s6_3other_b s6_3other_c s6_3other_d s6_3other_e s6_3other_f s6_3other_g s6_3other_h s6_3other_i s6_3other_j s6_3other_k s6_3other_l s6_3other_m s6_3other_n)

tab IUD_equipments


replace IUD_equipments = 0 if IUD_equipments < 10
replace IUD_equipments = 1 if IUD_equipments >= 10

label values IUD_equipments  yesno



**FEMALE STERILIZATION

foreach var in s6_7 s6_7a s6_7b s6_7c s6_7d s6_7e s6_7f s6_7g s6_7h s6_7j s6_8a s6_8b s6_8c s6_8d s6_8e s6_8f s6_8g s6_8h s6_8i s6_8j s6_8k s6_8l s6_8m s6_8n s6_8o s6_8p s6_8q s6_8r s6_8s s6_8t s6_8u s6_8v s6_8w s6_8x s6_8y s6_8z s6_8 s6_11a s6_11b s6_11c s6_11d s6_11e s6_11f s6_11g s6_11h s6_11i s6_11j s6_11k s6_11l s6_11m s6_11n s6_11o s6_11p s6_11q s6_11r s6_11s s6_11t s6_11u {
    replace `var' = . if `var' != 1
}

capture drop female_ster_equipments
egen female_ster_equipments  = rowtotal(s6_7 s6_7a s6_7b s6_7c s6_7d s6_7e s6_7f s6_7g s6_7h s6_7j s6_8a s6_8b s6_8c s6_8d s6_8e s6_8f s6_8g s6_8h s6_8i s6_8j s6_8k s6_8l s6_8m s6_8n s6_8o s6_8p s6_8q s6_8r s6_8s s6_8t s6_8u s6_8v s6_8w s6_8x s6_8y s6_8z s6_8 s6_11a s6_11b s6_11c s6_11d s6_11e s6_11f s6_11g s6_11h s6_11i s6_11j s6_11k s6_11l s6_11m s6_11n s6_11o s6_11p s6_11q s6_11r s6_11s s6_11t s6_11u)

tab female_ster_equipments


replace female_ster_equipments = 0 if female_ster_equipments < 10
replace female_ster_equipments = 1 if female_ster_equipments >= 10

label values female_ster_equipments  yesno




**MALE STERILIZATION

foreach var in s6_14a s6_14b s6_14c s6_14d s6_14e s6_14f s6_14g s6_14h s6_14i {
    replace `var' = . if `var' != 1
}

capture drop male_ster_equipments
egen male_ster_equipments  = rowtotal(s6_14a s6_14b s6_14c s6_14d s6_14e s6_14f s6_14g s6_14h s6_14i)

tab male_ster_equipments


replace male_ster_equipments = 0 if male_ster_equipments < 5
replace male_ster_equipments = 1 if male_ster_equipments >= 5

label values male_ster_equipments  yesno



putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"

putexcel A2 = "Table 5.1.1.4.1.  % of facilities that have adequate infrastructure/equipment/drug/supplies for providing each FP method", bold


putexcel B3 = "DUI", bold
putexcel C3 = "Male sterilization", bold
putexcel D3 = "Female sterilisation", bold




///////////////////////////////////////////////////////////////////////////////
***DUI
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab IUD_equipments, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel E5 = total_responses  , nformat(number_d0)
putexcel B5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab IUD_equipments s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
 putexcel E7 = total_dakar, nformat(number_d0)
 putexcel E8= total_diourbel, nformat(number_d0) 
 putexcel E9= total_fatick , nformat(number_d0) 
 putexcel E10= total_kaffrine , nformat(number_d0) 
 putexcel E11= total_kaolack , nformat(number_d0) 
 putexcel E12= total_kedougou , nformat(number_d0) 
 putexcel E13= total_kolda , nformat(number_d0) 
 putexcel E14= total_louga , nformat(number_d0) 
 putexcel E15= total_matam , nformat(number_d0) 
 putexcel E16= total_stlouis , nformat(number_d0) 
 putexcel E17= total_sedhiou , nformat(number_d0) 
 putexcel E18= total_tambacounda , nformat(number_d0) 
 putexcel E19= total_thies , nformat(number_d0) 
 putexcel E20= total_ziguinchor , nformat(number_d0) 

// Percentages "Oui"
 putexcel B7= pct_dakar, nformat(number_d2)
 putexcel B8= pct_diourbel, nformat(number_d2)
 putexcel B9= pct_fatick, nformat(number_d2)
 putexcel B10= pct_kaffrine, nformat(number_d2)
 putexcel B11= pct_kaolack, nformat(number_d2)
 putexcel B12= pct_kedougou, nformat(number_d2)
 putexcel B13= pct_kolda, nformat(number_d2)
 putexcel B14= pct_louga, nformat(number_d2)
 putexcel B15= pct_matam, nformat(number_d2)
 putexcel B16= pct_stlouis, nformat(number_d2)
 putexcel B17= pct_sedhiou, nformat(number_d2)
 putexcel B18= pct_tambacounda, nformat(number_d2)
 putexcel B19= pct_thies, nformat(number_d2)
 putexcel B20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab IUD_equipments s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel B22 = pct_eps, nformat(number_d2)
putexcel B23 = pct_cs , nformat(number_d2)
putexcel B24 = pct_ps , nformat(number_d2)
putexcel B25 = pct_case , nformat(number_d2)
putexcel B26 = pct_clinic , nformat(number_d2)


putexcel E22 = total_eps , nformat(number_d0)
putexcel E23 = total_cs , nformat(number_d0)
putexcel E24 = total_ps , nformat(number_d0)
putexcel E25 = total_case , nformat(number_d0)
putexcel E26 = total_clinic, nformat(number_d0)
putexcel close





///////////////////////////////////////////////////////////////////////////////
***Male sterilization
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab male_ster_equipments, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel C5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab male_ster_equipments s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
 

// Percentages "Oui"
 putexcel C7= pct_dakar, nformat(number_d2)
 putexcel C8= pct_diourbel, nformat(number_d2)
 putexcel C9= pct_fatick, nformat(number_d2)
 putexcel C10= pct_kaffrine, nformat(number_d2)
 putexcel C11= pct_kaolack, nformat(number_d2)
 putexcel C12= pct_kedougou, nformat(number_d2)
 putexcel C13= pct_kolda, nformat(number_d2)
 putexcel C14= pct_louga, nformat(number_d2)
 putexcel C15= pct_matam, nformat(number_d2)
 putexcel C16= pct_stlouis, nformat(number_d2)
 putexcel C17= pct_sedhiou, nformat(number_d2)
 putexcel C18= pct_tambacounda, nformat(number_d2)
 putexcel C19= pct_thies, nformat(number_d2)
 putexcel C20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab male_ster_equipments s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel C22 = pct_eps, nformat(number_d2)
putexcel C23 = pct_cs , nformat(number_d2)
putexcel C24 = pct_ps , nformat(number_d2)
putexcel C25 = pct_case , nformat(number_d2)
putexcel C26 = pct_clinic , nformat(number_d2)

putexcel close





///////////////////////////////////////////////////////////////////////////////
***Female sterilisation
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab female_ster_equipments, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel C5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab female_ster_equipments s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
 

// Percentages "Oui"
putexcel D7= pct_dakar, nformat(number_d2)
putexcel D8= pct_diourbel, nformat(number_d2)
putexcel D9= pct_fatick, nformat(number_d2)
putexcel D10= pct_kaffrine, nformat(number_d2)
putexcel D11= pct_kaolack, nformat(number_d2)
putexcel D12= pct_kedougou, nformat(number_d2)
putexcel D13= pct_kolda, nformat(number_d2)
putexcel D14= pct_louga, nformat(number_d2)
putexcel D15= pct_matam, nformat(number_d2)
putexcel D16= pct_stlouis, nformat(number_d2)
putexcel D17= pct_sedhiou, nformat(number_d2)
putexcel D18= pct_tambacounda, nformat(number_d2)
putexcel D19= pct_thies, nformat(number_d2)
putexcel D20= pct_ziguinchor, nformat(number_d2)
putexcel close



***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab female_ster_equipments s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel D22 = pct_eps, nformat(number_d2)
putexcel D23 = pct_cs , nformat(number_d2)
putexcel D24 = pct_ps , nformat(number_d2)
putexcel D25 = pct_case , nformat(number_d2)
putexcel D26 = pct_clinic , nformat(number_d2)

putexcel close



*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&88888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.1.1.4.2. % of facilities that have adequate infrastructure/equipment/drug/supplies for providing each MNCH service
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
*Delivery care	
use tabulation_data_withouts4, clear

foreach var in s3_3a s3_3b s3_3c s3_3d s3_3e s3_3g s3_3h s3_3f s3_3i s3_3j s3_3k s3_3l s3_3m s3_3n s3_3o s3_3p s3_3q s3_3r s3_3s s3_3t s3_3u s3_3v s3_3w s3_3x s3_3y s3_3z s3_3aa s3_3ab s3_3ac s3_3ad s3_3ae s3_3af s3_3ag s3_3ah s3_3ai s3_3aj s3_3ak s3_3al s3_3am s3_3an s3_3ao s3_3ap s3_3aq s3_3ar s3_3as s3_3at s3_3au s3_3av s3_3aw s3_3ax s3_3ay s3_3az s3_3ba s3_3bb s3_3bc s3_3bd s3_3be s3_3bf {
    replace `var' = . if `var' != 1
}

capture drop delivery_care_equipments
egen delivery_care_equipments  = rowtotal(s3_3a s3_3b s3_3c s3_3d s3_3e s3_3g s3_3h s3_3f s3_3i s3_3j s3_3k s3_3l s3_3m s3_3n s3_3o s3_3p s3_3q s3_3r s3_3s s3_3t s3_3u s3_3v s3_3w s3_3x s3_3y s3_3z s3_3aa s3_3ab s3_3ac s3_3ad s3_3ae s3_3af s3_3ag s3_3ah s3_3ai s3_3aj s3_3ak s3_3al s3_3am s3_3an s3_3ao s3_3ap s3_3aq s3_3ar s3_3as s3_3at s3_3au s3_3av s3_3aw s3_3ax s3_3ay s3_3az s3_3ba s3_3bb s3_3bc s3_3bd s3_3be s3_3bf)

tab delivery_care_equipments


replace delivery_care_equipments = 0 if delivery_care_equipments < 20
replace delivery_care_equipments = 1 if delivery_care_equipments >= 20

label values delivery_care_equipments  yesno





*Sick Newborn Care

foreach var in s3_14a s3_14b s3_14c s3_14d s3_14e s3_14f s3_14g s3_14h s3_14i s3_14j s3_14k s3_14l s3_14m s3_14n s3_14o s3_14p s3_14q s3_14r s3_14s s3_14t s3_14u s3_14v s3_14w s3_14x s3_14y s3_14z s3_14aa s3_14ab s3_14ac s3_14ad s3_14ae s3_14af s3_14ag s3_14ah s3_14ai s3_14aj s3_14ak s3_14al s3_14am s3_14an {
    replace `var' = . if `var' != 1
}

capture drop sick_newborn_care_equip
egen sick_newborn_care_equip  = rowtotal(s3_14a s3_14b s3_14c s3_14d s3_14e s3_14f s3_14g s3_14h s3_14i s3_14j s3_14k s3_14l s3_14m s3_14n s3_14o s3_14p s3_14q s3_14r s3_14s s3_14t s3_14u s3_14v s3_14w s3_14x s3_14y s3_14z s3_14aa s3_14ab s3_14ac s3_14ad s3_14ae s3_14af s3_14ag s3_14ah s3_14ai s3_14aj s3_14ak s3_14al s3_14am s3_14an)

tab sick_newborn_care_equip


replace sick_newborn_care_equip = 0 if sick_newborn_care_equip < 20
replace sick_newborn_care_equip = 1 if sick_newborn_care_equip >= 20

label values sick_newborn_care_equip  yesno





putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.2") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"

putexcel A2 = "Table 5.1.1.4.2. % of facilities that have adequate infrastructure/equipment/drug/supplies for providing each MNCH service", bold


putexcel B3 = "Delivery care", bold
putexcel C3 = "Sick Newborn Care", bold




///////////////////////////////////////////////////////////////////////////////
***Delivery care
///////////////////////////////////////////////////////////////////////////////


***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab delivery_care_equipments, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel C5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab delivery_care_equipments s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
 

// Percentages "Oui"
 putexcel B7= pct_dakar, nformat(number_d2)
 putexcel B8= pct_diourbel, nformat(number_d2)
 putexcel B9= pct_fatick, nformat(number_d2)
 putexcel B10= pct_kaffrine, nformat(number_d2)
 putexcel B11= pct_kaolack, nformat(number_d2)
 putexcel B12= pct_kedougou, nformat(number_d2)
 putexcel B13= pct_kolda, nformat(number_d2)
 putexcel B14= pct_louga, nformat(number_d2)
 putexcel B15= pct_matam, nformat(number_d2)
 putexcel B16= pct_stlouis, nformat(number_d2)
 putexcel B17= pct_sedhiou, nformat(number_d2)
 putexcel B18= pct_tambacounda, nformat(number_d2)
 putexcel B19= pct_thies, nformat(number_d2)
 putexcel B20= pct_ziguinchor, nformat(number_d2)
putexcel close




***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab delivery_care_equipments s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel B22 = pct_eps, nformat(number_d2)
putexcel B23 = pct_cs , nformat(number_d2)
putexcel B24 = pct_ps , nformat(number_d2)
putexcel B25 = pct_case , nformat(number_d2)
putexcel B26 = pct_clinic , nformat(number_d2)

putexcel close





///////////////////////////////////////////////////////////////////////////////
***Sick Newborn Care
///////////////////////////////////////////////////////////////////////////////

***&&&&& National &&&&&&&&&&&&&&&&&&&&&
tab sick_newborn_care_equip, matcell(freqs) matrow(names)
scalar total_responses = freqs[1,1] + freqs[2,1]
scalar count_oui = freqs[1,1]
scalar percent_oui = (count_oui / total_responses) * 100
display  count_oui
display percent_oui
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel C5 = percent_oui , nformat(number_d2)
putexcel close



***&&&&& Region &&&&&&&&&&&&&&&&&&&&&
tab sick_newborn_care_equip s1_1, matcell(freqs) matrow(names)
// Total per region
scalar  total_dakar = freqs[1,1] + freqs[2,1]
scalar  total_diourbel  = freqs[1,2] + freqs[2,2]
scalar  total_fatick    = freqs[1,3] + freqs[2,3]
scalar  total_kaffrine  = freqs[1,4] + freqs[2,4]
scalar  total_kaolack   = freqs[1,5] + freqs[2,5]
scalar  total_kedougou  = freqs[1,6] + freqs[2,6]
scalar  total_kolda     = freqs[1,7] + freqs[2,7]
scalar  total_louga     = freqs[1,8] + freqs[2,8]
scalar  total_matam        = freqs[1,9] + freqs[2,9]
scalar  total_stlouis      = freqs[1,10] + freqs[2,10]
scalar  total_sedhiou      = freqs[1,11] + freqs[2,11]
scalar  total_tambacounda  = freqs[1,12] + freqs[2,12]
scalar  total_thies        = freqs[1,13] + freqs[2,13]
scalar  total_ziguinchor   = freqs[1,14] + freqs[2,14]


// Percentages "Oui"
scalar  pct_dakar      = (freqs[1,1] /  total_dakar) * 100
scalar  pct_diourbel   = (freqs[1,2] / total_diourbel) * 100
scalar  pct_fatick     = (freqs[1,3] /  total_fatick) * 100
scalar  pct_kaffrine   = (freqs[1,4] /  total_kaffrine) * 100
scalar  pct_kaolack    = (freqs[1,5] /  total_kaolack) * 100
scalar  pct_kedougou   = (freqs[1,6] /  total_kedougou) * 100
scalar  pct_kolda      = (freqs[1,7] /  total_kolda) * 100
scalar  pct_louga      = (freqs[1,8] /  total_louga) * 100
scalar  pct_matam        = (freqs[1,9] /  total_matam) * 100
scalar  pct_stlouis      = (freqs[1,10] /  total_stlouis) * 100
scalar  pct_sedhiou      = (freqs[1,11] /  total_sedhiou) * 100
scalar  pct_tambacounda  = (freqs[1,12] /  total_tambacounda) * 100
scalar  pct_thies        = (freqs[1,13] /  total_thies) * 100
scalar  pct_ziguinchor   = (freqs[1,14] /  total_ziguinchor) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
 

// Percentages "Oui"
putexcel C7= pct_dakar, nformat(number_d2)
putexcel C8= pct_diourbel, nformat(number_d2)
putexcel C9= pct_fatick, nformat(number_d2)
putexcel C10= pct_kaffrine, nformat(number_d2)
putexcel C11= pct_kaolack, nformat(number_d2)
putexcel C12= pct_kedougou, nformat(number_d2)
putexcel C13= pct_kolda, nformat(number_d2)
putexcel C14= pct_louga, nformat(number_d2)
putexcel C15= pct_matam, nformat(number_d2)
putexcel C16= pct_stlouis, nformat(number_d2)
putexcel C17= pct_sedhiou, nformat(number_d2)
putexcel C18= pct_tambacounda, nformat(number_d2)
putexcel C19= pct_thies, nformat(number_d2)
putexcel C20= pct_ziguinchor, nformat(number_d2)
putexcel close


***&&&&& Type of the facility &&&&&&&&&&&&&&&&&&&&&
tab sick_newborn_care_equip s1_7, matcell(freqs) matrow(names)

scalar total_case   = freqs[1,1] + freqs[2,1]
scalar total_clinic    = freqs[1,2] + freqs[2,2]
scalar total_cs    = freqs[1,3] + freqs[2,3]
scalar total_eps  = freqs[1,4] + freqs[2,4]
scalar total_ps= freqs[1,5] + freqs[2,5]


scalar pct_case    = (freqs[1,1] / total_case) * 100
scalar pct_clinic     = (freqs[1,2] / total_clinic) * 100
scalar pct_cs     = (freqs[1,3] / total_cs) * 100
scalar pct_eps   = (freqs[1,4] / total_eps) * 100
scalar pct_ps = (freqs[1,5] / total_ps) * 100




putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.4.1") modify
putexcel C22 = pct_eps, nformat(number_d2)
putexcel C23 = pct_cs , nformat(number_d2)
putexcel C24 = pct_ps , nformat(number_d2)
putexcel C25 = pct_case , nformat(number_d2)
putexcel C26 = pct_clinic , nformat(number_d2)


putexcel close



























*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&88888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.1.1.5. % of facilities that have consistent availability of commodities, by FP method
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.1.1.5") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"
putexcel A27 = "FP methods", bold
putexcel A28 = "Male condom"
putexcel A29 = "Female condom"
putexcel A30 = "PCU" 
putexcel A31 = "Injectable-Depo Provera "
putexcel A32 = "Injectable - Sayana Press"
putexcel A33 = "Implants"
putexcel A34 = "PCO" 
putexcel A35 = "Progesterone-only pills" 
putexcel A36 = "IUD"
putexcel A37 = "Tubal rings"
putexcel A38 = "Pregnancy test kits"


putexcel A2 = "Table 5.1.1.5. % of facilities that have consistent availability of commodities, by FP method", bold


putexcel B3 = "Number of HF  that usually sell FP methods", bold
putexcel C3 = "Percent with any brand currently in stock", bold
putexcel D3 = "Percent of HF that had a stock in the last 3 months", bold































*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&88888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.2.1.2.1  % of facilities with sotck-outs in FP commodities, and the reasons					
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.2.1.2.1") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"
putexcel A27 = "FP methods", bold
putexcel A28 = "Male condom"
putexcel A29 = "Female condom"
putexcel A30 = "PCU" 
putexcel A31 = "Injectable-Depo Provera "
putexcel A32 = "Injectable - Sayana Press"
putexcel A33 = "Implants"
putexcel A34 = "PCO" 
putexcel A35 = "Progesterone-only pills" 
putexcel A36 = "IUD"
putexcel A37 = "Tubal rings"
putexcel A38 = "Pregnancy test kits"


putexcel A2 = "Table 5.2.1.2.1  % of facilities with sotck-outs in FP commodities, and the reasons", bold


putexcel B3 = "% HF with stock-outs in the last 3 months", bold
putexcel C3 = "No supply received", bold
putexcel D3 = "Budget constraints", bold
putexcel E3 = "Limited purchases options", bold
putexcel F3 = "Quality assurance issues", bold
putexcel G3 = "Other reasons", bold

*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&88888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.2.1.3.1. % of facilities that do not regularly provide FP services, by reasons				
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.2.1.3.1") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"
putexcel A27 = "FP methods", bold
putexcel A28 = "Male condom"
putexcel A29 = "Female condom"
putexcel A30 = "PCU" 
putexcel A31 = "Injectable-Depo Provera "
putexcel A32 = "Injectable - Sayana Press"
putexcel A33 = "Implants"
putexcel A34 = "PCO" 
putexcel A35 = "Progesterone-only pills" 
putexcel A36 = "IUD"
putexcel A37 = "Tubal rings"
putexcel A38 = "Pregnancy test kits"


putexcel A2 = "Table 5.2.1.3.1. % of facilities that do not regularly provide FP services, by reason", bold


putexcel B3 = "Number of facilities that do not regularly provide FP services", bold
putexcel C3 = "No trained provider", bold
putexcel D3 = "No furniture available", bold
putexcel E3 = "No infrastructure avalable", bold
putexcel F3 = "Clients don't want", bold
putexcel G3 = "Other reasons", bold







*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&88888888888888888888888888
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
***Table 5.2.1.3.2. % of facilities that do not regularly provide an MNCH service, by most commn reasons by MNCH components	
*&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
putexcel set "FP_Manitoba_HF.xlsx", sheet("T_5.2.1.3.2") modify
putexcel A5 = "National", bold
putexcel A6 = "Region", bold
putexcel A7 = "Dakar"
putexcel A8 = "Diourbel"
putexcel A9 = "Fatick"
putexcel A10 = "Kaffrine"
putexcel A11 = "Kaolack"
putexcel A12 = "Kedougou"
putexcel A13 = "Kolda"
putexcel A14 = "Louga"
putexcel A15 = "Matam"
putexcel A16 = "Saint-louis"
putexcel A17 = "Sedhiou"
putexcel A18 = "Tambacounda"
putexcel A19 = "Thies"
putexcel A20 = "Ziguinchor"
putexcel A21 = "Type of the facility", bold
putexcel A22 = "EPS"
putexcel A23 = "CS"
putexcel A24 = "PS"
putexcel A25 = "CASE"
putexcel A26 = "CLINIC"
putexcel A27 = "FP methods", bold
putexcel A28 = "Male condom"
putexcel A29 = "Female condom"
putexcel A30 = "PCU" 
putexcel A31 = "Injectable-Depo Provera "
putexcel A32 = "Injectable - Sayana Press"
putexcel A33 = "Implants"
putexcel A34 = "PCO" 
putexcel A35 = "Progesterone-only pills" 
putexcel A36 = "IUD"
putexcel A37 = "Tubal rings"
putexcel A38 = "Pregnancy test kits"

putexcel A2 = "Table 5.2.1.3.2. % of facilities that do not regularly provide an MNCH service, by most commn reasons by MNCH components", bold

putexcel B3 = "Number of facilities that do not regularly provide FP services", bold
putexcel C3 = "No trained provider", bold
putexcel D3 = "No furniture available", bold
putexcel E3 = "No infrastructure avalable", bold
putexcel F3 = "Clients don't want", bold
putexcel G3 = "Other reasons", bold