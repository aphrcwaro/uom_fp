* import_clinique_corr-section_41-personnel.do
*
* 	Imports and aggregates "QUESTIONNAIRE CLINIQUE PRIVEE-section_41-personnel" (ID: clinique_corr) data.
*
*	Inputs:  "F:/Dropbox_APHRC/Dropbox/FP-Project-APHRC-UM/Atelier_FP/Data/raw/QUESTIONNAIRE CLINIQUE PRIVEE-section_41-personnel.csv"
*	Outputs: "F:/Dropbox_APHRC/Dropbox/FP-Project-APHRC-UM/Atelier_FP/Data/raw/QUESTIONNAIRE CLINIQUE PRIVEE-section_41-personnel.dta"
*
*	Output by SurveyCTO March 26, 2025 5:39 PM.

* initialize Stata
clear all
set more off
set mem 100m

* initialize workflow-specific parameters
*	Set overwrite_old_data to 1 if you use the review and correction
*	workflow and allow un-approving of submissions. If you do this,
*	incoming data will overwrite old data, so you won't want to make
*	changes to data in your local .dta file (such changes can be
*	overwritten with each new import).
local overwrite_old_data 0

* initialize form-specific parameters
local csvfile "F:/Dropbox_APHRC/Dropbox/FP-Project-APHRC-UM/Atelier_FP/Data/raw/QUESTIONNAIRE CLINIQUE PRIVEE-section_41-personnel.csv"
local dtafile "F:/Dropbox_APHRC/Dropbox/FP-Project-APHRC-UM/Atelier_FP/Data/raw/clinic-section_41-personnel_raw.dta"
local corrfile "F:/Dropbox_APHRC/Dropbox/FP-Project-APHRC-UM/Atelier_FP/Data/raw/QUESTIONNAIRE CLINIQUE PRIVEE-section_41-personnel_corrections.csv"
local note_fields1 ""
local text_fields1 "numerob_produitf codeb_produitf produitb_labelg s4_1_8 s4_1_11 s4_1_12 s4_1_12au"
local date_fields1 ""
local datetime_fields1 ""

disp
disp "Starting import of: `csvfile'"
disp

* import data from primary .csv file
insheet using "`csvfile'", names clear

* drop extra table-list columns
cap drop reserved_name_for_field_*
cap drop generated_table_list_lab*

* continue only if there's at least one row of data to import
if _N>0 {
	* drop note fields (since they don't contain any real data)
	forvalues i = 1/100 {
		if "`note_fields`i''" ~= "" {
			drop `note_fields`i''
		}
	}
	
	* format date and date/time fields
	forvalues i = 1/100 {
		if "`datetime_fields`i''" ~= "" {
			foreach dtvarlist in `datetime_fields`i'' {
				cap unab dtvarlist : `dtvarlist'
				if _rc==0 {
					foreach dtvar in `dtvarlist' {
						tempvar tempdtvar
						rename `dtvar' `tempdtvar'
						gen double `dtvar'=.
						cap replace `dtvar'=clock(`tempdtvar',"MDYhms",2025)
						* automatically try without seconds, just in case
						cap replace `dtvar'=clock(`tempdtvar',"MDYhm",2025) if `dtvar'==. & `tempdtvar'~=""
						format %tc `dtvar'
						drop `tempdtvar'
					}
				}
			}
		}
		if "`date_fields`i''" ~= "" {
			foreach dtvarlist in `date_fields`i'' {
				cap unab dtvarlist : `dtvarlist'
				if _rc==0 {
					foreach dtvar in `dtvarlist' {
						tempvar tempdtvar
						rename `dtvar' `tempdtvar'
						gen double `dtvar'=.
						cap replace `dtvar'=date(`tempdtvar',"MDY",2025)
						format %td `dtvar'
						drop `tempdtvar'
					}
				}
			}
		}
	}

	* ensure that text fields are always imported as strings (with "" for missing values)
	* (note that we treat "calculate" fields as text; you can destring later if you wish)
	tempvar ismissingvar
	quietly: gen `ismissingvar'=.
	forvalues i = 1/100 {
		if "`text_fields`i''" ~= "" {
			foreach svarlist in `text_fields`i'' {
				cap unab svarlist : `svarlist'
				if _rc==0 {
					foreach stringvar in `svarlist' {
						quietly: replace `ismissingvar'=.
						quietly: cap replace `ismissingvar'=1 if `stringvar'==.
						cap tostring `stringvar', format(%100.0g) replace
						cap replace `stringvar'="" if `ismissingvar'==1
					}
				}
			}
		}
	}
	quietly: drop `ismissingvar'



	* label variables
	label variable key "Unique submission ID"
	cap label variable submissiondate "Date/time submitted"
	cap label variable formdef_version "Form version used on device"
	cap label variable review_status "Review status"
	cap label variable review_comments "Comments made during review"
	cap label variable review_corrections "Corrections made during review"


	label variable s4_1_3_0 "Combien de \${produitb_labelg} sont autorisés dans le service de gynécologie?"
	note s4_1_3_0: "Combien de \${produitb_labelg} sont autorisés dans le service de gynécologie?"

	label variable s4_1_3_1 "Combien de postes de \${produitb_labelg} sont actuellement vacants ?"
	note s4_1_3_1: "Combien de postes de \${produitb_labelg} sont actuellement vacants ?"

	label variable s4_1_4a "Combien de \${produitb_labelg} sont de sexe masculin ?"
	note s4_1_4a: "Combien de \${produitb_labelg} sont de sexe masculin ?"

	label variable s4_1_4b "Combien de \${produitb_labelg} sont de sexe féminin ?"
	note s4_1_4b: "Combien de \${produitb_labelg} sont de sexe féminin ?"

	label variable s4_1_6a "Combien de \${produitb_labelg} femmes ont reçu une formation complémentaire sur "
	note s4_1_6a: "Combien de \${produitb_labelg} femmes ont reçu une formation complémentaire sur le PF ?"

	label variable s4_1_6b "Combien de \${produitb_labelg} hommes ont reçu une formation complémentaire sur "
	note s4_1_6b: "Combien de \${produitb_labelg} hommes ont reçu une formation complémentaire sur le PF ?"

	label variable s4_1_7a "Combien de \${produitb_labelg} femmes fournissent actuellement des services de P"
	note s4_1_7a: "Combien de \${produitb_labelg} femmes fournissent actuellement des services de PF ?"

	label variable s4_1_7b "Combien de \${produitb_labelg} hommes fournissent actuellement des services de P"
	note s4_1_7b: "Combien de \${produitb_labelg} hommes fournissent actuellement des services de PF ?"

	label variable s4_1_8 "Quelles sont les méthodes que propose les \${produitb_labelg} dans la structure "
	note s4_1_8: "Quelles sont les méthodes que propose les \${produitb_labelg} dans la structure ?"

	label variable s4_1_9a "Combien de \${produitb_labelg} femmes ont reçu une formation supplémentaire reçu"
	note s4_1_9a: "Combien de \${produitb_labelg} femmes ont reçu une formation supplémentaire reçue sur la SMNI ?"

	label variable s4_1_9b "Combien de \${produitb_labelg} hommes ont reçu une formation supplémentaire reçu"
	note s4_1_9b: "Combien de \${produitb_labelg} hommes ont reçu une formation supplémentaire reçue sur la SMNI ?"

	label variable s4_1_10a "Combien \${produitb_labelg} femmes fournissent actuellement des services de SMNI"
	note s4_1_10a: "Combien \${produitb_labelg} femmes fournissent actuellement des services de SMNI ?"

	label variable s4_1_10b "Combien \${produitb_labelg} hommes fournissent actuellement des services de SMNI"
	note s4_1_10b: "Combien \${produitb_labelg} hommes fournissent actuellement des services de SMNI ?"

	label variable s4_1_11 "Quels sont les services de SMNI fournit par les services de SMNI \${produitb_lab"
	note s4_1_11: "Quels sont les services de SMNI fournit par les services de SMNI \${produitb_labelg} ?"

	label variable s4_1_11_2 "Combien de \${produitb_labelg} femmes ont reçu une formation sur l'échographie ?"
	note s4_1_11_2: "Combien de \${produitb_labelg} femmes ont reçu une formation sur l'échographie ?"

	label variable s4_1_11_3 "Combien de \${produitb_labelg} hommes ont reçu une formation sur l'échographie ?"
	note s4_1_11_3: "Combien de \${produitb_labelg} hommes ont reçu une formation sur l'échographie ?"

	label variable s4_1_12 "Pourquoi le poste de \${produitb_labelg} est-il actuellement vacant ?"
	note s4_1_12: "Pourquoi le poste de \${produitb_labelg} est-il actuellement vacant ?"

	label variable s4_1_12au "Veuillez préciser le motif"
	note s4_1_12au: "Veuillez préciser le motif"

	label variable s4_1_13 "En moyenne, depuis combien de temps le(les) poste(s) de \${produitb_labelg} est("
	note s4_1_13: "En moyenne, depuis combien de temps le(les) poste(s) de \${produitb_labelg} est(sont)-il(s) vacant ?"






	* append old, previously-imported data (if any)
	cap confirm file "`dtafile'"
	if _rc == 0 {
		* mark all new data before merging with old data
		gen new_data_row=1
		
		* pull in old data
		append using "`dtafile'"
		
		* drop duplicates in favor of old, previously-imported data if overwrite_old_data is 0
		* (alternatively drop in favor of new data if overwrite_old_data is 1)
		sort key
		by key: gen num_for_key = _N
		drop if num_for_key > 1 & ((`overwrite_old_data' == 0 & new_data_row == 1) | (`overwrite_old_data' == 1 & new_data_row ~= 1))
		drop num_for_key

		* drop new-data flag
		drop new_data_row
	}
	
	* save data to Stata format
	save "`dtafile'", replace

	* show codebook and notes
	codebook
	notes list
}

disp
disp "Finished import of: `csvfile'"
disp

* OPTIONAL: LOCALLY-APPLIED STATA CORRECTIONS
*
* Rather than using SurveyCTO's review and correction workflow, the code below can apply a list of corrections
* listed in a local .csv file. Feel free to use, ignore, or delete this code.
*
*   Corrections file path and filename:  F:/Dropbox_APHRC/Dropbox/FP-Project-APHRC-UM/Atelier_FP/Data/raw/QUESTIONNAIRE CLINIQUE PRIVEE-section_41-personnel_corrections.csv
*
*   Corrections file columns (in order): key, fieldname, value, notes

capture confirm file "`corrfile'"
if _rc==0 {
	disp
	disp "Starting application of corrections in: `corrfile'"
	disp

	* save primary data in memory
	preserve

	* load corrections
	insheet using "`corrfile'", names clear
	
	if _N>0 {
		* number all rows (with +1 offset so that it matches row numbers in Excel)
		gen rownum=_n+1
		
		* drop notes field (for information only)
		drop notes
		
		* make sure that all values are in string format to start
		gen origvalue=value
		tostring value, format(%100.0g) replace
		cap replace value="" if origvalue==.
		drop origvalue
		replace value=trim(value)
		
		* correct field names to match Stata field names (lowercase, drop -'s and .'s)
		replace fieldname=lower(subinstr(subinstr(fieldname,"-","",.),".","",.))
		
		* format date and date/time fields (taking account of possible wildcards for repeat groups)
		forvalues i = 1/100 {
			if "`datetime_fields`i''" ~= "" {
				foreach dtvar in `datetime_fields`i'' {
					* skip fields that aren't yet in the data
					cap unab dtvarignore : `dtvar'
					if _rc==0 {
						gen origvalue=value
						replace value=string(clock(value,"MDYhms",2025),"%25.0g") if strmatch(fieldname,"`dtvar'")
						* allow for cases where seconds haven't been specified
						replace value=string(clock(origvalue,"MDYhm",2025),"%25.0g") if strmatch(fieldname,"`dtvar'") & value=="." & origvalue~="."
						drop origvalue
					}
				}
			}
			if "`date_fields`i''" ~= "" {
				foreach dtvar in `date_fields`i'' {
					* skip fields that aren't yet in the data
					cap unab dtvarignore : `dtvar'
					if _rc==0 {
						replace value=string(clock(value,"MDY",2025),"%25.0g") if strmatch(fieldname,"`dtvar'")
					}
				}
			}
		}

		* write out a temp file with the commands necessary to apply each correction
		tempfile tempdo
		file open dofile using "`tempdo'", write replace
		local N = _N
		forvalues i = 1/`N' {
			local fieldnameval=fieldname[`i']
			local valueval=value[`i']
			local keyval=key[`i']
			local rownumval=rownum[`i']
			file write dofile `"cap replace `fieldnameval'="`valueval'" if key=="`keyval'""' _n
			file write dofile `"if _rc ~= 0 {"' _n
			if "`valueval'" == "" {
				file write dofile _tab `"cap replace `fieldnameval'=. if key=="`keyval'""' _n
			}
			else {
				file write dofile _tab `"cap replace `fieldnameval'=`valueval' if key=="`keyval'""' _n
			}
			file write dofile _tab `"if _rc ~= 0 {"' _n
			file write dofile _tab _tab `"disp"' _n
			file write dofile _tab _tab `"disp "CAN'T APPLY CORRECTION IN ROW #`rownumval'""' _n
			file write dofile _tab _tab `"disp"' _n
			file write dofile _tab `"}"' _n
			file write dofile `"}"' _n
		}
		file close dofile
	
		* restore primary data
		restore
		
		* execute the .do file to actually apply all corrections
		do "`tempdo'"

		* re-save data
		save "`dtafile'", replace
	}
	else {
		* restore primary data		
		restore
	}

	disp
	disp "Finished applying corrections in: `corrfile'"
	disp
}
