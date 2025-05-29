cap cd "D:\Dropbox\Dropbox\FP-Project-APHRC-UM\Worshop_FP_UM\Health_facility_tables"
cap cd "F:\Dropbox_APHRC\Dropbox\FP-Project-APHRC-UM\Worshop_FP_UM\Health_facility_tables"

martin

use all_hf_data_long_clean_merged_2705.dta, clear
preserve
duplicates report id 
duplicates list id 
duplicates tag id , generate(duptag)
browse if duptag > 0
drop s4_*
duplicates drop id, force
save health_facility_s4_dropped, replace
restore
use health_facility_s4_dropped.dta, clear

tab s1_7,m
tab s1_1, m

************* Infrastructures généraux **************************

//ssc install tabout
tabout s3_1a s3_1b s3_1c s3_1d s3_1e s3_1f s3_1g s3_1h s3_1i s3_1j s3_1k s3_1l s3_1m s3_1n  using "table_carract.xls", append c(col) npos(both) nlab(Number) 

************* Fourniture globale /overall supply **********************************

tabout s5_6 s1_7  using "table_fourniture_general.xls", append c(col) npos(both) nlab(Number) 

************* Fourniture des services de PF /Provision of FP services**************

cap drop pillule_fourni
gen pillule_fourni = .
replace pillule_fourni = 1 if s5_7a == 1 | s5_7a == 2 | s5_7a == 3 | s5_7a == 4
replace pillule_fourni = 0 if s5_7a == 5
tab pillule_fourni, m


cap drop injectable_fourni
gen injectable_fourni = .
replace injectable_fourni = 1 if s5_7b == 1 | s5_7b == 2 | s5_7b == 3 | s5_7b == 4
replace injectable_fourni = 0 if s5_7b == 5
tab injectable_fourni, m


cap drop condom_male_fourni
gen condom_male_fourni = .
replace condom_male_fourni = 1 if s5_7c == 1 | s5_7c == 2 | s5_7c == 3 | s5_7c == 4
replace condom_male_fourni = 0 if s5_7c == 5
tab condom_male_fourni, m

cap drop condom_female_fourni
gen condom_female_fourni = .
replace condom_female_fourni = 1 if s5_7d == 1 | s5_7d == 2 | s5_7d == 3 | s5_7d == 4
replace condom_female_fourni = 0 if s5_7d == 5
tab condom_female_fourni, m


cap drop pcu_fourni
gen pcu_fourni = .
replace pcu_fourni = 1 if s5_7e == 1 | s5_7e == 2 | s5_7e == 3 | s5_7e == 4
replace pcu_fourni = 0 if s5_7e == 5
tab pcu_fourni, m


cap drp dui_fourni
gen dui_fourni = .
replace dui_fourni = 1 if s5_7f == 1 | s5_7f == 2 | s5_7f == 3 | s5_7f == 4
replace dui_fourni = 0 if s5_7f == 5
tab dui_fourni, m

cap drop implant_fourni
gen implant_fourni = .
replace implant_fourni = 1 if s5_7g == 1 | s5_7g == 2 | s5_7g == 3 | s5_7g == 4
replace implant_fourni = 0 if s5_7g == 5
tab implant_fourni, m


cap drop sterilisation_f_fourni
gen sterilisation_f_fourni = .
replace sterilisation_f_fourni = 1 if s5_7h == 1 | s5_7h == 2 | s5_7h == 3 | s5_7h == 4
replace sterilisation_f_fourni = 0 if s5_7h == 5
tab sterilisation_f_fourni, m


cap drop sterilisation_m_fourni
gen sterilisation_m_fourni = .
replace sterilisation_m_fourni = 1 if s5_7i == 1 | s5_7i == 2 | s5_7i == 3 | s5_7i == 4
replace sterilisation_m_fourni = 0 if s5_7i == 5
tab sterilisation_m_fourni, m


cap drop mama_fourni
gen mama_fourni = .
replace mama_fourni = 1 if s5_7j == 1 | s5_7j == 2 | s5_7j == 3 | s5_7j == 4
replace mama_fourni = 0 if s5_7j == 5
tab mama_fourni, m


cap drop mjf_fourni
gen mjf_fourni = .
replace mjf_fourni = 1 if s5_7k == 1 | s5_7k == 2 | s5_7k == 3 | s5_7k == 4
replace mjf_fourni = 0 if s5_7k == 5
tab mjf_fourni, m



label val pillule_fourni injectable_fourni injectable_fourni condom_male_fourni condom_female_fourni pcu_fourni dui_fourni implant_fourni sterilisation_f_fourni sterilisation_m_fourni mama_fourni mjf_fourni yesno

tabout pillule_fourni injectable_fourni condom_male_fourni condom_female_fourni pcu_fourni dui_fourni implant_fourni sterilisation_f_fourni sterilisation_m_fourni mama_fourni mjf_fourni s1_7  using "table_fourniture.xls", append c(col) npos(both) nlab(Number) 

*Short-acting methods include pills, injectables, or the emergency contraceptive pill
*Long-acting and permanent methods include implants or IUDs

egen short_acting = rowtotal(pillule_fourni injectable_fourni pcu_fourni)

egen long_acting = rowtotal(implant_fourni dui_fourni sterilisation_f_fourni sterilisation_m_fourni)

gen short_acting_3 = .
replace short_acting_3 = 1 if short_acting == 3
replace short_acting_3 = 0 if short_acting != 3 & s5_6 ==1
label val  short_acting_3 yesno

gen long_acting_1 = .
replace long_acting_1 = 1 if long_acting  >= 1
replace long_acting_1 = 0 if long_acting == 0 & s5_6 ==1 
label values long_acting_1 yesno

tabout short_acting_3 long_acting_1 s1_7  using "table_fourniture.xls", append c(col) npos(both) nlab(Number) 




********************* Cascade availability **************************
/*Pilules=A, Injectables= B, Préservatif masculin =C, Préservatif féminin= D, Contraception d'urgence=E, DIU= F, Implants=G, Stérilisation féminine (Ligature des trompes)H, Stérilisation masculine/ Vasectomie= J, Allaitement maternel exclusif (MAMA)= K, Méthode des jours fixes (MJF)= L */

********************* DUI 
cap drop dui_daily
gen dui_daily = 0
replace dui_daily = 1  if s5_7f == 1 
replace dui_daily = .  if s5_6 == 2
label values dui_daily yesno

cap drop dui_daily_trained
gen dui_daily_trained = 0
replace dui_daily_trained = 1 if (dui_daily == 1) //& (( s4_1_5a_1!=0 & s4_1_5a_1!= 99& s4_1_5a_1!= 999)|( s4_1_5b_1!=0 & s4_1_5b_1!= 99& s4_1_5b_1!= 999))
replace dui_daily_trained = .  if s5_6 == 2
label values dui_daily_trained yesno


gen dui_daily_trained_stock = 0
replace dui_daily_trained_stock  = 1 if dui_daily_trained == 1 &  s6_16i == 1
replace dui_daily_trained_stock = . if s5_6 == 2
label values dui_daily_trained_stock yesno

 
gen dui_daily_trained_stock_items = 0
replace dui_daily_trained_stock_items = 1 if ((s6_3trav_a==1 & s6_3trav_b==1 &  s6_3trav_c==1 &  s6_3trav_d ==1 & s6_3trav_e==1 &  s6_3trav_f==1 &  s6_3trav_g==1 &  s6_3trav_h==1 &  s6_3trav_i ==1 & s6_3trav_j ==1 & s6_3trav_k==1 &  s6_3trav_l==1 &  s6_3trav_m==1 &  s6_3trav_n==1)|(s6_3coin_a ==1 &  s6_3coin_b ==1 & s6_3coin_c==1 &  s6_3coin_d==1 &  s6_3coin_e==1 &  s6_3coin_f ==1 & s6_3coin_g ==1 & s6_3coin_h==1 &  s6_3coin_i==1 &  s6_3coin_j ==1 & s6_3coin_k==1 &  s6_3coin_l ==1 & s6_3coin_m ==1 & s6_3coin_n==1)|(s6_3both_a ==1 & s6_3both_b ==1 & s6_3both_c==1 &  s6_3both_d==1 &  s6_3both_e==1 &  s6_3both_f==1 &  s6_3both_g ==1 & s6_3both_h ==1 & s6_3both_i==1 &  s6_3both_j ==1 & s6_3both_k==1 &  s6_3both_l==1 &  s6_3both_m==1 &  s6_3both_n==1)|(s6_3other_a==1 &  s6_3other_b==1 &  s6_3other_c==1 &  s6_3other_d==1 &  s6_3other_e==1 &  s6_3other_f==1 &  s6_3other_g==1 &  s6_3other_h==1 &  s6_3other_i ==1 & s6_3other_j==1 &  s6_3other_k==1 &  s6_3other_l==1 &  s6_3other_m==1 &  s6_3other_n==1)) & dui_daily_trained_stock == 1

replace dui_daily_trained_stock_items = . if s5_6 == 2

label values dui_daily_trained_stock_items yesno


cap drop dui_daily_trained_stock_items_p
gen dui_daily_trained_stock_items_p = 0
replace dui_daily_trained_stock_items_p = 1 if (dui_daily_trained_stock_items == 1)&((strpos(s8_5_A, "DIU")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "DIU")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "DIU")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "DIU")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "DIU")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "DIU")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "DIU")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "DIU")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "DIU")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "DIU")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "DIU")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))
replace dui_daily_trained_stock_items_p = . if s5_6 == 2

label values dui_daily_trained_stock_items_p yesno


********************* Injectables


gen inject_daily = 0
replace inject_daily = 1 if s5_7b == 1 
replace inject_daily = . if s5_6 == 2
label values inject_daily yesno



gen inject_daily_stock = 0
replace inject_daily_stock  = 1 if (inject_daily == 1) & ( s6_16d== 1 |  s6_16e==1)
replace inject_daily_stock  = . if s5_6 == 2 
label values inject_daily_stock yesno

cap drop inject_daily_stock_prov
gen inject_daily_stock_prov = 0
replace inject_daily_stock_prov  = 1 if (inject_daily_stock == 1)&((strpos(s8_5_A, "Injectable")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Injectable")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Injectable")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Injectable")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Injectable")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Injectable")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Injectable")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Injectable")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Injectable")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Injectable")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Injectable")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))
replace inject_daily_stock_prov  = . if s5_6 == 2
label values inject_daily_stock_prov yesno


tabout injectable_fourni inject_daily inject_daily_stock inject_daily_stock_prov s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number)



********************* Implants 

cap drop implant_daily
gen implant_daily = 0
replace implant_daily = 1 if s5_7g == 1
replace implant_daily = . if s5_6 == 2
label values implant_daily  yesno


cap drop implant_daily_trained
gen implant_daily_trained = 0
replace implant_daily_trained = 1 if (implant_daily == 1) //& (( s4_1_5a_1!=0 & s4_1_5a_1!= 99& s4_1_5a_1!= 999)|( s4_1_5b_1!=0 & s4_1_5b_1!= 99& s4_1_5b_1!= 999))
replace implant_daily_trained= .  if s5_6 == 2
label values implant_daily_trained yesno


gen     implant_daily_trained_stock = 0
replace implant_daily_trained_stock  = 1 if implant_daily_trained == 1 &   s6_16f == 1
replace implant_daily_trained_stock = . if s5_6 == 2
label values implant_daily_trained_stock yesno



gen implant_daily_stock_prov = 0
replace implant_daily_stock_prov  = 1 if (implant_daily_trained_stock == 1)&((strpos(s8_5_A, "Implant")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Implant")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Implant")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Implant")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Implant")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Implant")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Implant")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Implant")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Implant")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Implant")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Implant")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))

replace implant_daily_stock_prov = . if s5_6 == 2
label values implant_daily_stock_prov yesno


tabout implant_fourni implant_daily implant_daily_trained implant_daily_trained_stock implant_daily_stock_prov s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number)



*************  Condoms

* Male condoms

gen male_condoms_daily = 0
replace male_condoms_daily = 1 if  s5_7c == 1
replace male_condoms_daily = . if s5_6 == 2
label values male_condoms_daily  yesno



gen male_condoms_daily_stock = 0
replace male_condoms_daily_stock = 1 if (male_condoms_daily == 1) & (s6_16a == 1)
replace male_condoms_daily_stock  = . if s5_6 == 2 
label values male_condoms_daily_stock yesno






gen male_condoms_daily_stock_prov = 0
replace male_condoms_daily_stock_prov  = 1 if (male_condoms_daily_stock == 1)&((strpos(s8_5_A, "Préservatif masculin")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Préservatif masculin")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Préservatif masculin")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Préservatif masculin")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Préservatif masculin")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Préservatif masculin")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Préservatif masculin")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Préservatif masculin")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Préservatif masculin")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Préservatif masculin")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Préservatif masculin")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))

replace male_condoms_daily_stock_prov = . if s5_6 == 2
label values male_condoms_daily_stock_prov yesno


tabout condom_male_fourni male_condoms_daily male_condoms_daily_stock male_condoms_daily_stock_prov s1_7 using "table_cascade.xls", append c(col) npos(both) nlab(Number)




* Femaleale condoms
gen female_condoms_daily = 0
replace female_condoms_daily = 1 if  s5_7d == 1
replace female_condoms_daily = . if s5_6 == 2
label values female_condoms_daily  yesno



gen female_condoms_daily_stock = 0
replace female_condoms_daily_stock = 1 if (female_condoms_daily == 1) & ( s6_16b == 1)
replace female_condoms_daily_stock  = . if s5_6 == 2 
label values female_condoms_daily_stock yesno



gen female_condoms_daily_stock_prov = 0
replace female_condoms_daily_stock_prov  = 1 if (female_condoms_daily_stock == 1)&((strpos(s8_5_A, "Préservatif féminin")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Préservatif féminin")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Préservatif féminin")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Préservatif féminin")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Préservatif féminin")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Préservatif féminin")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Préservatif féminin")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Préservatif féminin")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Préservatif féminin")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Préservatif féminin")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Préservatif féminin")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))

replace female_condoms_daily_stock_prov = . if s5_6 == 2
label values female_condoms_daily_stock_prov yesno


tabout condom_female_fourni female_condoms_daily female_condoms_daily_stock female_condoms_daily_stock_prov s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number)



******** Pills


gen pills_daily = 0
replace pills_daily = 1 if  s5_7a == 1
replace pills_daily = . if s5_6 == 2
label values pills_daily  yesno



gen pills_daily_stock = 0
replace pills_daily_stock = 1 if (pills_daily == 1) & ( s6_16g == 1 | s6_16h ==1 )
replace pills_daily_stock  = . if s5_6 == 2 
label values pills_daily_stock yesno


cap drop pills_daily_stock_prov
gen pills_daily_stock_prov = 0
replace pills_daily_stock_prov  = 1 if (pills_daily_stock == 1)&((strpos(s8_5_A, "Pilule")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Pilule")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Pilule")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Pilule")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Pilule")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Pilule")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Pilule")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Pilule")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Pilule")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Pilule")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Pilule")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))

replace pills_daily_stock_prov = . if s5_6 == 2
label values pills_daily_stock_prov yesno

tabout pillule_fourni pills_daily pills_daily_stock pills_daily_stock_prov s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number)



******** Emergency pills


gen pcu_daily = 0
replace pcu_daily = 1 if  s5_7e  == 1
replace pcu_daily = . if s5_6 == 2
label values pcu_daily  yesno



gen pcu_daily_stock = 0
replace pcu_daily_stock = 1 if (pcu_daily == 1) & ( s6_16c == 1)
replace pcu_daily_stock  = . if s5_6 == 2 
label values pcu_daily_stock yesno



gen pcu_daily_stock_prov = 0
replace pcu_daily_stock_prov  = 1 if (pcu_daily_stock == 1)&((strpos(s8_5_A, "Contraception d'urgence")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Contraception d'urgence")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Contraception d'urgence")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Contraception d'urgence")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Contraception d'urgence")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Contraception d'urgence")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Contraception d'urgence")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Contraception d'urgence")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Contraception d'urgence")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Contraception d'urgence")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Contraception d'urgence")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))

replace pcu_daily_stock_prov = . if s5_6 == 2
label values pcu_daily_stock_prov yesno

tabout pcu_fourni pcu_daily pcu_daily_stock pcu_daily_stock_prov  s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number)



**************** Male seterilisation




gen male_ster_daily = 0
replace male_ster_daily = 1 if   s5_7i == 1
replace male_ster_daily = . if s5_6 == 2
label values male_ster_daily  yesno



gen male_ster_daily_stock = 0
replace male_ster_daily_stock = 1 if (male_ster_daily == 1) & (s6_15>=1)
replace male_ster_daily_stock  = . if s5_6 == 2 
label values male_ster_daily_stock yesno



gen male_ster_daily_stock_item = 0
replace male_ster_daily_stock_item = 1 if (male_ster_daily_stock == 1) & ( s6_14a==1 & s6_14b==1 & s6_14c==1 &  s6_14d==1 &  s6_14e==1 &  s6_14f==1 &  s6_14g==1 &  s6_14h==1 &  s6_14i==1)
replace male_ster_daily_stock_item  = . if s5_6 == 2 
label values male_ster_daily_stock_item  yesno




gen male_ster_daily_stock_item_prov=0
replace male_ster_daily_stock_item_prov  = 1 if (male_ster_daily_stock_item == 1)&((strpos(s8_5_A, "Stérilisation masculine/ Vasectomie")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Stérilisation masculine/ Vasectomie")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Stérilisation masculine/ Vasectomie")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Stérilisation masculine/ Vasectomie")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Stérilisation masculine/ Vasectomie")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Stérilisation masculine/ Vasectomie")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Stérilisation masculine/ Vasectomie")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Stérilisation masculine/ Vasectomie")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Stérilisation masculine/ Vasectomie")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Stérilisation masculine/ Vasectomie")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Stérilisation masculine/ Vasectomie")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))
replace male_ster_daily_stock_item_prov = . if s5_6 == 2 
label values male_ster_daily_stock_item_prov yesno


tabout sterilisation_m_fourni male_ster_daily male_ster_daily_stock male_ster_daily_stock_item male_ster_daily_stock_item_prov  s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number)



**************** female seterilisation




gen female_ster_daily = 0
replace female_ster_daily = 1 if    s5_7h == 1
replace female_ster_daily = . if s5_6 == 2
label values female_ster_daily  yesno



gen female_ster_daily_stock = 0
replace female_ster_daily_stock = 1 if (female_ster_daily == 1) & ((s6_9!=0 & s6_9!=9999)|(s6_12!=0 & s6_12!=9999 & s6_12!=998 & s6_12!=999))
replace female_ster_daily_stock  = . if s5_6 == 2 
label values female_ster_daily_stock yesno



gen female_ster_daily_stock_item = 0
replace female_ster_daily_stock_item = 1 if (female_ster_daily_stock == 1) & ((s6_11a==1 & s6_11b==1 & s6_11c==1 & s6_11d==1 & s6_11e==1 & s6_11f==1 & s6_11g==1 & s6_11h==1 & s6_11i==1 & s6_11j==1 & s6_11k==1 & s6_11l ==1 & s6_11m==1 & s6_11n==1 & s6_11o==1 & s6_11p==1 & s6_11q==1 & s6_11r==1 & s6_11s==1 & s6_11t==1 & s6_11u==1 )|( s6_8a==1 & s6_8b==1 & s6_8c==1 & s6_8d==1 & s6_8e==1 & s6_8f==1 & s6_8g==1 & s6_8h==1 & s6_8i==1 & s6_8j==1 & s6_8k==1 & s6_8l==1 & s6_8m==1 & s6_8n==1 & s6_8o==1 & s6_8p==1 & s6_8q==1 & s6_8r==1 & s6_8s==1 & s6_8t==1 & s6_8u==1 & s6_8v==1 & s6_8w==1 & s6_8x==1 & s6_8y==1 & s6_8z==1))
replace female_ster_daily_stock_item  = . if s5_6 == 2 
label values female_ster_daily_stock_item  yesno


gen female_ster_daily_stock_item_p=0
replace female_ster_daily_stock_item_p  = 1 if (female_ster_daily_stock_item == 1)&((strpos(s8_5_A, "Stérilisation féminine (Ligature des")&(s8_3_A!=0)&(s8_3_A!=998)&(s8_3_A!=999)&(s8_3_A!=9998)&(s8_3_A!=.)&(s8_3_A!=9998)&(s8_3_A!=.))|(strpos(s8_5_B, "Stérilisation féminine (Ligature des")&(s8_3_B!=0)&(s8_3_B!=998)&(s8_3_B!=999)&(s8_3_B!=9998)&(s8_3_B!=.)&(s8_3_B!=9998)&(s8_3_B!=.))|(strpos(s8_5_C, "Stérilisation féminine (Ligature des")&(s8_3_C!=0)&(s8_3_C!=998)&(s8_3_C!=999)&(s8_3_C!=9998)&(s8_3_C!=.)&(s8_3_C!=9998)&(s8_3_C!=.))|(strpos(s8_5_D, "Stérilisation féminine (Ligature des")&(s8_3_D!=0)&(s8_3_D!=998)&(s8_3_D!=999)&(s8_3_D!=9998)&(s8_3_D!=.)&(s8_3_D!=9998)&(s8_3_D!=.))|(strpos(s8_5_E, "Stérilisation féminine (Ligature des")&(s8_3_E!=0)&(s8_3_E!=998)&(s8_3_E!=999)&(s8_3_E!=9998)&(s8_3_E!=.)&(s8_3_E!=9998)&(s8_3_E!=.))|(strpos(s8_5_F, "Stérilisation féminine (Ligature des")&(s8_3_F!=0)&(s8_3_F!=998)&(s8_3_F!=999)&(s8_3_F!=9998)&(s8_3_F!=.)&(s8_3_F!=9998)&(s8_3_F!=.))|(strpos(s8_5_G, "Stérilisation féminine (Ligature des")&(s8_3_G!=0)&(s8_3_G!=998)&(s8_3_G!=999)&(s8_3_G!=9998)&(s8_3_G!=.)&(s8_3_G!=9998)&(s8_3_G!=.))|(strpos(s8_5_H, "Stérilisation féminine (Ligature des")&(s8_3_H!=0)&(s8_3_H!=998)&(s8_3_H!=999)&(s8_3_H!=9998)&(s8_3_H!=.)&(s8_3_H!=9998)&(s8_3_H!=.))|(strpos(s8_5_I, "Stérilisation féminine (Ligature des")&(s8_3_I!=0)&(s8_3_I!=998)&(s8_3_I!=999)&(s8_3_I!=9998)&(s8_3_I!=.)&(s8_3_I!=9998)&(s8_3_I!=.))|(strpos(s8_5_J, "Stérilisation féminine (Ligature des")&(s8_3_J!=0)&(s8_3_J!=998)&(s8_3_J!=999)&(s8_3_J!=9998)&(s8_3_J!=.)&(s8_3_J!=9998)&(s8_3_J!=.))|(strpos(s8_5_k, "Stérilisation féminine (Ligature des")&(s8_3_k!=0)&(s8_3_k!=998)&(s8_3_k!=999)&(s8_3_k!=9998)&(s8_3_k!=.)&(s8_3_k!=9998)&(s8_3_k!=.)))
replace female_ster_daily_stock_item_p = . if s5_6 == 2 
label values female_ster_daily_stock_item_p yesno


tabout sterilisation_f_fourni female_ster_daily female_ster_daily_stock female_ster_daily_stock_item female_ster_daily_stock_item_p  s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number)


tabout dui_fourni dui_daily dui_daily_trained dui_daily_trained_stock  dui_daily_trained_stock_items dui_daily_trained_stock_items_p s1_7  using "table_cascade.xls", append c(col) npos(both) nlab(Number) 



*************** Raison de non fourniture ************
gen pil_inf = 0
replace pil_inf = 1 if s5_10a_2==1 | s5_10a_3==1

tabout s5_10a_1 pil_inf s5_10a_4 s5_10a_5 s1_7 if   s5_7a == 5 using "table_raison non.xls", append  c(col) npos(both) nlab(Number) h1(pillule) 

		 
gen inj_inf = 0
replace inj_inf = 1 if s5_10b_2==1 | s5_10b_3==1
tabout   s5_10b_1 inj_inf s5_10b_4 s5_10b_5 s1_7 if   s5_7b == 5 using "table_raison non.xls", append  c(col) npos(both) nlab(Number) h1(Injectable) 

gen pm_inf = 0
replace pm_inf = 1 if s5_10c_2==1 | s5_10c_3==1
tabout s5_10c_1 pm_inf s5_10c_4 s5_10c_5  s1_7 if  s5_7c == 5 using "table_raison non.xls", append  c(col) npos(both)  nlab(Number) h1(Préservatif masculin) 

gen pf_in  = 0
replace pf_in = 1 if s5_10d_2==1 | s5_10d_3==1
tabout  s5_10d_1 pf_in s5_10d_4 s5_10d_5 s1_7 if   s5_7d == 5 using "table_raison non.xls", append c(col) npos(both) nlab(Number) h1(Préservatif féminin) 

gen pcu_in  = 0
replace pcu_in = 1 if s5_10e_2==1 | s5_10e_3==1
tabout  s5_10e_1 pcu_in s5_10e_4 s5_10e_5 s1_7 if   s5_7e == 5 using "table_raison non.xls", append c(col) npos(both) nlab(Number) h1(Contraception d'urgence) 

gen dui_in  = 0
replace dui_in = 1 if s5_10f_2==1 | s5_10f_3==1
tabout  s5_10f_1 dui_in s5_10f_4 s5_10f_5 s1_7 if   s5_7f == 5 using "table_raison non.xls", append c(col) npos(both) nlab(Number)  h1(DIU) 

gen implant_in  = 0
replace implant_in = 1 if s5_10g_2==1 | s5_10g_3==1
tabout s5_10g_1 implant_in s5_10g_4 s5_10g_5 s1_7 if  s5_7g == 5 using "table_raison non.xls", append c(col) npos(both) nlab(Number) h1(Implant) 

gen sf_in  = 0
replace sf_in = 1 if s5_10h_2==1 | s5_10h_3==1
tabout   s5_10h_1 sf_in s5_10h_4 s5_10h_5 s1_7 if  s5_7h == 5 using "table_raison non.xls", append c(col) npos(both) nlab(Number) h1(Stérilisation féminine (Ligature des trompes) ) 

gen sm_in  = 0
replace sm_in = 1 if s5_10i_2==1 | s5_10i_3==1
tabout  s5_10i_1 sm_in s5_10i_4 s5_10i_5   s1_7 if   s5_7i== 5 using "table_raison non.xls", append c(col) npos(both) nlab(Number) h1(Stérilisation Masculine ) 


****************** Rupture stock ****


tabout  s6_16a s1_7 if  (s5_7c!=-9 & s5_7c!=5 & s5_7c!=.) using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(Condom Masculin) 

tabout  s6_16b s1_7 if   (s5_7d!=-9 & s5_7d!=5 & s5_7d!=.) using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(Condom Feminin) 

tabout  s6_16c s1_7 if   (s5_7e!=-9 & s5_7e!=5 & s5_7e!=.) using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(PCU) 

tabout  s6_16d s1_7 if   (s5_7b!=-9 & s5_7b!=5 & s5_7b!=.) using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(Injectable depo) 

tabout  s6_16e s1_7 if   (s5_7b!=-9 & s5_7b!=5 & s5_7b!=.)  using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(Injectable sayana) 

tabout  s6_16f s1_7 if   (s5_7g!=-9 & s5_7g!=5 & s5_7g!=.)  using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(Implant) 

tabout  s6_16g s1_7 if   (s5_7a!=-9 & s5_7a!=5 & s5_7a!=.)  using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(PCO) 

tabout  s6_16h s1_7 if   (s5_7a!=-9 & s5_7a!=5 & s5_7a!=.) using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(Pillule progesterone) 

tabout  s6_16i s1_7 if   (s5_7f!=-9 & s5_7f!=5 & s5_7f!=.) using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(DUI) 

tabout  s6_16j s1_7 if  (s5_7h!=-9 & s5_7h!=5 & s5_7h!=.) using "rupture de stock.xls", append c(col) npos(both) nlab(Number) h1(Anneaux tubaire) 

**** Raison de rupture ***********

tabout  s6_19a_1 s6_19a_2 s6_19a_3 s6_19a_4 s6_19a_5 s1_7 if s6_16a==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(Condom Masculin) 

tabout  s6_19b_1 s6_19b_2 s6_19b_3 s6_19b_4 s6_19b_5  s1_7 if s6_16b==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(Condom Feminin) 

tabout  s6_19c_1 s6_19c_2 s6_19c_3 s6_19c_4 s6_19c_5 s1_7 if s6_16c ==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(PCU) 


tabout   s6_19d_1 s6_19d_2 s6_19d_3 s6_19d_4 s6_19d_5 s1_7 if s6_16d==3  using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(Injectable depo) 

tabout  s6_19e_1 s6_19e_2 s6_19e_3 s6_19e_4 s6_19e_5  s1_7 if  s6_16e==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(Injectable sayana) 

tabout   s6_19f_1 s6_19f_2 s6_19f_3 s6_19f_4 s6_19f_5 s1_7 if  s6_16f==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(Implant) 

tabout  s6_19g_1 s6_19g_2 s6_19g_3 s6_19g_4 s6_19g_5 s1_7  if  s6_16g==3  using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(PCO) 

tabout  s6_19h_1 s6_19h_2 s6_19h_3 s6_19h_4 s6_19h_5 s1_7  if  s6_16h==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(Pillule progesterone) 

tabout  s6_19i_1 s6_19i_2 s6_19i_3 s6_19i_4 s6_19i_5 s1_7 if  s6_16i==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(DUI) 

tabout  s6_19j_1 s6_19j_2 s6_19j_3 s6_19j_4 s6_19j_5 s1_7 if  s6_16j==3 using "table_raison rupture.xls", append c(col) npos(both) nlab(Number) h1(Anneaux tubaire) 


*save tabulation_data_withs4, replace
save tabulation_data_withouts4, replace