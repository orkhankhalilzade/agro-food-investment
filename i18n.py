"""Centralised English / Azerbaijani dictionaries.

Underlying business values (company, orchard, month, stage, ...) always stay in their
canonical form; only the displayed text is translated.
"""
from __future__ import annotations

LANGS = ["EN", "AZ"]

TEXT: dict[str, dict[str, str]] = {
    # header / layout
    "app_title": {"en": "Agro Food Investment", "az": "Agro Food Investment"},
    "app_subtitle": {"en": "IPM & Operational Cost Management", "az": "İPM və əməliyyat xərclərinin idarə edilməsi"},
    "tab_planning": {"en": "Planning", "az": "Planlaşdırma"},
    "tab_inventory": {"en": "Machinery & Equipment", "az": "Texnika və avadanlıq"},
    "save_note": {"en": "Data is kept only in this browser session. Save a file to continue later.",
                  "az": "Məlumat yalnız bu brauzer sessiyasında saxlanır. Sonra davam etmək üçün fayl saxlayın."},
    "save_btn": {"en": "Save progress (file)", "az": "İşi fayl kimi saxla"},
    "load_lbl": {"en": "Load a saved file", "az": "Saxlanmış faylı yüklə"},
    "load_ok": {"en": "File loaded: {n} operation(s).", "az": "Fayl yükləndi: {n} əməliyyat."},
    "load_bad": {"en": "Could not load the file: {err}", "az": "Faylı yükləmək mümkün olmadı: {err}"},
    "language": {"en": "Language", "az": "Dil"},
    # blocks
    "b1": {"en": "Basic Information", "az": "Əsas məlumatlar"},
    "b1_sub": {"en": "Company, orchard and planning period", "az": "Şirkət, bağ və planlaşdırma dövrü"},
    "b2": {"en": "Normatives", "az": "Normativlər"},
    "b2_sub": {"en": "Operational and financial parameters", "az": "Əməliyyat və maliyyə parametrləri"},
    "b3": {"en": "Manual Inputs", "az": "Əl ilə daxiletmə"},
    "b3_sub": {"en": "Operation, products, machinery and labor", "az": "Əməliyyat, preparatlar, texnika və əmək"},
    "b4": {"en": "Automatic Calculations & Results", "az": "Avtomatik hesablamalar və nəticələr"},
    "b4_sub": {"en": "Calculated with the formulas of the company Excel templates",
               "az": "Şirkət Excel şablonlarının formulları ilə hesablanır"},
    "b5": {"en": "Weekly Operations", "az": "Həftəlik əməliyyatlar"},
    "b5_sub": {"en": "All operations, filterable by company, orchard and period",
               "az": "Şirkət, bağ və dövrə görə süzülən bütün əməliyyatlar"},
    # block 1
    "company": {"en": "Company", "az": "Şirkət"},
    "orchard": {"en": "Orchard Name", "az": "Bağ adı"},
    "area": {"en": "Area (ha)", "az": "Sahə (ha)"},
    "stage": {"en": "Phenological Status", "az": "Fenoloji status"},
    "month": {"en": "Month", "az": "Ay"},
    "week": {"en": "Week", "az": "Həftə"},
    "select": {"en": "Select…", "az": "Seçin…"},
    "none_orchard": {"en": "None", "az": "Yoxdur"},
    "area_help": {"en": "Fixed value taken from the company Excel template.",
                  "az": "Şirkət Excel şablonundan götürülmüş sabit dəyər."},
    # block 2
    "n_shift": {"en": "Work Shift (h)", "az": "İş növbəsi (saat)"},
    "n_tank": {"en": "Tank per Area (unit)", "az": "Tank / sahə (vahid)"},
    "n_tank_help": {"en": "1 = round the number of tanks up, 0 = keep the fraction",
                    "az": "1 = tank sayını yuxarı yuvarlaqlaşdır, 0 = kəsri saxla"},
    "n_engine": {"en": "Engine Hours (incl. turns) ×", "az": "Mühərrik saatı (dönmələr daxil) ×"},
    "n_diesel": {"en": "Cost of Diesel (AZN/L)", "az": "Dizelin qiyməti (AZN/L)"},
    "n_lub": {"en": "Cost of Lubricants (AZN/L)", "az": "Sürtkü yağının qiyməti (AZN/L)"},
    # block 3
    "op_info": {"en": "Operation Information", "az": "Əməliyyat məlumatı"},
    "operation": {"en": "Operation", "az": "Əməliyyat"},
    "operation_auto": {"en": "{pt} (automatic number)", "az": "{pt} (avtomatik nömrə)"},
    "solution": {"en": "Spraying Solution per ha (L/ha)", "az": "1 ha üçün püskürtmə məhlulu (L/ha)"},
    "row_len": {"en": "Average Row Length (m)", "az": "Orta cərgə uzunluğu (m)"},
    "rows": {"en": "Rows", "az": "Cərgə sayı"},
    "max_dur": {"en": "Max. Duration of Operation (days)", "az": "Əməliyyatın maksimal müddəti (gün)"},
    "info_note": {"en": "Row length, rows, max. duration and speed are informational only: the Excel formulas do not use them.",
                  "az": "Cərgə uzunluğu, cərgə sayı, maksimal müddət və sürət yalnız məlumat üçündür: Excel formulları onlardan istifadə etmir."},
    "mach": {"en": "Machinery and Tractor", "az": "Texnika və traktor"},
    "tractor": {"en": "Tractor", "az": "Traktor"},
    "tractor_ph": {"en": "Choose from inventory or type", "az": "Siyahıdan seçin və ya yazın"},
    "n_tractors": {"en": "Number of Tractors", "az": "Traktorların sayı"},
    "speed": {"en": "Tractor Speed (km/h)", "az": "Traktorun sürəti (km/saat)"},
    "fuel_w": {"en": "Fuel Consumption with Sprayer (L/h)", "az": "Yanacaq sərfi, çiləyici ilə (L/saat)"},
    "oil_w": {"en": "Oil Consumption with Sprayer (L/h)", "az": "Yağ sərfi, çiləyici ilə (L/saat)"},
    "fuel_wo": {"en": "Fuel Consumption without Sprayer (L/h)", "az": "Yanacaq sərfi, çiləyicisiz (L/saat)"},
    "oil_wo": {"en": "Oil Consumption without Sprayer (L/h)", "az": "Yağ sərfi, çiləyicisiz (L/saat)"},
    "maint": {"en": "Machinery Maintenance Time (h)", "az": "Texnikanın texniki xidmət vaxtı (saat)"},
    "ipm": {"en": "IPM / Agricultural Products", "az": "İPM / kənd təsərrüfatı məhsulları"},
    "pests": {"en": "Pests", "az": "Zərərvericilər"},
    "diseases": {"en": "Diseases", "az": "Xəstəliklər"},
    "nutrition": {"en": "Macro- and Micronutrient Feeding", "az": "Makro və mikroelement qidalandırması"},
    "regulators": {"en": "Plant Regulators (acids, hormones, sun protection)",
                   "az": "Bitki tənzimləyiciləri (turşular, hormonlar, günəşdən qoruma)"},
    "adjuvant": {"en": "Adjuvant / pH Corrector", "az": "Adjuvant / pH korrektoru"},
    "herbicide": {"en": "Herbicide", "az": "Herbisid"},
    "suggest": {"en": "Suggestions from this orchard's program", "az": "Bu bağın proqramından təkliflər"},
    "program_loaded": {"en": "Filled from the Excel program for this week. You can edit it.",
                       "az": "Bu həftə üçün Excel proqramından dolduruldu. Dəyişə bilərsiniz."},
    "program_none": {"en": "The Excel template has no program for this week. Enter products manually.",
                     "az": "Excel şablonunda bu həftə üçün proqram yoxdur. Preparatları əl ilə yazın."},
    "program_pick": {"en": "Select company, orchard and week to load the program.",
                     "az": "Proqramı yükləmək üçün şirkəti, bağı və həftəni seçin."},
    "sprayer": {"en": "Sprayer Information", "az": "Çiləyici məlumatı"},
    "tank_cap": {"en": "Sprayer Tank Capacity (L)", "az": "Çiləyici çəninin tutumu (L)"},
    "reload": {"en": "Time for Tank Reloading (min)", "az": "Çənin doldurulma vaxtı (dəq)"},
    "prod": {"en": "Productivity (ha/h)", "az": "Məhsuldarlıq (ha/saat)"},
    "prod_help": {"en": "Entered manually, exactly as in the Excel templates.",
                  "az": "Excel şablonlarında olduğu kimi əl ilə daxil edilir."},
    "labor": {"en": "Labor Information", "az": "Əmək məlumatı"},
    "salary": {"en": "Gross Salary per Hour (AZN/h)", "az": "Saatlıq ümumi maaş (AZN/saat)"},
    "reset": {"en": "Reset", "az": "Sıfırla"},
    "add": {"en": "Add Operation", "az": "Əməliyyat əlavə et"},
    "update": {"en": "Update Operation", "az": "Əməliyyatı yenilə"},
    "cancel_edit": {"en": "Cancel editing", "az": "Redaktəni ləğv et"},
    "editing": {"en": "Editing {pt}", "az": "{pt} redaktə olunur"},
    "need_fields": {"en": "First select company, orchard, phenological status, month and week.",
                    "az": "Əvvəlcə şirkəti, bağı, fenoloji statusu, ayı və həftəni seçin."},
    "added": {"en": "{pt} added.", "az": "{pt} əlavə edildi."},
    "updated": {"en": "{pt} updated.", "az": "{pt} yeniləndi."},
    "deleted": {"en": "{pt} deleted.", "az": "{pt} silindi."},
    # block 4
    "r_total_cost": {"en": "Total Operational Cost", "az": "Ümumi əməliyyat xərci"},
    "r_salary": {"en": "Gross Salary", "az": "Ümumi maaş"},
    "r_fuel_total": {"en": "Total Fuel and Lubricants", "az": "Ümumi yanacaq və sürtkü"},
    "r_working_days": {"en": "Working Days", "az": "İş günləri"},
    "g_solution": {"en": "Solution and Spraying", "az": "Məhlul və çiləmə"},
    "g_prod": {"en": "Productivity and Time", "az": "Məhsuldarlıq və vaxt"},
    "g_fuel": {"en": "Fuel and Lubricants", "az": "Yanacaq və sürtkü yağları"},
    "r_total_solution": {"en": "Total Solution per Field", "az": "Sahə üzrə ümumi məhlul"},
    "r_tanks": {"en": "Tanks per Area", "az": "Sahə üzrə tank sayı"},
    "r_reload_total": {"en": "Total Time for Tank Reloading", "az": "Tankların doldurulmasının ümumi vaxtı"},
    "r_reload_per_ha": {"en": "Reloading per Ha", "az": "1 ha üçün doldurma"},
    "r_pure": {"en": "Pure Working Productivity", "az": "Xalis iş məhsuldarlığı"},
    "r_output": {"en": "Output per Shift per 1 Tractor", "az": "1 traktorun növbə ərzində məhsuldarlığı"},
    "r_time": {"en": "Time per Operation", "az": "Əməliyyat üzrə vaxt"},
    "r_engine": {"en": "Engine Hours (incl. turns)", "az": "Mühərrik saatı (dönmələr daxil)"},
    "r_diesel_active": {"en": "Diesel — Active", "az": "Dizel — aktiv"},
    "r_lub_active": {"en": "Lubricant — Active", "az": "Sürtkü yağı — aktiv"},
    "r_diesel_transit": {"en": "Diesel — Transit", "az": "Dizel — tranzit"},
    "r_lub_transit": {"en": "Lubricant — Transit", "az": "Sürtkü yağı — tranzit"},
    "r_diesel_cost": {"en": "Cost of Diesel", "az": "Dizelin dəyəri"},
    "r_lub_cost": {"en": "Cost of Lubricants", "az": "Sürtkü yağının dəyəri"},
    "missing": {"en": "Missing input", "az": "Məlumat çatışmır"},
    "missing_list": {"en": "Missing input: {list}", "az": "Çatışmayan məlumatlar: {list}"},
    "save_load": {"en": "Save / Load", "az": "Saxla / Yüklə"},
    # block 5
    "all": {"en": "All", "az": "Hamısı"},
    "export": {"en": "Export to Excel", "az": "Excel-ə ixrac et"},
    "export_none": {"en": "Several companies have operations: pick one in the Company filter to export.",
                    "az": "Bir neçə şirkətin əməliyyatı var: ixrac üçün Şirkət süzgəcindən birini seçin."},
    "export_empty": {"en": "Add at least one operation to export.", "az": "İxrac üçün ən azı bir əməliyyat əlavə edin."},
    "export_info": {"en": "The Excel file will contain all {n} operation(s) of {company}.",
                    "az": "Excel faylında {company} şirkətinin bütün {n} əməliyyatı olacaq."},
    "count_info": {"en": "{n} operation(s) saved for {company}. They are collected in the table in block 5.",
                   "az": "{company} üçün {n} əməliyyat saxlanılıb. Onlar 5-ci blokdakı cədvəldə toplanır."},
    "st_area": {"en": "Total planned area", "az": "Ümumi planlaşdırılan sahə"},
    "st_ops": {"en": "Operations", "az": "Əməliyyatlar"},
    "st_hours": {"en": "Total working hours", "az": "Ümumi iş saatları"},
    "st_diesel": {"en": "Total diesel", "az": "Ümumi dizel"},
    "st_cost": {"en": "Total operational cost", "az": "Ümumi əməliyyat xərci"},
    "st_incomplete": {"en": "{n} operation(s) are incomplete; missing values are not counted.",
                      "az": "{n} əməliyyat natamamdır; çatışmayan dəyərlər hesaba daxil edilmir."},
    "col_company": {"en": "Company", "az": "Şirkət"},
    "col_orchard": {"en": "Orchard Name", "az": "Bağ adı"},
    "col_week": {"en": "Week", "az": "Həftə"},
    "col_operation": {"en": "Operation", "az": "Əməliyyat"},
    "col_area": {"en": "Area (ha)", "az": "Sahə (ha)"},
    "col_tractor": {"en": "Tractor", "az": "Traktor"},
    "col_days": {"en": "Working Days", "az": "İş günləri"},
    "col_diesel": {"en": "Diesel (L)", "az": "Dizel (L)"},
    "col_cost": {"en": "Total Operational Cost (AZN)", "az": "Ümumi əməliyyat xərci (AZN)"},
    "col_status": {"en": "Status", "az": "Status"},
    "status_complete": {"en": "Complete", "az": "Tamamlanıb"},
    "status_incomplete": {"en": "Incomplete", "az": "Natamam"},
    "empty_table": {"en": "No operations match the filters. Add one above.",
                    "az": "Süzgəclərə uyğun əməliyyat yoxdur. Yuxarıda əlavə edin."},
    "select_row": {"en": "Select a row in the table to edit or delete it.",
                   "az": "Redaktə etmək və ya silmək üçün cədvəldən sətir seçin."},
    "edit": {"en": "Edit", "az": "Redaktə et"},
    "delete": {"en": "Delete", "az": "Sil"},
    "confirm_delete": {"en": "Delete operation {pt} ({orchard}, week {week})?",
                       "az": "{pt} əməliyyatı silinsin? ({orchard}, həftə {week})"},
    "confirm": {"en": "Yes, delete", "az": "Bəli, sil"},
    "cancel": {"en": "Cancel", "az": "Ləğv et"},
    # inventory (column titles are bilingual and static, see app.py)
    "inv_help": {"en": "Add, edit or delete rows. Each company fills in its own machinery manually; nothing is pre-filled.",
                 "az": "Sətir əlavə edin, dəyişin və ya silin. Hər şirkət öz texnikasını əl ilə doldurur; heç nə əvvəlcədən yazılmayıb."},
    "inv_tractors": {"en": "Tractor / Machinery Inventory", "az": "Traktor / texnika siyahısı"},
    "inv_impl": {"en": "Implement / Equipment Inventory", "az": "Qoşqu / avadanlıq siyahısı"},
}

MONTH_AZ = {"January": "Yanvar", "February": "Fevral", "March": "Mart", "April": "Aprel", "May": "May",
            "June": "İyun", "July": "İyul", "August": "Avqust", "September": "Sentyabr",
            "October": "Oktyabr", "November": "Noyabr", "December": "Dekabr"}
STAGE_AZ = {"Dormancy": "Yuxu dövrü", "Bud Swelling": "Tumurcuqların şişməsi", "Flowering": "Çiçəkləmə",
            "Fruit Development": "Meyvənin inkişafı", "Ripening and Harvest": "Yetişmə və yığım",
            "Harvest": "Yığım", "Harvest and Dormancy": "Yığım və yuxu dövrü"}

# Missing-input names (keys of calc.OperationInputs)
INPUT_KEYS = {
    "area_ha": "area", "solution_per_ha": "solution", "fuel_with_sprayer": "fuel_w",
    "oil_with_sprayer": "oil_w", "fuel_without_sprayer": "fuel_wo", "oil_without_sprayer": "oil_wo",
    "maintenance_time_h": "maint", "tank_capacity_l": "tank_cap", "productivity_ha_h": "prod",
    "reload_minutes": "reload", "tractors": "n_tractors", "salary_per_hour": "salary",
}


def tr(lang: str, key: str, **kw) -> str:
    s = TEXT[key]["az" if lang == "AZ" else "en"]
    return s.format(**kw) if kw else s


def month_label(lang: str, month: str) -> str:
    return MONTH_AZ[month] if lang == "AZ" else month


def stage_label(lang: str, stage: str) -> str:
    return STAGE_AZ[stage] if lang == "AZ" else stage


def input_label(lang: str, calc_key: str) -> str:
    """Short label of a manual input, without the unit in brackets."""
    full = tr(lang, INPUT_KEYS[calc_key])
    return full.split(" (")[0]
