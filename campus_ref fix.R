
read_xlsx("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/01.16.2024/data fix/Weekly Safety Stock Compliance Report v4 rolling 53 weeks - 01.16.2024.xlsx",
          sheet = "SS metrics") -> ss_metrics


ss_metrics %>% 
  janitor::clean_names() %>% 
  dplyr::select(ref, campus, location, campus_name, item) %>% 
  dplyr::mutate(campus = ifelse(location == "430", 43, campus),
                campus_name = ifelse(location == "430", "BHM", campus_name)) %>% 
  dplyr::mutate(campus_ref = paste0(campus, "-", item)) %>% 
  writexl::write_xlsx("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/01.16.2024/data fix/campus_ref.xlsx")



