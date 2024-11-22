library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)


## Read SSC Report
ss_metrics <- read_xlsx("S:/Supply Chain Projects/Data Source (SCE)/Analytics/safety stock vs forecast by loc/ssmetrics.xlsx")


ss_metrics %>%
  janitor::clean_names() %>%
  mutate(year_month = format(as.Date(date, format = "%m/%d/%Y"), "%Y-%m")) %>% 
  filter(item == "24262KFC") %>% 
  group_by(campus, year_month) %>%
  summarise(
    total_safety_stock = sum(campus_ss_on_hand, na.rm = TRUE),
    total_balance_usable = sum(balance_usable, na.rm = TRUE),
    total_cases_below_ss = sum(campus_case_below_ss, na.rm = TRUE),
    count_campus_refs = n_distinct(campus_ref),  
    safety_stock_percentage = (sum(safety_stock, na.rm = TRUE) / sum(balance_usable, na.rm = TRUE)) * 100,
    ss_adherence_percentage = (sum(campus_ss_on_hand, na.rm = TRUE) / sum(campus_ss, na.rm = TRUE)) * 100,
    ss_ratio = total_balance_usable / total_safety_stock
  ) %>%
  ungroup() %>% 
  mutate(campus = as.character(campus)) -> ss_metrics_analysis


## Read Forecast
dsx <- read_excel("S:/Global Shared Folders/Large Documents/S&OP/Demand Planning/BI Forecast Backup/2024/DSX Forecast Backup - 2024.11.18.xlsx")

dsx[-1,] -> dsx
colnames(dsx) <- dsx[1, ]
dsx[-1, ] -> dsx



dsx %>%
  janitor::clean_names() %>%
  mutate(
    year_month = format(as.Date(paste0(substr(forecast_month_year_id, 1, 4), "-", substr(forecast_month_year_id, 5, 6), "-01")), "%Y-%m"), 
    campus = as.character(product_manufacturing_location_code)
  ) %>%
  filter(product_label_sku_code == "24262-KFC") %>% 
  group_by(campus, year_month) %>%
  mutate(adjusted_forecast_cases = as.double(adjusted_forecast_cases)) %>% 
  summarize(forecasted = sum(adjusted_forecast_cases, na.rm = TRUE)) -> dsx_analysis




# Perform the join
ss_metrics_analysis %>%
  left_join(dsx_analysis, by = c("campus" = "campus", "year_month" = "year_month")) %>%
  filter(year_month %in% intersect(ss_metrics_analysis$year_month, dsx_analysis$year_month)) %>%
  mutate(forecasted = as.double(forecasted)) -> final_result


# Write the result
write_xlsx(final_result, "S:/Supply Chain Projects/Data Source (SCE)/Analytics/safety stock vs forecast by loc/ssmetrics_analysis.xlsx")



