library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)

specific_date <- as.Date("2024-06-04")

# (Path revision needed) load main board (mega data) ----
load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Safety_Stock_Compliance/RPA/venturafoods_SafetyStockCompliance_RPA/rds files/ssmetrics_mainboard_05_28_2024.rds")


ssmetrics_mainboard %>%
  dplyr::mutate(ref = gsub("-", "_", ref),
                campus_ref = gsub("-", "_", campus_ref)) -> ssmetrics_mainboard

readr::type_convert(ssmetrics_mainboard) -> ssmetrics_mainboard


############################### Phase 1 ############################
# Campus Abb
campus_abb <- read_excel("S:/Supply Chain Projects/RStudio/Campus Abb.xlsx")

# Category (From BI) ---- 
category_bi <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2022/12.19.2022/BI Category and Platform and pack size.xlsx")

category_bi[-1, ] -> category_bi
colnames(category_bi) <- category_bi[1, ]
category_bi[-1, ] -> category_bi

category_bi %>% 
  dplyr::select(1, 3, 6) %>% 
  dplyr::rename(Item = "SKU Code",
                Category = "Product Category Name",
                Platform = "Product Platform Description") %>% 
  dplyr::mutate(Item = gsub("-", "", Item)) -> category_bi

# Stock type ----
load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Safety_Stock_Compliance/RPA/venturafoods_SafetyStockCompliance_RPA/rds files/stock_type.rds")

# (Path revision needed) Macro-platform (change this only when there's a change) ----
macro_platform <- read_excel("S:/Supply Chain Projects/RStudio/Macro-platform.xlsx",
                             col_names = FALSE)

colnames(macro_platform) <- macro_platform[1, ]
macro_platform[-1, ] -> macro_platform

colnames(macro_platform)[2] <- "macro_platform"

# (Path revision needed) Location_Name (change this only when there's a change) ----
location_name <- read_excel("S:/Supply Chain Projects/RStudio/Location_Name.xlsx",
                            col_names = FALSE)

colnames(location_name) <- location_name[1, ]
location_name[-1, ] -> location_name

location_name %>% 
  dplyr::mutate(Location = as.numeric(Location)) -> location_name

# (Path revision needed) priority_Sku_uniques (change this only when there's a change) ----
priority_sku <- read_excel("S:/Supply Chain Projects/RStudio/Priority_Sku_and_uniques.xlsx",
                           col_names = FALSE)

colnames(priority_sku) <- priority_sku[1, ]
priority_sku[-1, ] -> priority_sku

colnames(priority_sku)[1] <- "priority_sku"

priority_sku %>% 
  dplyr::mutate(Item = priority_sku) -> priority_sku

# (Path revision needed) oil allocation (change this only when there's a change) ----
oil_aloc <- read_excel("S:/Supply Chain Projects/RStudio/oil allocation.xlsx",
                       col_names = FALSE)

colnames(oil_aloc) <- oil_aloc[1, ]
oil_aloc[-1, ] -> oil_aloc

colnames(oil_aloc)[1] <- "Item"
colnames(oil_aloc)[2] <- "oil_aloc"
colnames(oil_aloc)[3] <- "comp_desc"


# (Path revision needed) Inventory Model  (Make sure to remove the password of the original .xlsx file) ----
# Make sure with the password: Elli

# S:Drive - Supply Chain Project - Logistics - SCP - Cost Saving Reporting 

inventory_model <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Optimization by Location - Finished Goods LIVE.xlsx",
                              col_names = FALSE, sheet = "Fin Goods")

inventory_model[-1:-7, ] -> inventory_model
colnames(inventory_model) <- inventory_model[1, ]
inventory_model[-1, ] -> inventory_model


inventory_model %>% 
  janitor::clean_names() %>%
  dplyr::rename(ref = ship_ref,
                mfg_line = mfg_loc_line) %>%
  dplyr::rename_with(~ ifelse(startsWith(., "max_capacity"), "max_capacity", .)) %>%
  dplyr::rename(max_capacity = psa_percent_3_mos_average_by_mfg_platform) %>% 
  dplyr::select(ref, mfg_line, max_capacity) %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) %>% 
  dplyr::mutate(max_capacity = as.numeric(max_capacity)) -> inventory_model



# Campus reference
campus_ref <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Campus reference.xlsx",
                         col_names = FALSE)

colnames(campus_ref) <- campus_ref[1, ]
campus_ref[-1, ] -> campus_ref

colnames(campus_ref)[1] <- "Location"
colnames(campus_ref)[3] <- "Campus"
colnames(campus_ref)[4] <- "campus_no"

campus_ref %>% 
  dplyr::mutate(Campus = replace(Campus, is.na(Campus), 0)) -> campus_ref



# Lot Status
Lot_Status <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Lot Status Code.xlsx",
                         col_names = FALSE)

colnames(Lot_Status) <- Lot_Status[1, ]
Lot_Status[-1, ] -> Lot_Status

Lot_Status %>% 
  dplyr::rename(Lot_Status = "Lot status",
                Hold_Status = "Hard/Soft Hold") %>% 
  dplyr::select(Lot_Status, Hold_Status) -> Lot_Status

# previous SS_Metrics file (This is the most recent file format before we changed "Type" and "Stocking Type Description") - Do not change this until further notice
ssmetrics_pre <- read_excel("S:/Supply Chain Projects/LOGISTICS/SCP/Cost Saving Reporting/SS Compliance/Archive/Copy of Safety Stock Compliance Report Data v3 - 06.05.23.xlsx",
                            col_names = FALSE)

ssmetrics_pre[-1, ] -> ssmetrics_pre
colnames(ssmetrics_pre) <- ssmetrics_pre[1, ]
ssmetrics_pre[-1, ] -> ssmetrics_pre
names(ssmetrics_pre) <- str_replace_all(names(ssmetrics_pre), c(" " = "_"))
names(ssmetrics_pre) <- str_replace_all(names(ssmetrics_pre), c("/" = "_"))


# (Path revision needed) Planner_address Change Directory only when you need to ----
Planner_address <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Address Book/Address Book - 2024.06.03.xlsx", 
                              sheet = "employee", col_types = c("text", 
                                                                "text", "text", "text", "text"))

names(Planner_address) <- str_replace_all(names(Planner_address), c(" " = "_"))

colnames(Planner_address)[1] <- "Planner_No"

Planner_address %>% 
  dplyr::select(1:2) -> Planner_address



# (Path revision needed) exception report ----
exception_report <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/exception report.xlsx")

readr::type_convert(exception_report) -> exception_report

exception_report[-1:-2, ] -> exception_report

colnames(exception_report) <- exception_report[1, ]
exception_report[-1, ] -> exception_report

colnames(exception_report)[1] <- "B_P"
colnames(exception_report)[2] <- "ItemNo"
colnames(exception_report)[3] <- "Buyer"
colnames(exception_report)[4] <- "Planner"
colnames(exception_report)[5] <- "Supplier_No"
colnames(exception_report)[6] <- "Payee Number"
colnames(exception_report)[7] <- "MPF or Line"
colnames(exception_report)[8] <- "Order Policy Code"
colnames(exception_report)[9] <- "Order Policy Value"
colnames(exception_report)[10] <- "Plan Code"
colnames(exception_report)[11] <- "Fence Rule"
colnames(exception_report)[12] <- "Plan Fence Days"
colnames(exception_report)[13] <- "Msg Display Fence"
colnames(exception_report)[14] <- "Freeze Fence"
colnames(exception_report)[15] <- "Leadtime Days"
colnames(exception_report)[16] <- "Reorder MIN"
colnames(exception_report)[17] <- "Reorder MAX"
colnames(exception_report)[18] <- "Reorder Multiple"
colnames(exception_report)[19] <- "Safety Stock"
colnames(exception_report)[20] <- "Reorder Point"
colnames(exception_report)[21] <- "Reorder Qty"
colnames(exception_report)[22] <- "Avg Demand Weeks"
colnames(exception_report)[23] <- "Formula Type"
colnames(exception_report)[24] <- "Sort Code"
colnames(exception_report)[25] <- "Schedule Group"
colnames(exception_report)[26] <- "Model"
colnames(exception_report)[27] <- "Description"
colnames(exception_report)[28] <- "UOM"
colnames(exception_report)[29] <- "PL QTY"
colnames(exception_report)[30] <- "Planning Formula"
colnames(exception_report)[31] <- "Costing Formula"
colnames(exception_report)[32] <- "null"

names(exception_report) <- str_replace_all(names(exception_report), c(" " = "_"))


exception_report %>% 
  dplyr::mutate(ref = paste0(B_P, "_", ItemNo)) %>% 
  dplyr::relocate(ref) -> exception_report

exception_report[!duplicated(exception_report[,c("ref")]),] -> exception_report

# exception report Planner NA to 0
exception_report %>% 
  dplyr::mutate(Planner = replace(Planner, is.na(Planner), 0)) -> exception_report



# (Path revision needed) custord custord ----
# Open Customer Order File pulling ----  Change Directory ----
custord <- read.xlsx("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/US and CAN OO BT where status _ J.xlsx",
                     colNames = FALSE)

custord %>% 
  dplyr::slice(c(-1, -3)) -> custord

colnames(custord) <- custord[1, ]
custord[-1, ] -> custord


custord %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(product_label_sku = gsub("-", "", product_label_sku)) %>% 
  dplyr::mutate(ref = paste0(location, "_", product_label_sku)) %>% 
  dplyr::mutate(oo_cases = as.double(oo_cases),
                oo_cases = ifelse(is.na(oo_cases), 0, oo_cases),
                b_t_open_order_cases = as.double(b_t_open_order_cases),
                b_t_open_order_cases = ifelse(is.na(b_t_open_order_cases), 0, b_t_open_order_cases)) %>%
  dplyr::mutate(Qty = oo_cases + b_t_open_order_cases) %>% 
  dplyr::mutate(sales_order_requested_ship_date = as_date(as.integer(sales_order_requested_ship_date), origin = "1899-12-30")) %>% 
  dplyr::select(ref, product_label_sku, location, Qty, sales_order_requested_ship_date) %>% 
  dplyr::rename(Item = product_label_sku,
                Location = location,
                date = sales_order_requested_ship_date) %>% 
  dplyr::group_by(ref, Item, Location, date) %>% 
  dplyr::summarise(Qty = sum(Qty)) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= specific_date & date < specific_date +7, "Y", "N")) %>% 
  dplyr::relocate(in_next_7_days, .after = Location) -> custord




# Custord pivot
reshape2::dcast(custord, ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(total_cust_order = N + Y) -> custord_pivot



# (Path revision needed) Custord wo ----
wo <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/Wo.xlsx")


wo %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= specific_date & date < specific_date+7, "Y", "N")) %>% 
  dplyr::rename(Item = item,
                Location = location,
                Qty = production_scheduled_cases) %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(N = as.integer(N)) -> wo_pivot



# (Path revision needed) Custord Receipt ----
receipt <- read.csv("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/DSXIE/2024/06.04/receipt.csv",
                    header = FALSE)


receipt %>% 
  dplyr::select(-1) %>% 
  dplyr::slice(-1) %>% 
  dplyr::rename(aa = V2) %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::select(-"3", -"8") %>% 
  dplyr::rename(aa = "1") %>% 
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp,-"4") %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                po_no = "6",
                date = "7") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref, Item, Location) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= specific_date & date < specific_date+7, "Y", "N") ) -> receipt



# Receipt Pivot
receipt %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) -> receipt_pivot  


# (Path revision needed) Custord PO ----
po <- read.csv("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/DSXIE/2024/06.04/po.csv",
               header = FALSE)

po %>% 
  dplyr::select(-1) %>% 
  dplyr::slice(-1) %>% 
  dplyr::rename(aa = V2) %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::select(-"3") %>% 
  dplyr::rename(aa = "1") %>% 
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp, -"4", -"8") %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                po_no = "6",
                date = "7") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref, Item, Location) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= specific_date & date< specific_date + 7, "Y", "N") ) -> po


# PO_Pivot 
po %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(N = as.integer(N),
                Y = as.integer(Y)) -> PO_Pivot




# ###############################################################################################################################
# ######################################## From here, this should be deactivated after two location resolved ####################
# ###############################################################################################################################



JDOH_complete <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/inventory.xlsx",
                            sheet = "FG")
JDOH_complete_2 <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/inventory.xlsx",
                              sheet = "RM")


JDOH_complete[-1, ] -> JDOH_complete
colnames(JDOH_complete) <- JDOH_complete[1, ]
JDOH_complete[-1, ] -> JDOH_complete

JDOH_complete %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>% 
  dplyr::mutate(item = str_replace(item, "^0+(?!$)", "")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::select(ref, inventory_hold_status, current_inventory_balance) %>% 
  dplyr::mutate(current_inventory_balance = as.double(current_inventory_balance)) -> JDOH_complete_1


JDOH_complete_1 %>% 
  dplyr::group_by(ref, inventory_hold_status) %>% 
  dplyr::summarise(current_inventory_balance = sum(current_inventory_balance)) %>% 
  tidyr::pivot_wider(names_from = inventory_hold_status, 
                     values_from = current_inventory_balance) %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(useable = replace(useable, is.na(useable), 0),
                hard_hold = replace(hard_hold, is.na(hard_hold), 0),
                soft_hold = replace(soft_hold, is.na(soft_hold), 0)) %>%
  dplyr::mutate(Balance_Usable = useable + soft_hold,
                Balance_Hold = hard_hold,
                On_Hand = Balance_Usable + Balance_Hold) %>% 
  dplyr::select(-useable, -hard_hold, -soft_hold) -> JDOH_complete_1




JDOH_complete %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>% 
  dplyr::mutate(item = str_replace(item, "^0+(?!$)", "")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::select(ref, location, item, description) %>% 
  dplyr::distinct() %>% 
  dplyr::left_join(JDOH_complete_1, by = "ref") -> JDOH_complete



JDOH_complete %>% 
  dplyr::left_join(exception_report %>% dplyr::select(ref, Safety_Stock), by = "ref") %>% 
  dplyr::mutate(Safety_Stock = replace(Safety_Stock, is.na(Safety_Stock), 0)) %>% 
  dplyr::mutate(Safety_Stock = as.double(Safety_Stock)) %>% 
  dplyr::left_join(exception_report %>% dplyr::select(ref, Planner), by = "ref") %>% 
  dplyr::rename(Planner_No = Planner) %>% 
  dplyr::mutate(Planner_No = replace(Planner_No, is.na(Planner_No), 0)) %>% 
  dplyr::left_join(Planner_address) %>% 
  dplyr::mutate(Alpha_Name = replace(Alpha_Name, is.na(Alpha_Name), 0)) %>% 
  dplyr::rename(Planner_Name = Alpha_Name,
                Description = description,
                Item = item,
                Location = location) -> JDOH_complete



JDOH_complete %>% 
  filter(!str_starts(Description, "PWS ") & 
           !str_starts(Description, "SUB ") & 
           !str_starts(Description, "THW ") & 
           !str_starts(Description, "PALLET")) -> JDOH_complete





################### RM ####################

JDOH_complete_2[-1, ] -> JDOH_complete_2
colnames(JDOH_complete_2) <- JDOH_complete_2[1, ]
JDOH_complete_2[-1, ] -> JDOH_complete_2

JDOH_complete_2 %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>% 
  dplyr::mutate(item = str_replace(item, "^0+(?!$)", "")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::select(ref, inventory_hold_status, current_inventory_balance) %>% 
  dplyr::mutate(current_inventory_balance = as.double(current_inventory_balance)) -> JDOH_complete_2_2


JDOH_complete_2_2 %>% 
  dplyr::group_by(ref, inventory_hold_status) %>% 
  dplyr::summarise(current_inventory_balance = sum(current_inventory_balance)) %>% 
  tidyr::pivot_wider(names_from = inventory_hold_status, 
                     values_from = current_inventory_balance) %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(useable = replace(useable, is.na(useable), 0),
                hard_hold = replace(hard_hold, is.na(hard_hold), 0),
                soft_hold = replace(soft_hold, is.na(soft_hold), 0)) %>%
  dplyr::mutate(Balance_Usable = useable + soft_hold,
                Balance_Hold = hard_hold,
                On_Hand = Balance_Usable + Balance_Hold) %>% 
  dplyr::select(-useable, -hard_hold, -soft_hold) -> JDOH_complete_2_2




JDOH_complete_2 %>% 
  janitor::clean_names() %>% 
  dplyr::mutate(item = gsub("-", "", item)) %>% 
  dplyr::mutate(item = str_replace(item, "^0+(?!$)", "")) %>% 
  dplyr::mutate(ref = paste0(location, "_", item)) %>% 
  dplyr::select(ref, location, item, description) %>% 
  dplyr::distinct() %>% 
  dplyr::left_join(JDOH_complete_2_2, by = "ref") -> JDOH_complete_2



JDOH_complete_2 %>% 
  dplyr::left_join(exception_report %>% dplyr::select(ref, Safety_Stock), by = "ref") %>% 
  dplyr::mutate(Safety_Stock = replace(Safety_Stock, is.na(Safety_Stock), 0)) %>% 
  dplyr::mutate(Safety_Stock = as.double(Safety_Stock)) %>% 
  dplyr::left_join(exception_report %>% dplyr::select(ref, Planner), by = "ref") %>% 
  dplyr::rename(Planner_No = Planner) %>% 
  dplyr::mutate(Planner_No = replace(Planner_No, is.na(Planner_No), 0)) %>% 
  dplyr::left_join(Planner_address) %>% 
  dplyr::mutate(Alpha_Name = replace(Alpha_Name, is.na(Alpha_Name), 0)) %>% 
  dplyr::rename(Planner_Name = Alpha_Name,
                Description = description,
                Item = item,
                Location = location) -> JDOH_complete_2



JDOH_complete_2 %>% 
  filter(!str_starts(Description, "PWS ") & 
           !str_starts(Description, "SUB ") & 
           !str_starts(Description, "THW ") & 
           !str_starts(Description, "PALLET")) -> JDOH_complete_2


rbind(JDOH_complete, JDOH_complete_2) -> JDOH_complete

####################################################################################################################################################
####################################################################################################################################################
############################################################### 6/5/2023 Update ####################################################################
####################################################################################################################################################
####################################################################################################################################################


# Add Location 430
loc_430_for_jdoh <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/430.xlsx")
loc_430_for_jdoh[-1:-3, ] -> loc_430_for_jdoh
colnames(loc_430_for_jdoh) <- loc_430_for_jdoh[1, ]
loc_430_for_jdoh[-1, ] -> loc_430_for_jdoh
loc_430_for_jdoh[-nrow(loc_430_for_jdoh), ] -> loc_430_for_jdoh


loc_430_for_jdoh %>% 
  janitor::clean_names() %>% 
  dplyr::select(sku, base_available_qty, base_on_hand_qty) %>% 
  tidyr::separate(sku, c("1", "2", "3"), sep = "-") %>% 
  janitor::clean_names() %>% 
  dplyr::select(x1, base_available_qty, base_on_hand_qty) %>% 
  dplyr::rename(Item = x1,
                Balance_Usable = base_available_qty,
                On_Hand = base_on_hand_qty) %>% 
  dplyr::mutate(Balance_Usable = as.numeric(Balance_Usable),
                On_Hand = as.numeric(On_Hand)) %>% 
  dplyr::mutate(Item = str_replace(Item, "^0+(?!$)", "")) %>% 
  dplyr::mutate(ref = paste0("430", "_", Item),
                Location = "430",
                Safety_Stock = 0,
                Planner_No = "113444",
                Planner_Name = "WHITE, STEPHANIE",
                Description = "",
                Balance_Hold = 0) %>% 
  dplyr::relocate(ref, Location, Item, Description, Balance_Usable, Balance_Hold, On_Hand, Safety_Stock, Planner_No, Planner_Name) %>% 
  dplyr::mutate(Safety_Stock = as.double(Safety_Stock))  -> loc_430_for_jdoh




rbind(JDOH_complete, loc_430_for_jdoh) -> JDOH_complete



############ 25 & 55 label inventory ##########
lot_status_code <- read_excel("S:/Supply Chain Projects/Data Source (SCE)/Lot Status Code.xlsx")

lot_status_code %>% 
  janitor::clean_names() %>% 
  dplyr::select(lot_status, hard_soft_hold) %>% 
  dplyr::mutate(lot_status = ifelse(is.na(lot_status), "Useable", lot_status),
                hard_soft_hold = ifelse(is.na(hard_soft_hold), "Useable", hard_soft_hold)) %>% 
  dplyr::rename(status = lot_status) -> lot_status_code



jde_inv_for_25_55_label <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/JDE 25,55.xlsx")

jde_inv_for_25_55_label[-1:-5, ] -> jde_inv_for_25_55_label
colnames(jde_inv_for_25_55_label) <- jde_inv_for_25_55_label[1, ]
jde_inv_for_25_55_label[-1, ] -> jde_inv_for_25_55_label


jde_inv_for_25_55_label %>% 
  janitor::clean_names() %>% 
  dplyr::rename(b_p = bp,
                item = item_number) %>% 
  dplyr::mutate(status = ifelse(is.na(status), "Useable", status)) %>% 
  dplyr::mutate(item = as.numeric(item),
                on_hand = as.numeric(on_hand),
                b_p = as.numeric(b_p)) %>% 
  dplyr::filter(!is.na(item)) %>% 
  dplyr::left_join(lot_status_code, by = "status") %>% 
  dplyr::select(-status) %>% 
  pivot_wider(names_from = hard_soft_hold, values_from = on_hand, values_fn = list(on_hand = sum)) %>% 
  janitor::clean_names() %>% 
  replace_na(list(useable = 0, soft_hold = 0, hard_hold = 0)) %>% 
  dplyr::left_join(exception_report %>% 
                     janitor::clean_names() %>% 
                     dplyr::select(item_no, mpf_or_line) %>% 
                     dplyr::rename(item = item_no,
                                   label = mpf_or_line) %>% 
                     dplyr::mutate(item = as.double(item)) %>% 
                     dplyr::filter(label == "LBL") %>% 
                     dplyr::distinct(item, label)) %>% 
  dplyr::filter(!is.na(label)) %>% 
  dplyr::select(-label) %>% 
  dplyr::mutate(ref = paste0(b_p, "_", item)) %>% 
  dplyr::mutate(useable = useable + soft_hold) %>% 
  dplyr::select(-soft_hold) %>% 
  dplyr::mutate(on_hand = useable + hard_hold) %>%
  dplyr::select(ref, b_p, item, description, useable, hard_hold, on_hand) %>% 
  dplyr::left_join(exception_report %>% dplyr::select(ref, Safety_Stock, Planner), by = "ref") %>% 
  dplyr::rename(Planner_No = Planner) %>% 
  dplyr::mutate(Planner_No = replace(Planner_No, is.na(Planner_No), 0)) %>% 
  dplyr::left_join(Planner_address) %>% 
  dplyr::mutate(Alpha_Name = replace(Alpha_Name, is.na(Alpha_Name), 0)) %>% 
  dplyr::rename(Planner_Name = Alpha_Name) %>% 
  dplyr::rename(Location = b_p,
                Item = item,
                Description = description,
                Balance_Usable = useable,
                Balance_Hold = hard_hold,
                On_Hand = on_hand) -> jde_inv_for_25_55_label


rbind(JDOH_complete, jde_inv_for_25_55_label) -> JDOH_complete


## Bring over Inv_Bal ##
inv_bal <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/inv_bal.xlsx")
inv_bal[-1:-2, ] -> inv_bal
colnames(inv_bal) <- inv_bal[1, ]
inv_bal[-1, ] -> inv_bal


inv_bal %>% 
  janitor::clean_names() %>% 
  readr::type_convert() %>% 
  dplyr::mutate(bp = as.numeric(bp)) %>% 
  dplyr::mutate(ref = paste0(bp, "_", item)) %>%
  dplyr::select(ref, type, gl_class) %>% 
  dplyr::rename(Stock_Type = type,
                GL_Class = gl_class) %>% 
  dplyr::distinct() -> inv_bal_1


JDOH_complete %>% 
  dplyr::left_join(inv_bal_1, by = "ref") %>% 
  dplyr::mutate(Lot_Status = "") %>% 
  dplyr::select(ref, Location, Item, Stock_Type, Description, Balance_Usable, Balance_Hold, Lot_Status, On_Hand, Safety_Stock, GL_Class, Planner_No, Planner_Name) -> JDOH_complete





################################################################################################################################
############################################## up to here. two location resovled -> deactivated ################################
################################################################################################################################


################################# SS Metrics ##################################
ssmetrics <- ""
data.frame(ssmetrics) -> ssmetrics

ssmetrics %>%
  dplyr::select(-ssmetrics) %>% 
  dplyr::mutate(ref = "",
                Location = "",
                Item = "",
                Stock_Type = "",
                Description = "",
                Balance_Usable = "",
                Balance_Hold = "",
                Lot_Status = "",
                On_Hand = "",
                Safety_Stock = "",
                GL_Class = "",
                Planner_No = "",
                Planner_Name = "") -> ssmetrics


rbind(ssmetrics, JDOH_complete) -> ssmetrics

ssmetrics[-1, ] -> ssmetrics

readr::type_convert(ssmetrics) -> ssmetrics

ssmetrics %>% 
  dplyr::mutate(date = specific_date) %>% 
  dplyr::relocate(date, .after = ref) -> ssmetrics


# Type - vlookup
ssmetrics_pre[-which(duplicated(ssmetrics_pre$Item)),] -> ssmetrics_pre_1

merge(ssmetrics, ssmetrics_pre_1[, c("Item", "Type")], by = "Item", all.x = TRUE) -> ssmetrics



# Stocking Type Description - vlookup
merge(ssmetrics, stock_type[, c("Stock_Type", "Stocking_Type_Description")], by = "Stock_Type", all.x = TRUE) -> ssmetrics


# MTO/MTS - vlookup
merge(ssmetrics, exception_report[, c("ref", "Order_Policy_Code")], by = "ref", all.x = TRUE) %>% 
  dplyr::rename(MTO_MTS = Order_Policy_Code) -> ssmetrics


# Lot Status, Hold Status - vlookup
ssmetrics %>% 
  dplyr::mutate(Lot_Status = replace(Lot_Status, is.na(Lot_Status),"")) -> ssmetrics

merge(ssmetrics, Lot_Status[, c("Lot_Status", "Hold_Status")], by = "Lot_Status", all.x = TRUE) -> ssmetrics

ssmetrics %>% 
  dplyr::mutate(Hold_Status.y = ifelse(Lot_Status == "", "", Hold_Status)) %>% 
  dplyr::select(-Hold_Status) %>% 
  dplyr::rename(Hold_Status = Hold_Status.y) -> ssmetrics



# MPF - vlookup
merge(ssmetrics, exception_report[, c("ref", "MPF_or_Line")], by = "ref", all.x = TRUE) %>% 
  dplyr::rename(MPF = MPF_or_Line) -> ssmetrics



# MTO/MTS
ssmetrics %>% 
  dplyr::mutate(MTO_MTS = ifelse(is.na(MTO_MTS) & Stocking_Type_Description == "Obsolete - Use Up", "DNRR", MTO_MTS)  ) %>% 
  dplyr::mutate(MPF = ifelse(is.na(MPF) & Stocking_Type_Description == "Obsolete - Use Up", "DNRR", MPF)  ) %>% 
  dplyr::mutate(MTO_MTS = ifelse(is.na(MTO_MTS) & Stocking_Type_Description == "Consigned Inventory", "N/A", MTO_MTS)  ) %>% 
  dplyr::mutate(MPF = ifelse(is.na(MPF) & Stocking_Type_Description == "Consigned Inventory", "N/A", MPF)  ) -> ssmetrics

# split the data with Type NA and Type !NA

ssmetrics %>% 
  dplyr::filter(!is.na(Type)) -> ssmetrics_1

ssmetrics %>% 
  dplyr::filter(is.na(Type)) -> ssmetrics_2

# Type N/A
ssmetrics_mainboard %>% 
  dplyr::select(Item, Type) -> ssmetrics_mainboard_type

ssmetrics_mainboard_type[!duplicated(ssmetrics_mainboard_type[,c("Item", "Type")]),] -> ssmetrics_mainboard_type

merge(ssmetrics_2, ssmetrics_mainboard_type[, c("Item", "Type")], by = "Item", all.x = TRUE) %>% 
  dplyr::relocate(Type.y, .after = Type.x) %>% 
  dplyr::select(-Type.x) %>% 
  dplyr::rename(Type = Type.y) -> ssmetrics_2


rbind(ssmetrics_1, ssmetrics_2) -> ssmetrics

ssmetrics %>% 
  dplyr::mutate(Type = ifelse(Stocking_Type_Description == "WIP", "WIP", Type)) -> ssmetrics

ssmetrics %>% 
  dplyr::filter(Type %in% c("Finished Goods", "Ingredients", "Label", "Packaging", NA)) %>% 
  dplyr::arrange(Type) -> ssmetrics



#####################################################

# category vlookup from category_bi
category_bi[!duplicated(category_bi[,c("Item", "Category", "Platform")]),] -> category_bi
merge(ssmetrics, category_bi[, c("Item", "Category")], by = "Item", all.x = TRUE) -> ssmetrics  


# Split the data for Category with NA and !NA
ssmetrics %>% 
  dplyr::filter(!is.na(Category)) -> ssmetrics_cat_not_na


ssmetrics %>% 
  dplyr::filter(is.na(Category)) -> ssmetrics_cat_na

ssmetrics_cat_na %>% 
  dplyr::mutate(Category = ifelse(Type == "Packaging" | Type == "Ingredients" | Type == "Label", Type, NA)) -> ssmetrics_cat_na

ssmetrics_cat_na %>% 
  dplyr::filter(!is.na(Category)) -> ssmetrics_cat_passed

ssmetrics_cat_na %>% 
  dplyr::filter(is.na(Category)) -> cat_mega


ssmetrics_mainboard %>% 
  dplyr::select(Item, Category) -> ssmetrics_mainboard_cat

ssmetrics_mainboard_cat[!duplicated(ssmetrics_mainboard_cat[,c("Item", "Category")]),] -> ssmetrics_mainboard_cat
ssmetrics_mainboard_cat[-which(duplicated(ssmetrics_mainboard_cat$Item)),] -> ssmetrics_mainboard_cat


merge(cat_mega, ssmetrics_mainboard_cat[, c("Item", "Category")], by = "Item", all.x = TRUE) %>% 
  dplyr::select(-Category.x) %>% 
  dplyr::rename(Category = Category.y) -> cat_mega  



rbind(ssmetrics_cat_not_na, ssmetrics_cat_passed, cat_mega) -> ssmetrics



# Platform vlookup from category_bi
merge(ssmetrics, category_bi[, c("Item", "Platform")], by = "Item", all.x = TRUE) -> ssmetrics  

# Split the data for Platform with NA and !NA
ssmetrics %>% 
  dplyr::filter(!is.na(Platform)) -> ssmetrics_plt_not_na

ssmetrics %>% 
  dplyr::filter(is.na(Platform)) -> ssmetrics_plt_na

ssmetrics_plt_na %>% 
  dplyr::mutate(Platform = ifelse(Type == "Packaging" | Type == "Ingredients" | Type == "Label", Type, NA)) -> ssmetrics_plt_na

ssmetrics_plt_na %>% 
  dplyr::filter(!is.na(Platform)) -> ssmetrics_plt_passed

ssmetrics_plt_na %>% 
  dplyr::filter(is.na(Platform)) -> plt_mega


ssmetrics_mainboard %>% 
  dplyr::select(Item, Platform) -> ssmetrics_mainboard_plt

ssmetrics_mainboard_plt[!duplicated(ssmetrics_mainboard_plt[,c("Item", "Platform")]),] -> ssmetrics_mainboard_plt
ssmetrics_mainboard_plt[-which(duplicated(ssmetrics_mainboard_plt$Item)),] -> ssmetrics_mainboard_plt

merge(plt_mega, ssmetrics_mainboard_plt[, c("Item", "Platform")], by = "Item", all.x = TRUE) %>% 
  dplyr::select(-Platform.x) %>% 
  dplyr::rename(Platform = Platform.y) -> plt_mega  


rbind(ssmetrics_plt_not_na, ssmetrics_plt_passed, plt_mega) -> ssmetrics


#####################################################

# Pivot Hold Qty
ssmetrics %>% 
  dplyr::filter(Hold_Status %in% c("", "Soft")) %>% 
  reshape2::dcast(date + Location + Item + Description ~ . , value.var = "Balance_Hold", sum) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref) %>% 
  dplyr::rename(Balance_Hold = ".") -> Pivot_hold_qty

# Pivot itmbal
ssmetrics %>% 
  reshape2::dcast(date + Location + Item + Description + Type + Stocking_Type_Description + Planner_Name + MTO_MTS + MPF + Safety_Stock +
                    Balance_Usable + Category + Platform ~ .) %>% 
  dplyr::select(-.) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref) -> Pivot_itmbal


merge(Pivot_itmbal, Pivot_hold_qty[, c("ref", "Balance_Hold")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Balance_Hold = replace(Balance_Hold, is.na(Balance_Hold), 0)) %>% 
  dplyr::relocate(-ref)-> Pivot_itmbal



##################### Final SS Metrics #####################
Pivot_itmbal -> ssmetrics_final

# campus & campus_ref
merge(ssmetrics_final, campus_ref[, c("Location", "Campus")], by = "Location", all.x = TRUE) %>% 
  dplyr::mutate(campus_ref = ifelse(Campus != 0, paste0(Campus, "-", Item), ref)) %>% 
  dplyr::mutate(Campus = gsub(0, "", Campus)) %>% 
  dplyr::relocate(Campus) -> ssmetrics_final

# Current SS Alert
ssmetrics_final %>% 
  dplyr::mutate(current_ss_alert = ifelse(Safety_Stock == 0, "N/A", 
                                          ifelse(Balance_Usable + Balance_Hold < Safety_Stock, "Below SS", "OK"))) -> ssmetrics_final




# Total Cust Order
merge(ssmetrics_final, custord_pivot[, c("ref", "total_cust_order")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(total_cust_order = replace(total_cust_order, is.na(total_cust_order), 0)) -> ssmetrics_final


# cust_order_qty_in_the_next_5_days
merge(ssmetrics_final, custord_pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::rename(cust_order_in_the_next_5_days = Y) -> ssmetrics_final

# wo_qty_in_the_next_5_days
merge(ssmetrics_final, wo_pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::rename(wo_qty_in_the_next_5_days = Y) -> ssmetrics_final

# receipt_qty_in_the_next_5_days
merge(ssmetrics_final, receipt_pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::rename(receipt_qty_in_the_next_5_days = Y) -> ssmetrics_final

# po_qty_in_the_next_5_days
merge(ssmetrics_final, PO_Pivot[, c("ref", "Y")], by = "ref", all.x = TRUE) %>% 
  dplyr::mutate(Y = replace(Y, is.na(Y), 0)) %>% 
  dplyr::rename(po_qty_in_the_next_5_days = Y) -> ssmetrics_final


# SS Alert after Cust Order in the next 5 days + WO & Receipt
ssmetrics_final %>% 
  dplyr::mutate(ss_alert_after = ifelse(Safety_Stock == 0, "N/A",
                                        ifelse((Balance_Usable + Balance_Hold + wo_qty_in_the_next_5_days + 
                                                  receipt_qty_in_the_next_5_days + 
                                                  po_qty_in_the_next_5_days - cust_order_in_the_next_5_days) >= Safety_Stock, "OK",
                                               ifelse(wo_qty_in_the_next_5_days + receipt_qty_in_the_next_5_days + 
                                                        po_qty_in_the_next_5_days == 0, 
                                                      "Below SS with no supply","Below SS")))) -> ssmetrics_final


# Campus SS
plyr::ddply(ssmetrics_final, "campus_ref", transform, campus_ss = sum(Safety_Stock)) -> ssmetrics_final


# campus_total_available
plyr::ddply(ssmetrics_final, "campus_ref", transform, campus_total_available_1 = sum(Balance_Usable)) -> ssmetrics_final
plyr::ddply(ssmetrics_final, "campus_ref", transform, campus_total_available_2 = sum(Balance_Hold)) -> ssmetrics_final

ssmetrics_final %>% 
  dplyr::mutate(campus_total_available_1 = ifelse(is.na(campus_total_available_1), 0, campus_total_available_1),
                campus_total_available_2 = ifelse(is.na(campus_total_available_2), 0, campus_total_available_2)) %>%
  dplyr::mutate(campus_total_available = campus_total_available_1 + campus_total_available_2) %>% 
  dplyr::select(-campus_total_available_1, -campus_total_available_2) -> ssmetrics_final

# month, FY, year
ssmetrics_final %>% 
  mutate(month = month(date),
         year = year(date),
         FY = ifelse(specific_date < make_date(year = year(specific_date), month = 4, day = 1),
                     paste("FY", year(date)),
                     paste("FY", year(date) + 1))) -> ssmetrics_final



# Macro-platform
ssmetrics_final %>% 
  dplyr::left_join(macro_platform, by = "Platform") %>% 
  dplyr::mutate(macro_platform = ifelse(is.na(macro_platform), Type, macro_platform)) -> ssmetrics_final

# Location_Name
ssmetrics_final %>% 
  dplyr::left_join(location_name %>% select(1, 2), by = "Location") -> ssmetrics_final



# Label
ssmetrics_final %>% 
  dplyr::mutate(Label = ifelse(Type == "Finished Goods", "label", NA)) -> ssmetrics_final

ssmetrics_final$Item -> temp_item

substr(temp_item, nchar(temp_item)-2, nchar(temp_item)) -> temp_label
cbind(ssmetrics_final, temp_label) -> ssmetrics_final

ssmetrics_final %>% 
  dplyr::mutate(Label = ifelse(Label == "label", temp_label, NA)) %>% 
  dplyr::select(-temp_label) -> ssmetrics_final


# Sku_has_ss
ssmetrics_final %>% 
  dplyr::mutate(Sku_has_ss = ifelse(Safety_Stock > 0, 1, 0)) -> ssmetrics_final


# Sku_greater_or_equal_ss
ssmetrics_final %>% 
  dplyr::mutate(Sku_greater_or_equal_ss = ifelse(Sku_has_ss == 1 & Balance_Usable + Balance_Hold >= Safety_Stock, 1, 0)) -> ssmetrics_final


# Sku less ss
ssmetrics_final %>% 
  dplyr::mutate(Sku_less_ss = ifelse(Safety_Stock > (Balance_Usable + Balance_Hold), 1, 0)) -> ssmetrics_final

# Sku less ss with supply
ssmetrics_final %>% 
  dplyr::mutate(Sku_less_ss_with_supply = ifelse(Safety_Stock > (Balance_Usable + Balance_Hold) & 
                                                   (wo_qty_in_the_next_5_days + receipt_qty_in_the_next_5_days + 
                                                      po_qty_in_the_next_5_days > 0), 1, 0)) -> ssmetrics_final


# campus Sku has ss
ssmetrics_final %>% 
  dplyr::mutate(campus_Sku_ss =ifelse(campus_ss != 0 & campus_ss > 0, 1, 0)) -> ssmetrics_final

# campus Sku greater equal ss
ssmetrics_final %>% 
  dplyr::mutate(campus_Sku_greater_equal_ss = ifelse(campus_Sku_ss == 1 & campus_total_available >= campus_ss, 1, 0)) -> ssmetrics_final


# campus Sku less ss
ssmetrics_final %>% 
  dplyr::mutate(campus_Sku_less_ss = ifelse(campus_ss > campus_total_available, 1, 0)) -> ssmetrics_final

# Priority Sku or unique RM
ssmetrics_final %>% 
  dplyr::left_join(priority_sku, by = "Item") %>% 
  dplyr::mutate(priority_sku_unique = ifelse(is.na(priority_sku), "N", "Y")) %>% 
  dplyr::select(-priority_sku) -> ssmetrics_final


# oil allocation sku
ssmetrics_final %>% 
  dplyr::left_join(oil_aloc %>% dplyr::select(1:2), by = "Item") %>% 
  dplyr::mutate(oil_aloc_2 = ifelse(Type != "Finished Goods", Type, NA)) %>% 
  dplyr::mutate(oil_aloc_3 = ifelse(is.na(oil_aloc) & is.na(oil_aloc_2), "non oil allocation", NA)) %>% 
  dplyr::mutate(oil_allocation = oil_aloc,
                oil_allocation = ifelse(is.na(oil_aloc), oil_aloc_2, oil_aloc),
                oil_allocation = ifelse(is.na(oil_allocation), oil_aloc_3, oil_allocation)) %>% 
  dplyr::select(-oil_aloc, -oil_aloc_2, -oil_aloc_3) -> ssmetrics_final


# mfg_line & max capacity
ssmetrics_final %>% 
  dplyr::left_join(inventory_model, by = "ref") %>% 
  dplyr::mutate(mfg_line = ifelse(is.na(mfg_line), Type, mfg_line)) %>% 
  dplyr::mutate(max_capacity = replace(max_capacity, is.na(max_capacity), 0)) -> ssmetrics_final

# Capacity Status
ssmetrics_final %>% 
  dplyr::mutate(capacity_status = ifelse(Type == "Finished Goods",
                                         ifelse(max_capacity > 0.85, "Constrained", 
                                                ifelse(max_capacity < 0.75, "OK", "Check")), Type)) -> ssmetrics_final


# max_capacity retouch
ssmetrics_final %>% 
  dplyr::mutate(max_capacity = paste0(round(100*max_capacity, 0), "%")) -> ssmetrics_final


# month final touch
ssmetrics_final %>% 
  dplyr::mutate(month = recode(month, "1" = "Jan", "2" = "Feb", "3" = "Mar", "4" = "Apr", "5" = "May", "6" = "Jun", "7" = "Jul", "8" = "Aug", "9" = "Sep", "10" = "Oct", "11" = "Nov", "12" = "Dec")) -> ssmetrics_final

## RELOCATING ##
ssmetrics_final %>% 
  dplyr::relocate(month, FY, year, Category, Platform, macro_platform, Location_Name, Campus, date, Location, Item, Description, Type,
                  Stocking_Type_Description, Planner_Name, MTO_MTS, MPF, Safety_Stock, Balance_Usable, Balance_Hold, ref, campus_ref,
                  Label, mfg_line, max_capacity, capacity_status, current_ss_alert, total_cust_order, cust_order_in_the_next_5_days,
                  wo_qty_in_the_next_5_days, receipt_qty_in_the_next_5_days, po_qty_in_the_next_5_days, ss_alert_after, Sku_has_ss,
                  Sku_greater_or_equal_ss, Sku_less_ss, Sku_less_ss_with_supply, priority_sku_unique, oil_allocation,
                  campus_ss, campus_total_available, campus_Sku_ss, campus_Sku_greater_equal_ss, campus_Sku_less_ss) -> ssmetrics_final


ssmetrics_final %>% 
  dplyr::mutate(MTO_MTS = ifelse(Location == 430, 4, MTO_MTS)) -> ssmetrics_final

# Final Touch
ssmetrics_final %>% 
  dplyr::filter(MTO_MTS == 4) -> ssmetrics_final



# (Path revision needed) #### weekly result ##### ----
ssmetrics_final -> ssmetrics_final_2
ssmetrics_final_2 %>% 
  dplyr::mutate(ref = gsub("_", "-", ref),
                campus_ref = gsub("_", "-", campus_ref),
                date = format(as.Date(date), "%m/%d/%y")) -> ssmetrics_final_2





###########################################################################################################################################
########################################################### Fixing the data ###############################################################
###########################################################################################################################################

# Location 430 MPF wipped out fix
ssmetrics_final_2 %>% 
  dplyr::mutate(MPF = ifelse(Location == 430, "PKG", MPF)) -> ssmetrics_final_2


# MPF error fix
ssmetrics_final_2 %>% 
  dplyr::mutate(MPF = ifelse(is.na(MPF) & Stocking_Type_Description == "Raw Material", Type, MPF)) %>% 
  dplyr::mutate(MPF = ifelse(MPF == "Packaging", "PKG",
                             ifelse(MPF == "Label", "LBL",
                                    ifelse(MPF == "Ingredients", "ING", MPF)))) -> ssmetrics_final_2



# completed sku list import (fix Category & Platform) ----
# https://edgeanalytics.venturafoods.com/MicroStrategy/servlet/mstrWeb
completed_sku_list <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/Completed SKU list - Linda.xlsx")
completed_sku_list[-1:-2, ]  %>% 
  janitor::clean_names() %>% 
  dplyr::select(x6, x9, x11) %>% 
  dplyr::rename(Item = x6,
                Category = x9,
                Platform = x11) %>% 
  dplyr::mutate(Item = gsub("-", "", Item)) -> completed_sku_list

completed_sku_list[!duplicated(completed_sku_list[,c("Item")]),] -> completed_sku_list

completed_sku_list %>% 
  dplyr::select(Item, Category) -> completed_sku_list_category


completed_sku_list %>% 
  dplyr::select(Item, Platform) -> completed_sku_list_platform



ssmetrics_final_2 %>%
  dplyr::select(-Category, -Platform) %>% 
  dplyr::left_join(completed_sku_list_category) %>% 
  dplyr::left_join(completed_sku_list_platform) %>% 
  dplyr::relocate(c(Category, Platform), .after = year) %>% 
  dplyr::mutate(Category = ifelse(Stocking_Type_Description == "Raw Material" & MPF != "PKG" & MPF != "ING" & MPF != "LBL",
                                  Type, Category)) %>% 
  dplyr::mutate(Platform = ifelse(Stocking_Type_Description == "Raw Material" & MPF != "PKG" & MPF != "ING" & MPF != "LBL",
                                  Type, Platform)) -> ssmetrics_final_2



ssmetrics_final_2 %>% 
  dplyr::mutate(Category = ifelse(Stocking_Type_Description == "Raw Material" & MPF == "PKG", "Packaging",
                                  ifelse(Stocking_Type_Description == "Raw Material" & MPF == "LBL", "Label",
                                         ifelse(Stocking_Type_Description == "Raw Material" & MPF == "ING", "Ingredients", Category)))) %>% 
  dplyr::mutate(Platform = ifelse(Stocking_Type_Description == "Raw Material" & MPF == "PKG", "Packaging",
                                  ifelse(Stocking_Type_Description == "Raw Material" & MPF == "LBL", "Label",
                                         ifelse(Stocking_Type_Description == "Raw Material" & MPF == "ING", "Ingredients", Platform)))) -> ssmetrics_final_2


ssmetrics_final_2 %>% 
  dplyr::mutate(Category = ifelse(is.na(Category) & Type == "Packaging", "Packaging",
                                  ifelse(is.na(Category) & Type == "Label", "Label",
                                         ifelse(is.na(Category) & Type == "Ingredients", "Ingridents", Category)))) %>% 
  dplyr::mutate(Platform = ifelse(is.na(Platform) & Type == "Packaging", "Packaging",
                                  ifelse(is.na(Platform) & Type == "Label", "Label",
                                         ifelse(is.na(Platform) & Type == "Ingredients", "Ingridents", Platform)))) -> ssmetrics_final_2





# Macro Platform Fix
ssmetrics_final_2 %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "Finished Goods" | is.na(macro_platform), "test", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "Packaging" | Platform == "Ingredients" | Platform == "Label", Platform, "test")) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "BOTTLES" | Platform == "JARS", "BOTTLES/JARS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "ALL SOUP BASES, FLAVORS" | Platform == "PAN COATING", "CO-PACK", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "DRUMS, TOTES", "DRUMS, TOTES", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "JUGS - W/ HANDLE", "JUGS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "LARGE POUCH > 6 oz" | Platform == "NGSD-NEXT GEN SAUCE DISPENSER", "LARGE POUCH", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "PC CUP - 37MM" | Platform == "PC CUP - 51MM" | 
                                          Platform == "PC CUP - 60MM" | Platform == "PC CUP - 75MM" | Platform == "PC CUP - RECTG" | 
                                          Platform == "PC POUCH <= 6 OZ, 3 SIDE SEAL" | Platform == "PC POUCH <= 6 OZ, 4 SIDE SEAL", "PC'S", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "PRINTS" | Platform == "QTRS", "PRINTS/QTRS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "TRAYS", "TRAYS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "TUB/BOWL - OVAL <= 5#" | Platform == "TUB/BOWL - ROUND <= 5#" | 
                                          Platform == "TUB/BOWL - SQUARE <= 5#" | Platform == "TUBS, PAILS, CARTONS > 5#", "TUB/BOWL", macro_platform)) -> ssmetrics_final_2



# MPF (PFG fix)
ssmetrics_final_2 %>% 
  dplyr::mutate(MPF = ifelse(MPF == "PFG", "PKG", MPF)) -> ssmetrics_final_2


# mfg_line raw material fix
ssmetrics_final_2 %>% 
  dplyr::mutate(mfg_line = ifelse(is.na(mfg_line) & Stocking_Type_Description == "Raw Material", Category, mfg_line)) %>% 
  dplyr::mutate(mfg_line = ifelse(mfg_line == "Packaging", "PKG",
                                  ifelse(mfg_line == "Label", "LBL",
                                         ifelse(mfg_line == "Ingredients", "ING", mfg_line)))) -> ssmetrics_final_2


# Type
ssmetrics_final_2 %>% 
  dplyr::mutate(Type = ifelse(is.na(Type) & Stocking_Type_Description == "Raw Material", Category,
                              ifelse(is.na(Type) & Stocking_Type_Description == "Make", "Finished Goods",
                                     ifelse(is.na(Type) & Stocking_Type_Description == "Transfer", "Finished Goods",
                                            ifelse(is.na(Type) & Stocking_Type_Description == "Purchased", "Finished Goods", Type))))) -> ssmetrics_final_2

# MTO/MTS
ssmetrics_final_2 %>% 
  dplyr::mutate(MTO_MTS = "MTS") -> ssmetrics_final_2 


# Dups remove
ssmetrics_final_2[!duplicated(ssmetrics_final_2[,c("ref", "oil_allocation")]),] -> ssmetrics_final_2



################### Additional Code revise 5/30/2023 ######################

ssmetrics_final_2 %>% 
  dplyr::mutate(Platform = ifelse(Platform == "Ingridents", "Ingredients", Platform)) %>% 
  dplyr::mutate(Category = ifelse(Category == "Ingridents", "Ingredients", Category)) -> ssmetrics_final_2


################### Additional Code revise 6/05/2023 ######################
ssmetrics_final_2 %>% 
  dplyr::mutate(Planner_Name = replace(Planner_Name, is.na(Planner_Name), 0)) -> ssmetrics_final_2

ssmetrics_final_2 %>% 
  dplyr::mutate(capacity_status = ifelse(Type == "Finished Goods", 
                                         ifelse(max_capacity > 0.85, "Constrained",
                                                ifelse(max_capacity < 0.75, "OK", "Check")), Type)) -> ssmetrics_final_2



################### Additional Code revise 6/08/2023 ######################
ssmetrics_final_2 %>% 
  dplyr::mutate(type_2 = ifelse(Type == "Packaging", "Raw Material", 
                                ifelse(Type == "Label", "Raw Material",
                                       ifelse(Type == "Ingredients", "Raw Material", Type))),
                stocking_type_2 = ifelse(type_2 == "Raw Material", Type, Stocking_Type_Description)) %>% 
  dplyr::select(-Type, -Stocking_Type_Description) %>% 
  dplyr::rename(Type = type_2,
                Stocking_Type_Description = stocking_type_2) %>% 
  dplyr::relocate(c(Type, Stocking_Type_Description), .after = Description) -> ssmetrics_final_2


################### Additional Code revise 7/10/2023 ######################
campus_ref %>% 
  dplyr::select(Location, Campus) %>% 
  dplyr::mutate(Location = as.double(Location),
                Campus = as.double(Campus)) -> campus_ref_2

ssmetrics_final_2 %>% 
  merge(campus_ref_2[, c("Location", "Campus")], by = "Location", all.x = TRUE)  -> ssmetrics_final_2


ssmetrics_final_2 %>% 
  dplyr::relocate(Location, .after = date) -> ssmetrics_final_2

################### Additional Code revise 8/15/2023 #######################
campus_abb %>%
  janitor::clean_names() -> campus_abb

ssmetrics_final_2 %>% 
  dplyr::select(-Campus.x) %>%
  dplyr::rename(campus = Campus.y) %>% 
  dplyr::left_join(campus_abb) -> ssmetrics_final_2

################### Additional Code revise 10/03/2023 #######################
pre_ss_metrics <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/05.28.2024/SS Metrics 0528.xlsx")

pre_ss_metrics %>% 
  dplyr::select(Item, Category, Platform, macro_platform, Type, Stocking_Type_Description, mfg_line, capacity_status) %>% 
  dplyr::rename(Category_2 = Category,
                Platform_2 = Platform,
                macro_platform_2 = macro_platform,
                Type_2 = Type,
                Stocking_Type_Description_2 = Stocking_Type_Description,
                mfg_line_2 = mfg_line,
                capacity_status_2 = capacity_status) -> pre_ss_metrics

pre_ss_metrics[!duplicated(pre_ss_metrics[,c("Item")]),] -> pre_ss_metrics

ssmetrics_final_2 %>% 
  dplyr::left_join(pre_ss_metrics) %>% 
  dplyr::mutate(Category = ifelse(is.na(Category), Category_2, Category),
                Platform = ifelse(is.na(Platform), Platform_2, Platform),
                macro_platform = ifelse(is.na(macro_platform), macro_platform_2, macro_platform),
                Type = ifelse(is.na(Type), Type_2, Type),
                Stocking_Type_Description = ifelse(is.na(Stocking_Type_Description), Stocking_Type_Description_2, Stocking_Type_Description),
                mfg_line = ifelse(is.na(mfg_line), mfg_line_2, mfg_line),
                capacity_status = ifelse(is.na(capacity_status), capacity_status_2, capacity_status)) %>% 
  dplyr::select(-Category_2, -Platform_2, -macro_platform_2, -Type_2, -Stocking_Type_Description_2, -mfg_line_2, -capacity_status_2) -> ssmetrics_final_2


################### Additional Code revise 10/24/2023 #######################
iqr_fg <- readxl::read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/weekly run data/2024/05.28.2024/Finished Goods Inventory Health Adjusted Forward (IQR) NEW TEMPLATE - 05.28.2024.xlsx",
                             sheet = "Location FG")


iqr_fg %>% 
  data.frame() %>% 
  dplyr::slice(-1:-2) -> iqr_fg_mfg_site

colnames(iqr_fg_mfg_site) <- iqr_fg_mfg_site[1, ]

iqr_fg_mfg_site %>% 
  dplyr::slice(-1) %>% 
  janitor::clean_names() %>% 
  dplyr::select(ref, mfg_ref) %>% 
  tidyr::separate(mfg_ref, c("mfg_site", "item"), sep = "-") %>% 
  dplyr::select(-item)  -> iqr_fg_mfg_site


ssmetrics_final_2 %>% 
  dplyr::left_join(iqr_fg_mfg_site) %>% 
  dplyr::mutate(mfg_site = ifelse(is.na(mfg_site), "NA", mfg_site)) %>% 
  dplyr::relocate(mfg_site, .before = date) -> ssmetrics_final_2



####### 11/1/2023 #######
ssmetrics_final_2 %>% 
  dplyr::mutate(Label = ifelse(is.na(Label), "NA", Label)) -> ssmetrics_final_2


####### 01/16/2024 #######

ssmetrics_final_2 %>% 
  dplyr::mutate(campus = ifelse(Location == "430", 43, campus),
                campus_name = ifelse(Location == "430", "BHM", campus_name)) %>% 
  dplyr::mutate(campus_ref = paste0(campus, "_", Item)) -> ssmetrics_final_2


####### 02/20/2024 #######
ssmetrics_final_2 %>% 
  dplyr::relocate(c(campus, campus_name), .after = Location_Name) -> ssmetrics_final_2


####### 03/06/2024 #######

ssmetrics_final_2 %>%
  left_join(ssmetrics_pre_1 %>% select(Category, Item), by = "Item") %>%
  dplyr::mutate(Category = ifelse(is.na(Category.x), Category.y, Category.x)) %>%
  dplyr::select(-Category.y, -Category.x) %>% 
  dplyr::relocate(c(Category), .after = year) -> ssmetrics_final_2

ssmetrics_final_2 %>%
  left_join(ssmetrics_pre_1 %>% select(Platform, Item), by = "Item") %>%
  dplyr::mutate(Platform = ifelse(is.na(Platform.x), Platform.y, Platform.x)) %>%
  dplyr::select(-Platform.y, -Platform.x) %>% 
  dplyr::relocate(c(Platform), .after = Category) -> ssmetrics_final_2




# Macro Platform Fix
ssmetrics_final_2 %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "Finished Goods" | is.na(macro_platform), "test", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "Packaging" | Platform == "Ingredients" | Platform == "Label", Platform, "test")) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "BOTTLES" | Platform == "JARS", "BOTTLES/JARS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "ALL SOUP BASES, FLAVORS" | Platform == "PAN COATING", "CO-PACK", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "DRUMS, TOTES", "DRUMS, TOTES", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "JUGS - W/ HANDLE", "JUGS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "LARGE POUCH > 6 oz" | Platform == "NGSD-NEXT GEN SAUCE DISPENSER", "LARGE POUCH", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "PC CUP - 37MM" | Platform == "PC CUP - 51MM" | 
                                          Platform == "PC CUP - 60MM" | Platform == "PC CUP - 75MM" | Platform == "PC CUP - RECTG" | 
                                          Platform == "PC POUCH <= 6 OZ, 3 SIDE SEAL" | Platform == "PC POUCH <= 6 OZ, 4 SIDE SEAL", "PC'S", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "PRINTS" | Platform == "QTRS", "PRINTS/QTRS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "TRAYS", "TRAYS", macro_platform)) %>% 
  dplyr::mutate(macro_platform = ifelse(macro_platform == "test" | Platform == "TUB/BOWL - OVAL <= 5#" | Platform == "TUB/BOWL - ROUND <= 5#" | 
                                          Platform == "TUB/BOWL - SQUARE <= 5#" | Platform == "TUBS, PAILS, CARTONS > 5#", "TUB/BOWL", macro_platform)) -> ssmetrics_final_2



ssmetrics_final_2 %>%
  left_join(ssmetrics_pre_1 %>% select(Type, Item), by = "Item") %>%
  dplyr::mutate(Type = ifelse(is.na(Type.x), Type.y, Type.x)) %>%
  dplyr::select(-Type.y, -Type.x) %>% 
  dplyr::relocate(c(Type), .after = Description) -> ssmetrics_final_2

ssmetrics_final_2 %>%
  left_join(ssmetrics_pre_1 %>% select(Stocking_type_description, Item) %>% rename(Stocking_Type_Description = Stocking_type_description), by = "Item") %>%
  dplyr::mutate(Stocking_Type_Description = ifelse(is.na(Stocking_Type_Description.x), Stocking_Type_Description.y, Stocking_Type_Description.x)) %>%
  dplyr::select(-Stocking_Type_Description.y, -Stocking_Type_Description.x) %>% 
  dplyr::relocate(c(Stocking_Type_Description), .after = Type) -> ssmetrics_final_2

ssmetrics_final_2 %>%
  mutate(mfg_line = ifelse(is.na(mfg_line), MPF, mfg_line))  -> ssmetrics_final_2






ssmetrics_final_2 %>%
  mutate(
    Category = ifelse(is.na(Category) & startsWith(Description, "RSC"), "packaging", Category),
    Platform = ifelse(is.na(Platform) & startsWith(Description, "RSC"), "packaging", Platform),
    macro_platform = ifelse(is.na(macro_platform) & startsWith(Description, "RSC"), "packaging", macro_platform),
    Stocking_Type_Description = ifelse(is.na(Stocking_Type_Description) & startsWith(Description, "RSC"), "packaging", Stocking_Type_Description),
    capacity_status = ifelse(is.na(capacity_status) & startsWith(Description, "RSC"), "packaging", capacity_status),
    oil_allocation = ifelse(is.na(oil_allocation) & startsWith(Description, "RSC"), "packaging", oil_allocation),
    Type = ifelse(is.na(Type) & startsWith(Description, "RSC"), "Raw Material", Type),
    MPF = ifelse(is.na(MPF) & startsWith(Description, "RSC"), "PKG", MPF),
    mfg_line = ifelse(is.na(mfg_line) & startsWith(Description, "RSC"), "PKG", mfg_line)
  ) -> ssmetrics_final_2


ssmetrics_final_2 %>%
  left_join(exception_report %>% 
              select(ItemNo, Description) %>% 
              rename(Item = ItemNo) %>% 
              distinct(), 
            by = "Item") %>%
  mutate(Description = ifelse(is.na(Description.x), Description.y, Description.x)) %>%
  select(-Description.x, -Description.y) %>% 
  relocate(Description, .after = Item) -> ssmetrics_final_2

####### 03/06/2024 #######
ssmetrics_final_2 %>% 
  dplyr::mutate(campus_ss_on_hand = ifelse(campus_total_available > campus_ss, campus_ss, campus_total_available)) %>% 
  dplyr::relocate(campus_ss_on_hand, .after = campus_total_available)-> ssmetrics_final_2

ssmetrics_final_2 %>% 
  dplyr::mutate(type_2 = ifelse(Type == "Packaging", "Raw Material", 
                                ifelse(Type == "Label", "Raw Material",
                                       ifelse(Type == "Ingredients", "Raw Material", Type))),
                stocking_type_2 = ifelse(type_2 == "Raw Material", Category, Stocking_Type_Description)) %>% 
  dplyr::select(-Type, -Stocking_Type_Description) %>% 
  dplyr::rename(Type = type_2,
                Stocking_Type_Description = stocking_type_2) %>% 
  dplyr::relocate(c(Type, Stocking_Type_Description), .after = Description) -> ssmetrics_final_2

ssmetrics_final_2 %>% 
  dplyr::mutate(campus_ref = gsub("_", "-", campus_ref)) -> ssmetrics_final_2

#####################################################################################################################################
#####################################################################################################################################
#####################################################################################################################################


writexl::write_xlsx(ssmetrics_final_2, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/SS Metrics 0604.xlsx") 



#####################################################################################################################################
###################################################### save & update mega data ######################################################
#####################################################################################################################################

ssmetrics_final_2 %>% 
  dplyr::mutate(ref = gsub("-", "_", ref),
                campus_ref = gsub("-", "_", campus_ref)) -> ssmetrics_final


colnames(ssmetrics_mainboard) <- colnames(ssmetrics_final)

readr::type_convert(ssmetrics_mainboard) -> ssmetrics_mainboard






ssmetrics_mainboard %>% 
  dplyr::mutate(ref = gsub("_", "-", ref),
                campus_ref = gsub("_", "-", campus_ref)) -> ssmetrics_mainboard



# (Path revision needed) ----
save(ssmetrics_mainboard, file = "C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Safety_Stock_Compliance/RPA/venturafoods_SafetyStockCompliance_RPA/rds files/ssmetrics_mainboard_06_04_2024.rds")




#######################################################################################################################################
#######################################################################################################################################
#######################################################################################################################################

file.copy("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/05.28.2024/Weekly Safety Stock Compliance Report v4 rolling 53 weeks - 05.28.2024.xlsb",
          "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Weekly Run Files/2024/06.04.2024/Weekly Safety Stock Compliance Report v4 rolling 53 weeks - 06.04.2024.xlsb")
