library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)

# ssmetrics_mainboard <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/ssmetrics_main_board.xlsx",
#                          col_names = FALSE)

# load main board (mega data) ----
load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Safety_Stock_Compliance/RPA/venturafoods_SafetyStockCompliance_RPA/rds files/ssmetrics_mainboard.rds")

colnames(ssmetrics_mainboard) <- ssmetrics_mainboard[1, ]
ssmetrics_mainboard[-1, ] -> ssmetrics_mainboard
names(ssmetrics_mainboard) <- str_replace_all(names(ssmetrics_mainboard), c(" " = "_"))
names(ssmetrics_mainboard) <- str_replace_all(names(ssmetrics_mainboard), c("/" = "_"))
names(ssmetrics_mainboard) <- str_replace_all(names(ssmetrics_mainboard), c("-" = "_"))


ssmetrics_mainboard %>% 
  dplyr::mutate(Ref = gsub("-", "_", Ref),
                Campus_Ref = gsub("-", "_", Campus_Ref)) -> ssmetrics_mainboard

readr::type_convert(ssmetrics_mainboard) -> ssmetrics_mainboard

colnames(ssmetrics_mainboard)[8] <- "campus"
colnames(ssmetrics_mainboard)[9] <- "date"
colnames(ssmetrics_mainboard)[12] <- "Description"
colnames(ssmetrics_mainboard)[14] <- "Stocking_Type_Description"
colnames(ssmetrics_mainboard)[17] <- "MPF_Line"
colnames(ssmetrics_mainboard)[18] <- "Safety_Stock"
colnames(ssmetrics_mainboard)[19] <- "Balance_Usable"
colnames(ssmetrics_mainboard)[20] <- "Balance_Hold"
colnames(ssmetrics_mainboard)[21] <- "ref"
colnames(ssmetrics_mainboard)[22] <- "campus_ref"



############################### Phase 1 ############################
# Stock type ----
load("C:/Users/slee/OneDrive - Ventura Foods/Stan/R Codes/Projects/Safety_Stock_Compliance/RPA/venturafoods_SafetyStockCompliance_RPA/rds files/stock_type.rds")

# Macro-platform (change this only when there's a change) ----
macro_platform <- read_excel("S:/Supply Chain Projects/RStudio/Macro-platform.xlsx",
                             col_names = FALSE)

colnames(macro_platform) <- macro_platform[1, ]
macro_platform[-1, ] -> macro_platform

colnames(macro_platform)[2] <- "macro_platform"

# Location_Name (change this only when there's a change) ----
location_name <- read_excel("S:/Supply Chain Projects/RStudio/Location_Name.xlsx",
                             col_names = FALSE)

colnames(location_name) <- location_name[1, ]
location_name[-1, ] -> location_name

location_name %>% 
  dplyr::mutate(Location = as.numeric(Location)) -> location_name

# priority_Sku_uniques (change this only when there's a change) ----
priority_sku <- read_excel("S:/Supply Chain Projects/RStudio/Priority_Sku_and_uniques.xlsx",
                             col_names = FALSE)

colnames(priority_sku) <- priority_sku[1, ]
priority_sku[-1, ] -> priority_sku

colnames(priority_sku)[1] <- "priority_sku"

priority_sku %>% 
  dplyr::mutate(Item = priority_sku) -> priority_sku

# oil allocation (change this only when there's a change) ----
oil_aloc <- read_excel("S:/Supply Chain Projects/RStudio/oil allocation.xlsx",
                           col_names = FALSE)

colnames(oil_aloc) <- oil_aloc[1, ]
oil_aloc[-1, ] -> oil_aloc

colnames(oil_aloc)[1] <- "Item"
colnames(oil_aloc)[2] <- "oil_aloc"
colnames(oil_aloc)[3] <- "comp_desc"


# Inventory Model  (Make sure to remove the password of the original .xlsx file) ----

inventory_model <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/SS Optimization by Location - Finished Goods July 2022.xlsx",
                              col_names = FALSE, sheet = "Fin Goods")

inventory_model[-1:-7, ] -> inventory_model
colnames(inventory_model) <- inventory_model[1, ]
inventory_model[-1, ] -> inventory_model

colnames(inventory_model)[5] <- "ref"
colnames(inventory_model)[17] <- "mfg_line"
colnames(inventory_model)[32] <- "max_capacity"

inventory_model %>% 
  dplyr::select(ref, mfg_line, max_capacity) %>% 
  dplyr::mutate(ref = gsub("-", "_", ref)) %>% 
  dplyr::mutate(max_capacity = as.numeric(max_capacity)) -> inventory_model



# Campus reference
campus_ref <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/campus reference.xlsx",
                         col_names = FALSE)

colnames(campus_ref) <- campus_ref[1, ]
campus_ref[-1, ] -> campus_ref

colnames(campus_ref)[1] <- "Location"
colnames(campus_ref)[3] <- "Campus"
colnames(campus_ref)[4] <- "campus_no"

campus_ref %>% 
  dplyr::mutate(Campus = replace(Campus, is.na(Campus), 0)) -> campus_ref



# Lot Status
Lot_Status <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Lot Status Code.xlsx",
                         col_names = FALSE)

colnames(Lot_Status) <- Lot_Status[1, ]
Lot_Status[-1, ] -> Lot_Status

Lot_Status %>% 
  dplyr::rename(Lot_Status = "Lot status",
                Hold_Status = "Hard/Soft Hold") %>% 
  dplyr::select(Lot_Status, Hold_Status) -> Lot_Status

# previous SS_Metrics file ----
ssmetrics_pre <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/Copy of Safety Stock Compliance Report Data v3 - 06.20.22.xlsx",
                             col_names = FALSE)

ssmetrics_pre[-1, ] -> ssmetrics_pre
colnames(ssmetrics_pre) <- ssmetrics_pre[1, ]
ssmetrics_pre[-1, ] -> ssmetrics_pre
names(ssmetrics_pre) <- str_replace_all(names(ssmetrics_pre), c(" " = "_"))
names(ssmetrics_pre) <- str_replace_all(names(ssmetrics_pre), c("/" = "_"))


# Planner_address Change Directory only when you need to ----
Planner_address <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/IQR Automation/FG/FG test/Address Book - 06.08.22.xlsx", 
                              sheet = "Sheet1", col_types = c("text", 
                                                              "text", "text", "text", "text"))

names(Planner_address) <- str_replace_all(names(Planner_address), c(" " = "_"))

colnames(Planner_address)[1] <- "Planner_No"

Planner_address %>% 
  dplyr::select(1:2) -> Planner_address

# JDE VF Item Branch - Work with Item Branch ----
JD_item_branch <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/86, 226, 381 Sheet 4 - Copy.xlsx",
                             col_names = FALSE)

colnames(JD_item_branch) <- JD_item_branch[1, ]
JD_item_branch[-1, ] -> JD_item_branch

names(JD_item_branch) <- str_replace_all(names(JD_item_branch), c(" " = "_"))
names(JD_item_branch) <- str_replace_all(names(JD_item_branch), c("/" = "_"))
names(JD_item_branch) <- str_replace_all(names(JD_item_branch), c("2" = "second"))

colnames(JD_item_branch)[1] <- "Location"
colnames(JD_item_branch)[2] <- "Item"

readr::type_convert(JD_item_branch) -> JD_item_branch

JD_item_branch %>%
  dplyr::mutate(ref = paste0(Location, "_", Item)) -> JD_item_branch

# exception report ----
exception_report <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/exception report 06.20.22 - Copy.xlsx", 
                               sheet = "Sheet1",
                               col_types = c("text", "text", "text", 
                                             "text", "numeric", "text", "text", "text", 
                                             "text", "text", "text", "text", "text", 
                                             "text", "numeric", "numeric", "numeric", 
                                             "numeric", "numeric", "numeric", 
                                             "numeric", "text", "text", "text", 
                                             "text", "text", "text", "text", "numeric", 
                                             "text", "text", "text"))

exception_report[-1:-2, -32] -> exception_report

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

names(exception_report) <- str_replace_all(names(exception_report), c(" " = "_"))


exception_report %<>% 
  dplyr::mutate(ref = paste0(B_P, "_", ItemNo)) %>% 
  dplyr::relocate(ref) 


# custord custord ----
# Open Customer Order File pulling ----  Change Directory ----
custord <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/wo receipt custord po - 06.20.22 - Copy.xlsx", 
                      sheet = "custord", col_names = FALSE)



custord %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8", "9"), sep = "~") %>% 
  dplyr::select(-"3", -"6", -"7", -"8") -> custord

custord %>% 
  dplyr::rename(aa = "1") %>% 
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp, -"4") %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                date = "9") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item),
                in_next_7_days = ifelse(date >= Sys.Date() & date < Sys.Date() +7, "Y", "N")) %>% 
  dplyr::relocate(ref, Item, Location, in_next_7_days) -> custord

# Custord pivot
reshape2::dcast(custord, ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(total_cust_order = N + Y) -> custord_pivot



# Custord wo ----
wo <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/wo receipt custord po - 06.20.22 - Copy.xlsx", 
                 sheet = "wo", col_names = FALSE)


wo %>% 
  dplyr::rename(aa = "...1") %>% 
  tidyr::separate(aa, c("1", "2", "3", "4", "5", "6", "7", "8"), sep = "~") %>% 
  dplyr::select(-"3") %>% 
  dplyr::rename(aa = "1") %>%  
  tidyr::separate(aa, c("global", "rp", "Item")) %>% 
  dplyr::select(-global, -rp, -"4", -"8") %>% 
  dplyr::rename(Location = "2",
                Qty = "5",
                wo_no = "6",
                date = "7") %>% 
  dplyr::mutate(Qty = as.double(Qty),
                date = as.Date(date)) %>% 
  dplyr::mutate(Location = sub("^0+", "", Location)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref, Item, Location) %>% 
  dplyr::mutate(in_next_7_days = ifelse(date >= Sys.Date() & date < Sys.Date()+7, "Y", "N") )-> wo

# wo pivot
wo %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(N = as.integer(N)) -> wo_pivot



# Custord Receipt ----
receipt <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/wo receipt custord po - 06.20.22 - Copy.xlsx", 
                      sheet = "receipt", col_names = FALSE)


receipt %>% 
  dplyr::rename(aa = "...1") %>% 
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
  dplyr::mutate(in_next_7_days = ifelse(date >= Sys.Date() & date < Sys.Date()+7, "Y", "N") ) -> receipt



# Receipt Pivot
receipt %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) -> receipt_pivot  


# Custord PO ----
po <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/wo receipt custord po - 06.20.22 - Copy.xlsx", 
                 sheet = "po", col_names = FALSE)

po %>% 
  dplyr::rename(aa = "...1") %>% 
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
  dplyr::mutate(in_next_7_days = ifelse(date >= Sys.Date() & date< Sys.Date() + 7, "Y", "N") ) -> po


# PO_Pivot 
po %>% 
  reshape2::dcast(ref ~ in_next_7_days, value.var = "Qty", sum) %>% 
  dplyr::mutate(N = as.integer(N),
                Y = as.integer(Y)) -> PO_Pivot



# JD - OH ----
JDOH <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/Copy of JD_OH_SS_20220620 - Copy.xlsx", 
                   sheet = "itmbal", col_names = FALSE)


JDOH[-1, ] -> JDOH
colnames(JDOH) <- JDOH[1, ]
JDOH[-1, ] -> JDOH

colnames(JDOH)[1] <- "Location"
colnames(JDOH)[2] <- "Item"
colnames(JDOH)[3] <- "Stock_Type"
colnames(JDOH)[4] <- "Description"
colnames(JDOH)[5] <- "Balance_Usable"
colnames(JDOH)[6] <- "Balance_Hold"
colnames(JDOH)[7] <- "Lot_Status"
colnames(JDOH)[8] <- "On_Hand"
colnames(JDOH)[9] <- "Safety_Stock"
colnames(JDOH)[10] <- "GL_Class"
colnames(JDOH)[11] <- "Planner_No"
colnames(JDOH)[12] <- "Planner_Name"

readr::type_convert(JDOH) -> JDOH

JDOH %>% 
  dplyr::filter(Location != 86 & Location!= 226 & Location != 381) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item)) %>% 
  dplyr::relocate(ref) -> JDOH


# Inventory Analysis for all locations ----
Inv_FG <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/Inventory Report for all locations - 06.20.22 - Copy.xlsx", 
                     sheet = "FG", col_names = FALSE)

Inv_RM <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/Inventory Report for all locations - 06.20.22 - Copy.xlsx", 
                     sheet = "RM", col_names = FALSE)


Inv_FG[-1:-3, ] -> Inv_FG
Inv_RM[-1:-2, ] -> Inv_RM

colnames(Inv_RM) <- Inv_RM[1, ]
Inv_RM[-1, ] -> Inv_RM

colnames(Inv_RM)[2] <- "Location_Name"
colnames(Inv_RM)[5] <- "Description"

names(Inv_RM) <- str_replace_all(names(Inv_RM), c(" " = "_"))

colnames(Inv_FG)[1] <- "Location"
colnames(Inv_FG)[2] <- "Location_Name"
colnames(Inv_FG)[3] <- "Campus"
colnames(Inv_FG)[4] <- "Item"
colnames(Inv_FG)[5] <- "Description"
colnames(Inv_FG)[6] <- "Inventory_Status_Code"
colnames(Inv_FG)[7] <- "Hold_Status"
colnames(Inv_FG)[8] <- "Current_Inventory_Balance"

dplyr::bind_rows(Inv_RM, Inv_FG) -> Inv_all

Inv_all %>% 
  dplyr::mutate(Item = gsub("-", "", Item)) %>% 
  dplyr::mutate(ref = paste0(Location, "_", Item),
                mfg_ref = paste0(Campus, "_", Item)) %>% 
  dplyr::relocate(ref, mfg_ref) -> Inv_all

readr::type_convert(Inv_all) -> Inv_all

Inv_all %>% 
  dplyr::mutate(item_length = nchar(Item)) -> Inv_all

Inv_all %>% filter(item_length == 8) -> fg 
Inv_all %>% filter(item_length == 5 | item_length == 9) -> rm

rm %>% 
  dplyr::mutate(Item = sub("^0+", "", Item)) -> rm

rbind(rm, fg) -> Inv_all
Inv_all %>% 
  dplyr::select(-item_length) -> Inv_all


############################################################################################################################
########################### From here, This should be activated after two locations are resolved ###########################
############################################################################################################################

# Inv_all %>%
#   reshape2::dcast(Location + Item + Description ~ Hold_Status, value.var = "Current_Inventory_Balance", sum) %>%
#   dplyr::rename(Balance_Usable = Useable,
#                 Hard_Hold = "Hard Hold",
#                 Soft_Hold = "Soft Hold") %>%
#   dplyr::mutate(Stock_Type = "",
#                 Balance_Hold = Hard_Hold + Soft_Hold,
#                 Lot_Status = "",
#                 On_Hand = "",
#                 Safety_Stock = "",
#                 GL_Class = "",
#                 Planner_No = "",
#                 Planner_Name = "",
#                 ref = paste0(Location, "_", Item)) %>%
#   dplyr::select(-Hard_Hold, -Soft_Hold) %>%
#   dplyr::relocate(Location, Item, Stock_Type, Description, Balance_Usable, Balance_Hold)-> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Lot Status
# Inv_all_pivot_for_JDOH %>%
#   merge(Inv_all[, c("ref", "Inventory_Status_Code")], by = "ref", all.x = TRUE) %>%
#   relocate(Inventory_Status_Code, .after = Lot_Status) %>%
#   dplyr::select(-Lot_Status) %>%
#   dplyr::rename(Lot_Status = Inventory_Status_Code) %>%
#   dplyr::relocate(ref) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Stock_Type
# Inv_all_pivot_for_JDOH %>%
#   merge(JD_item_branch[, c("ref", "Stocking_Type")], by = "ref", all.x = TRUE) %>%
#   dplyr::relocate(Stocking_Type, .after = Stock_Type) %>%
#   dplyr::select(-Stock_Type) %>%
#   dplyr::rename(Stock_Type = Stocking_Type) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Safety_stock
# Inv_all_pivot_for_JDOH %>%
#   merge(exception_report[, c("ref", "Safety_Stock")], by = "ref", all.x = TRUE) %>%
#   dplyr::relocate(Safety_Stock.y, .after = Safety_Stock.x) %>%
#   dplyr::select(-Safety_Stock.x) %>%
#   dplyr::rename(Safety_Stock = Safety_Stock.y) %>%
#   dplyr::mutate(Safety_Stock = replace(Safety_Stock, is.na(Safety_Stock), 0)) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Planner_No
# Inv_all_pivot_for_JDOH %>%
#   merge(exception_report[, c("ref", "Planner")], by = "ref", all.x = TRUE) %>%
#   dplyr::relocate(Planner, .after = Planner_No) %>%
#   dplyr::select(-Planner_No) %>%
#   dplyr::rename(Planner_No = Planner) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Planner_Name
# Inv_all_pivot_for_JDOH %>%
#   merge(Planner_address[, c("Planner_No", "Alpha_Name")], by = "Planner_No", all.x = TRUE) %>%
#   dplyr::relocate(Alpha_Name, .after = Planner_Name) %>%
#   dplyr::select(-Planner_Name) %>%
#   dplyr::rename(Planner_Name = Alpha_Name) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Lot Status
# Inv_all_pivot_for_JDOH %>%
#   merge(Inv_all[, c("ref", "Inventory_Status_Code")], by = "ref", all.x = TRUE) %>%
#   relocate(Inventory_Status_Code, .after = Lot_Status) %>%
#   dplyr::select(-Lot_Status) %>%
#   dplyr::rename(Lot_Status = Inventory_Status_Code) %>%
#   dplyr::relocate(ref) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Stock_Type
# Inv_all_pivot_for_JDOH %>%
#   merge(JD_item_branch[, c("ref", "Stocking_Type")], by = "ref", all.x = TRUE) %>%
#   dplyr::relocate(Stocking_Type, .after = Stock_Type) %>%
#   dplyr::select(-Stock_Type) %>%
#   dplyr::rename(Stock_Type = Stocking_Type) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Safety_stock
# Inv_all_pivot_for_JDOH %>%
#   merge(exception_report[, c("ref", "Safety_Stock")], by = "ref", all.x = TRUE) %>%
#   dplyr::relocate(Safety_Stock.y, .after = Safety_Stock.x) %>%
#   dplyr::select(-Safety_Stock.x) %>%
#   dplyr::rename(Safety_Stock = Safety_Stock.y) %>%
#   dplyr::mutate(Safety_Stock = replace(Safety_Stock, is.na(Safety_Stock), 0)) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Planner_No
# Inv_all_pivot_for_JDOH %>%
#   merge(exception_report[, c("ref", "Planner")], by = "ref", all.x = TRUE) %>%
#   dplyr::relocate(Planner, .after = Planner_No) %>%
#   dplyr::select(-Planner_No) %>%
#   dplyr::rename(Planner_No = Planner) -> Inv_all_pivot_for_JDOH
# 
# # Inv_all_pivot_for_JDOH - Planner_Name
# Inv_all_pivot_for_JDOH %>%
#   merge(Planner_address[, c("Planner_No", "Alpha_Name")], by = "Planner_No", all.x = TRUE) %>%
#   dplyr::relocate(Alpha_Name, .after = Planner_Name) %>%
#   dplyr::select(-Planner_Name) %>%
#   dplyr::rename(Planner_Name = Alpha_Name) -> Inv_all_pivot_for_JDOH
# 
# rbind(JDOH, Inv_all_pivot_for_JDOH) -> JDOH_complete
# 
# # JDOH_complete - On_Hand
# JDOH_complete %>%
#   dplyr::mutate(On_Hand = Balance_Usable + Balance_Hold) -> JDOH_complete

# ###############################################################################################################################
# ######################################## From here, this should be deactivated after two location resolved ####################
# ###############################################################################################################################
# 
# Change directory (MicroStrategy Inventory Analysis from Cassandra) ----
Inv_cassandra <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/Inventory Analysis (Cassandra) - Copy.xlsx",
                            col_names = FALSE)

Inv_cassandra[-1:-2, ] -> Inv_cassandra
colnames(Inv_cassandra) <- Inv_cassandra[1, ]
Inv_cassandra[-1, c(2:5, 13:14, 24)] -> Inv_cassandra

Location_temp <- 226
campus_temp <- 86
tibble(Location_temp, campus_temp) -> loc226_campus
loc226_campus %>%
  dplyr::rename(Location = Location_temp,
                Campus = campus_temp) -> loc226_campus

Inv_cassandra %>%
  dplyr::filter(Location == "226" | Location == "381") %>%
  dplyr::mutate(Location = as.numeric(Location)) %>%
  dplyr::left_join(loc226_campus, by = "Location") %>%
  dplyr::rename(Item = Sku,
                Location_Name = "Location Nm",
                Inventory_Status_Code = "Inventory Status",
                Hold_Status = "Inventory Hold Status",
                Current_Inventory_Balance = "Inventory Qty (Cases)") %>%
  dplyr::mutate(Item = gsub("-", "", Item),
                ref = paste0(Location, "_", Item),
                mfg_ref = paste0(Campus, "_", Item)) %>%
  dplyr::relocate(ref, mfg_ref, Location, Location_Name, Campus, Item, Description, Inventory_Status_Code,
                  Hold_Status, Current_Inventory_Balance) -> Inv_cassandra

rbind(Inv_all, Inv_cassandra) -> Inv_all


readr::type_convert(Inv_all) -> Inv_all

# Inv_all_pivot
Inv_all %>%
  reshape2::dcast(Location + Item + Description ~ Hold_Status, value.var = "Current_Inventory_Balance", sum) -> Inv_all_pivot

Inv_all %>%
  dplyr::filter(Location == 86 | Location == 226 | Location == 381) -> Inv_all_86_226_381

Inv_all_86_226_381 %>%
  reshape2::dcast(Location + Item + Description ~ Hold_Status, value.var = "Current_Inventory_Balance", sum) -> Inv_all_pivot_86_226_381

Inv_all_pivot_86_226_381 %>%
  dplyr::rename(Balance_Usable = Useable,
                Hard_Hold = "Hard Hold",
                Soft_Hold = "Soft Hold") %>%
  dplyr::mutate(Stock_Type = "",
                Balance_Hold = Hard_Hold + Soft_Hold,
                Lot_Status = "",
                On_Hand = "",
                Safety_Stock = "",
                GL_Class = "",
                Planner_No = "",
                Planner_Name = "",
                ref = paste0(Location, "_", Item)) %>%
  dplyr::select(-Hard_Hold, -Soft_Hold) %>%
  dplyr::relocate(Location, Item, Stock_Type, Description, Balance_Usable, Balance_Hold)  -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Lot Status
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(Inv_all_86_226_381[, c("ref", "Inventory_Status_Code")], by = "ref", all.x = TRUE) %>%
  relocate(Inventory_Status_Code, .after = Lot_Status) %>%
  dplyr::select(-Lot_Status) %>%
  dplyr::rename(Lot_Status = Inventory_Status_Code) %>%
  dplyr::relocate(ref) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Stock_Type
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(JD_item_branch[, c("ref", "Stocking_Type")], by = "ref", all.x = TRUE) %>%
  dplyr::relocate(Stocking_Type, .after = Stock_Type) %>%
  dplyr::select(-Stock_Type) %>%
  dplyr::rename(Stock_Type = Stocking_Type) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Safety_stock
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(exception_report[, c("ref", "Safety_Stock")], by = "ref", all.x = TRUE) %>%
  dplyr::relocate(Safety_Stock.y, .after = Safety_Stock.x) %>%
  dplyr::select(-Safety_Stock.x) %>%
  dplyr::rename(Safety_Stock = Safety_Stock.y) %>%
  dplyr::mutate(Safety_Stock = replace(Safety_Stock, is.na(Safety_Stock), 0)) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Planner_No
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(exception_report[, c("ref", "Planner")], by = "ref", all.x = TRUE) %>%
  dplyr::relocate(Planner, .after = Planner_No) %>%
  dplyr::select(-Planner_No) %>%
  dplyr::rename(Planner_No = Planner) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Planner_Name
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(Planner_address[, c("Planner_No", "Alpha_Name")], by = "Planner_No", all.x = TRUE) %>%
  dplyr::relocate(Alpha_Name, .after = Planner_Name) %>%
  dplyr::select(-Planner_Name) %>%
  dplyr::rename(Planner_Name = Alpha_Name) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Lot Status
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(Inv_all_86_226_381[, c("ref", "Inventory_Status_Code")], by = "ref", all.x = TRUE) %>%
  relocate(Inventory_Status_Code, .after = Lot_Status) %>%
  dplyr::select(-Lot_Status) %>%
  dplyr::rename(Lot_Status = Inventory_Status_Code) %>%
  dplyr::relocate(ref) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Stock_Type
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(JD_item_branch[, c("ref", "Stocking_Type")], by = "ref", all.x = TRUE) %>%
  dplyr::relocate(Stocking_Type, .after = Stock_Type) %>%
  dplyr::select(-Stock_Type) %>%
  dplyr::rename(Stock_Type = Stocking_Type) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Safety_stock
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(exception_report[, c("ref", "Safety_Stock")], by = "ref", all.x = TRUE) %>%
  dplyr::relocate(Safety_Stock.y, .after = Safety_Stock.x) %>%
  dplyr::select(-Safety_Stock.x) %>%
  dplyr::rename(Safety_Stock = Safety_Stock.y) %>%
  dplyr::mutate(Safety_Stock = replace(Safety_Stock, is.na(Safety_Stock), 0)) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Planner_No
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(exception_report[, c("ref", "Planner")], by = "ref", all.x = TRUE) %>%
  dplyr::relocate(Planner, .after = Planner_No) %>%
  dplyr::select(-Planner_No) %>%
  dplyr::rename(Planner_No = Planner) -> Inv_all_pivot_86_226_381_for_JDOH

# Inv_all_pivot_86_226_381_for_JDOH - Planner_Name
Inv_all_pivot_86_226_381_for_JDOH %>%
  merge(Planner_address[, c("Planner_No", "Alpha_Name")], by = "Planner_No", all.x = TRUE) %>%
  dplyr::relocate(Alpha_Name, .after = Planner_Name) %>%
  dplyr::select(-Planner_Name) %>%
  dplyr::rename(Planner_Name = Alpha_Name) -> Inv_all_pivot_86_226_381_for_JDOH



# combine with JDOH & Inv_all_pivot_86_226_381
# Lot Status - vlookup from Inv_all_86_226_381
rbind(JDOH, Inv_all_pivot_86_226_381_for_JDOH) -> JDOH_complete

# JDOH_complete - On_Hand
JDOH_complete %>%
  dplyr::mutate(On_Hand = Balance_Usable + Balance_Hold) -> JDOH_complete

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
  dplyr::mutate(date = Sys.Date()) %>% 
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

# MTO - 4
ssmetrics %>% 
  dplyr::filter(MTO_MTS == 4) %>% 
  dplyr::filter(Type %in% c("Finished Goods", "Ingredients", "Label", "Packaging")) -> ssmetrics

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

rbind(ssmetrics, ssmetrics_2) -> ssmetrics



#####################################################
# what if still N/A? that's new items

# this is where you will need to read Micro Strategy from Linda's access grant
types_for_na <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/.xlsx",
                                 col_names = FALSE)





# Categories vlookup from mega data
ssmetrics_mainboard %>% 
  dplyr::select(Item, Category) -> ssmetrics_mainboard_cat

ssmetrics_mainboard_cat[!duplicated(ssmetrics_mainboard_cat[,c("Item", "Category")]),] -> ssmetrics_mainboard_cat
ssmetrics_mainboard_cat[-which(duplicated(ssmetrics_mainboard_cat$Item)),] -> ssmetrics_mainboard_cat

merge(ssmetrics, ssmetrics_mainboard_cat[, c("Item", "Category")], by = "Item", all.x = TRUE) -> ssmetrics  

# here you have N/A for new SKUs. need MS ----
category_platform_for_new <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/Category and Platform and pack size.xlsx",
                           col_names = FALSE)

category_platform_for_new[-1:-2, ] -> category_platform_for_new
colnames(category_platform_for_new) <- category_platform_for_new[1, ]
category_platform_for_new[-1, ] -> category_platform_for_new

category_platform_for_new %>% 
  dplyr::select(1,2,8,9) -> category_platform_for_new


colnames(category_platform_for_new)[1] <- "Location"
colnames(category_platform_for_new)[2] <- "Item"
colnames(category_platform_for_new)[3] <- "Category"
colnames(category_platform_for_new)[4] <- "Platform"


category_platform_for_new %>% 
  dplyr::mutate(Item = gsub("-", "", Item)) %>% 
  dplyr::select(Item, Category) -> category_for_new
category_for_new[-which(duplicated(category_for_new$Item)),] -> category_for_new

ssmetrics %>% 
  dplyr::mutate(Category = ifelse(Category == is.na(Category), left_join(category_for_new, by = "Item"), Category)) -> ssmetrics


# Platform vlookup from mega data
ssmetrics_mainboard %>% 
  dplyr::select(Item, Platform) -> ssmetrics_mainboard_plat

ssmetrics_mainboard_plat[!duplicated(ssmetrics_mainboard_plat[,c("Item", "Platform")]),] -> ssmetrics_mainboard_plat
ssmetrics_mainboard_plat[-which(duplicated(ssmetrics_mainboard_plat$Item)),] -> ssmetrics_mainboard_plat

merge(ssmetrics, ssmetrics_mainboard_plat[, c("Item", "Platform")], by = "Item", all.x = TRUE) -> ssmetrics  

# here you have N/A for new SKUs. need MS
category_platform_for_new %>% 
  dplyr::mutate(Item = gsub("-", "", Item)) %>% 
  dplyr::select(Item, Platform) -> platform_for_new
platform_for_new[-which(duplicated(platform_for_new$Item)),] -> platform_for_new

ssmetrics %>% 
  dplyr::mutate(Platform = ifelse(Platform == is.na(Platform), left_join(platform_for_new, by = "Item"), Platform)) -> ssmetrics
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
  dplyr::mutate(campus_total_available = campus_total_available_1 + campus_total_available_2) %>% 
  dplyr::select(-campus_total_available_1, -campus_total_available_2) -> ssmetrics_final

# month, FY, year
ssmetrics_final %>% 
  dplyr::mutate(month = lubridate::month(date),
                year = lubridate::year(date),
                FY = paste("FY", lubridate::year(date)+1)) -> ssmetrics_final



# Macro-platform
ssmetrics_final %>% 
  dplyr::left_join(macro_platform, by = "Platform") %>% 
  dplyr::mutate(macro_platform = replace(macro_platform, is.na(macro_platform), Type)) -> ssmetrics_final

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
priority_sku[-which(duplicated(priority_sku$priority_sku)),] -> priority_sku

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
  dplyr::mutate(mfg_line = replace(mfg_line, is.na(mfg_line), Type)) %>% 
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




##### weekly result ##### ----
writexl::write_xlsx(ssmetrics_final, "SS Metrics 0620.xlsx") 




#####################################################################################################################################
###################################################### save & update mega data ######################################################
#####################################################################################################################################
colnames(ssmetrics_mainboard) <- colnames(ssmetrics_final)

ssmetrics_mainboard %>% 
  dplyr::mutate(date = as.integer(date),
                date = as.Date(date, origin = "1899-12-30")) -> ssmetrics_mainboard

readr::type_convert(ssmetrics_mainboard) -> ssmetrics_mainboard

ssmetrics_mainboard %>% 
  dplyr::mutate(max_capacity = replace(max_capacity, is.na(max_capacity), 0)) %>% 
  dplyr::mutate(max_capacity = paste0(round(100*max_capacity, 0), "%")) %>% 
  dplyr::mutate(Safety_Stock = as.double(Safety_Stock)) -> ssmetrics_mainboard



# Check the first line to see the earliest day of the data ----
ssmetrics_mainboard %>% head()

ssmetrics_mainboard %>% 
  dplyr::filter(date != "2021-06-21") %>% 
  dplyr::bind_rows(ssmetrics_final) -> ssmetrics_mainboard







save(ssmetrics_mainboard, file = "ssmetrics_mainboard_7_12_22.rds")


writexl::write_xlsx(ssmetrics_mainboard, "SS Metrics_mainboard_7_12_22.xlsx") 




# Still need to do line (784 ~ 814) for a new items. 









# colnames(ssmetrics_mainboard)[1]	<-	"Month"
# colnames(ssmetrics_mainboard)[2]	<-	"FY"
# colnames(ssmetrics_mainboard)[3]	<-	"Year"
# colnames(ssmetrics_mainboard)[4]	<-	"Category"
# colnames(ssmetrics_mainboard)[5]	<-	"Platform"
# colnames(ssmetrics_mainboard)[6]	<-	"Macro-Platform"
# colnames(ssmetrics_mainboard)[7]	<-	"Loc Name"
# colnames(ssmetrics_mainboard)[8]	<-	"Campus"
# colnames(ssmetrics_mainboard)[9]	<-	"Date"
# colnames(ssmetrics_mainboard)[10]	<-	"Location"
# colnames(ssmetrics_mainboard)[11]	<-	"Item"
# colnames(ssmetrics_mainboard)[12]	<-	"Description"
# colnames(ssmetrics_mainboard)[13]	<-	"Type"
# colnames(ssmetrics_mainboard)[14]	<-	"Stocking type description"
# colnames(ssmetrics_mainboard)[15]	<-	"Planner Name"
# colnames(ssmetrics_mainboard)[16]	<-	"MTO/MTS"
# colnames(ssmetrics_mainboard)[17]	<-	"MPF/Line#"
# colnames(ssmetrics_mainboard)[18]	<-	"SafetyStock"
# colnames(ssmetrics_mainboard)[19]	<-	"BalanceUsable"	
# colnames(ssmetrics_mainboard)[20]	<-	"Sum of BalanceHold(exclude hard hold)"
# colnames(ssmetrics_mainboard)[21]	<-	"Ref"
# colnames(ssmetrics_mainboard)[22]	<-	"Campus Ref"
# colnames(ssmetrics_mainboard)[23]	<-	"Label"
# colnames(ssmetrics_mainboard)[24]	<-	"mfg-line"
# colnames(ssmetrics_mainboard)[25]	<-	"Capacity"
# colnames(ssmetrics_mainboard)[26]	<-	"Capacity Status"
# colnames(ssmetrics_mainboard)[27]	<-	"Current SS Alert"
# colnames(ssmetrics_mainboard)[28]	<-	"Total Cust Order"
# colnames(ssmetrics_mainboard)[29]	<-	"Cust Order qty in the next 5 days"
# colnames(ssmetrics_mainboard)[30]	<-	"WO in next 5 days"
# colnames(ssmetrics_mainboard)[31]	<-	"Receipt in next 5 days"
# colnames(ssmetrics_mainboard)[32]	<-	"Open PO in next 5 days"
# colnames(ssmetrics_mainboard)[33]	<-	"SS Alert after Cust Order in the next 5 days + WO & Receipt"
# colnames(ssmetrics_mainboard)[34]	<-	"SKU has SS"
# colnames(ssmetrics_mainboard)[35]	<-	"SKU >= SS"
# colnames(ssmetrics_mainboard)[36]	<-	"SKU < SS"
# colnames(ssmetrics_mainboard)[37]	<-	"SKU < SS with supply"
# colnames(ssmetrics_mainboard)[38]	<-	"Priority SKU or Unique RM"
# colnames(ssmetrics_mainboard)[39]	<-	"Oil Allocation SKU"
# colnames(ssmetrics_mainboard)[40]	<-	"Campus SS"
# colnames(ssmetrics_mainboard)[41]	<-	"Campus Total Available"
# colnames(ssmetrics_mainboard)[42]	<-	"Campus SKU has SS"
# colnames(ssmetrics_mainboard)[43]	<-	"Campus SKU >= SS"
# colnames(ssmetrics_mainboard)[44]	<-	"Campus SKU < SS"
