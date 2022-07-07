library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(readxlsb)
# ssmetrics_mainboard <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/ssmetrics_main_board.xlsx",
#                          col_names = FALSE)

load("ssmetrics_mainboard.rds")

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
# Stock type
load("stock_type.rds")

# Lot Status
Lot_Status <- read_excel("S:/Supply Chain Projects/Linda Liang/reference files/Lot Status Code.xlsx",
                         col_names = FALSE)

colnames(Lot_Status) <- Lot_Status[1, ]
Lot_Status[-1, ] -> Lot_Status

Lot_Status %>% 
  dplyr::rename(Lot_Status = "Lot status",
                Hold_Status = "Hard/Soft Hold") -> Lot_Status

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
reshape2::dcast(custord, ref ~ in_next_7_days, value.var = "Qty", sum) -> custord_pivot




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

# ################################################################################################################################
# ############################################## up to here. two location resovled -> deactivated ################################
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
merge(ssmetrics, Lot_Status[, c("Lot_Status", "Hold_Status")], by = "Lot_Status", all.x = TRUE) -> ssmetrics

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


# what if still N/A? that's new items

types_for_na <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Automation/raw/.xlsx",
                                 col_names = FALSE)

## waiting for Linda's access grant

# Categories and platforms
## waiting for Linda's response about RM






# for new item -> Linda sent me the dossier 
# Categories and platforms -> same logic

# 40:10 we move to pivot
# add platform category
# add vlookup formula for Type for ssmetrics_na






