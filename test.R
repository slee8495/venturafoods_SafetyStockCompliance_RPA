test <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Book1.xlsx")

# do the intersect, there are two columns; micro and inv_bal
setdiff(test$micro, test$inv_bal)

# how do you find something exists in micro, but not in inv_bal?
test %>% 
  filter(!micro %in% inv_bal)

test %>% 
  filter(!inv_bal %in% micro) -> test_result

write_xlsx(test_result, "C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Safety Stock Compliance/Book1_result.xlsx")
