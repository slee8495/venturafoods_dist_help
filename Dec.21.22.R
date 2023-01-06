library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)



file <- readxl::read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Distribution Team Help/12.21.22/Book2.xlsx")


# Pivoting
file %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::group_by(order_number, or_ty, original_or_num, actual_ship, x2nd_item_number, description_1, uom, customer_po,
                  ship_to, sold_to, sold_to_name, branch_plant) %>%
  
  dplyr::summarise(quantity = sum(quantity),
                   quantity_ordered = sum(quantity_ordered),
                   quantity_shipped = sum(quantity_shipped)) %>% 
  
  dplyr::relocate(quantity, .after = description_1) %>%
  dplyr::relocate(quantity_ordered, .after = uom) %>% 
  dplyr::relocate(quantity_shipped, .after = quantity_ordered) %>% 
  
  dplyr::ungroup() %>% 
  dplyr::mutate(actual_ship = as.Date(actual_ship)) %>% 
  data.frame() -> tab_1

# Relocating

tab_1 %>% 
  dplyr::relocate(order_number,	or_ty,	original_or_num,	actual_ship,	x2nd_item_number,	description_1,	quantity,
                  uom, quantity_ordered, quantity_shipped, customer_po, ship_to,	sold_to,	sold_to_name, branch_plant) -> for_tab_2

tab_1 %>% 
  dplyr::relocate(order_number,	or_ty,	original_or_num,	actual_ship,	x2nd_item_number,	description_1,	quantity,
                  uom, quantity_ordered, quantity_shipped, customer_po, ship_to,	sold_to,	sold_to_name, branch_plant) -> tab_1


# Renaming
colnames(tab_1)[1] <- "Order Number"
colnames(tab_1)[2] <- "Or Ty"
colnames(tab_1)[3] <- "Original Or Num"
colnames(tab_1)[4] <- "Actual Ship"
colnames(tab_1)[5] <- "2nd Item Number"
colnames(tab_1)[6] <- "Description 1"
colnames(tab_1)[7] <- "Quantity"
colnames(tab_1)[8] <- "UOM"
colnames(tab_1)[9] <- "Quantity Ordered"
colnames(tab_1)[10] <- "Quantity Shipped"
colnames(tab_1)[11] <- "Customer PO"
colnames(tab_1)[12] <- "Ship To"
colnames(tab_1)[13] <- "Sold To"
colnames(tab_1)[14] <- "Sold To Name"
colnames(tab_1)[15] <- "Branch/Plant"



########## Tab 2 ##########


# cont of order number
reshape2::dcast(for_tab_2, sold_to_name ~ . , value.var = "order_number", length) %>% 
  dplyr::rename(orders = ".") -> table_1
reshape2::dcast(for_tab_2, sold_to_name ~ . , value.var = "quantity", sum) %>% 
  dplyr::rename(cases = ".") -> table_2



reshape2::dcast(for_tab_2, sold_to_name + order_number ~ . , value.var = "quantity", length) %>% 
  dplyr::rename(count = ".") %>% 
  plyr::ddply("sold_to_name", transform, count_2 = sum(count)) %>% 
  dplyr::group_by(sold_to_name, count_2) %>% 
  dplyr::count() %>% 
  dplyr::mutate(lines_per_order_avg = count_2 / n) %>%
  dplyr::ungroup() %>% 
  dplyr::select(sold_to_name, lines_per_order_avg) -> table_3




table_1 %>% 
  dplyr::left_join(table_2) %>% 
  dplyr::left_join(table_3) %>% 
  dplyr::mutate(cs_per_order_avg = cases / orders) %>% 
  dplyr::mutate(cases_per_line = cs_per_order_avg / lines_per_order_avg) %>% 
  dplyr::mutate(cases = round(cases, 0),
                lines_per_order_avg = round(lines_per_order_avg, 1),
                cs_per_order_avg = round(cs_per_order_avg, 1),
                cases_per_line = round(cases_per_line, 2)) %>% 
  dplyr::relocate(sold_to_name, orders, cases, cs_per_order_avg, lines_per_order_avg, cases_per_line) -> tab_2




# export to excel

openxlsx::createWorkbook("example") -> example
openxlsx::addWorksheet(example, "tab_name_1")
openxlsx::addWorksheet(example, "tab_name_2")

openxlsx::writeDataTable(example, "tab_name_1", tab_1)
openxlsx::writeDataTable(example, "tab_name_2", tab_2)


openxlsx::saveWorkbook(example, file = "Sales_Data.xlsx")



