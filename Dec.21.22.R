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


tab_1






