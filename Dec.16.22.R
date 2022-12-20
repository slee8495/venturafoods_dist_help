library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)
library(janitor)
library(lubridate)



file <- readxl::read_excel("622-Torlake Sales Data  for R.xlsx")


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
  data.frame() -> file_2

# Relocating
file_2 %>% 
  dplyr::relocate(order_number,	or_ty,	original_or_num,	actual_ship,	x2nd_item_number,	description_1,	quantity,
                  uom, quantity_ordered, quantity_shipped, customer_po, ship_to,	sold_to,	sold_to_name, branch_plant) -> file_2


# Renaming
colnames(file_2)[1] <- "Order Number"
colnames(file_2)[2] <- "Or Ty"
colnames(file_2)[3] <- "Original Or Num"
colnames(file_2)[4] <- "Actual Ship"
colnames(file_2)[5] <- "2nd Item Number"
colnames(file_2)[6] <- "Description 1"
colnames(file_2)[7] <- "Quantity"
colnames(file_2)[8] <- "UOM"
colnames(file_2)[9] <- "Quantity Ordered"
colnames(file_2)[10] <- "Quantity Shipped"
colnames(file_2)[11] <- "Customer PO"
colnames(file_2)[12] <- "Ship To"
colnames(file_2)[13] <- "Sold To"
colnames(file_2)[14] <- "Sold To Name"
colnames(file_2)[15] <- "Branch/Plant"
  

writexl::write_xlsx(file_2, "Sales Data.xlsx")
