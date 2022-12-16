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

file %>% 
  data.frame() %>% 
  janitor::clean_names() %>% 
  dplyr::group_by(order_number, or_ty, original_or_num, last_status, next_status, request_date, original_promised, original_promised_time,
                  actual_ship, x2nd_item_number, description_1, uom, location, lot_serial_number, invoice_number,
                  customer_po, order_co, hd_cd, ship_to, sold_to, sold_to_name, secondary_quantity, secondary_uo_m, extended_amount,
                  requested_time, original_or_type, orig_or_co, x3rd_item_number, shipment_number, pick_number, delivery_number,
                  unit_price, price_uom, foreign_unit_price, foreign_extended_amount, order_date, short_item_no, document_number,
                  doc_ty, doc_co, scheduled_pick, scheduled_pick_time, actual_ship_time, invoice_date, cancel_date, g_l_date, promised_delivery,
                  promised_delivery_time, branch_plant, ln_ty, rel_ord, rel_ord_type, related_po_wo_no, related_po_wo_line_no,
                  transaction_originator, carrier_number, f_h, mod_trn, zone_no, deliver_to, pull_signal) %>% 
  dplyr::summarise(quantity = sum(quantity),
                   quantity_ordered = sum(quantity_ordered),
                   quantity_shipped = sum(quantity_shipped),
                   quantity_canceled = sum(quantity_canceled),
                   quantity_backordered = sum(quantity_backordered)) %>% 
  dplyr::relocate(quantity, .after = description_1) %>%
  dplyr::relocate(quantity_ordered, .after = uom) %>% 
  dplyr::relocate(quantity_shipped, .after = quantity_ordered) %>% 
  dplyr::relocate(quantity_canceled, .after = quantity_shipped) %>% 
  dplyr::relocate(quantity_backordered, .after = quantity_canceled) %>% 
  dplyr::ungroup() %>% 
  data.frame() -> file_2


writexl::write_xlsx(file_2, "Sales Data.xlsx")
