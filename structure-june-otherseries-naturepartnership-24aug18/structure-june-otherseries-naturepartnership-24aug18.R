library(tidyverse)
library(tidyxl)
library(unpivotr)

tidy <-
  xlsx_cells("./structure-june-otherseries-naturepartnership-24aug18.xlsx") %>%
  dplyr::filter(sheet != "Metadata",
                !is_blank,
                row >= 6L) %>%
  nest(-sheet) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("NNW", "header1") %>%
                      behead("N", "header2") %>%
                      behead("W", "local_nature_partnership") %>%
                      mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
                      select(header1, header2, local_nature_partnership, value = numeric))) %>%
  unnest()

