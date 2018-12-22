library(tidyverse)
library(tidyxl)
library(unpivotr)

tidy <-
  xlsx_cells("./structure-june-Englandsizebands-22nov18.xlsx") %>%
  dplyr::filter(!is_blank, row >= 6L) %>%
  behead("NNW", "year") %>%
  behead("N", "unit") %>%
  behead("WNW", "produce") %>%
  behead("W", "size_band") %>%
  dplyr::filter(!is.na(unit), # cell BI18
                produce != "Livestock Type") %>% # duplicate rows of column headers
  mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
  select(row, col, year, unit, produce, size_band, value = numeric, suppressed) %>%
  dplyr::filter(!is.na(unit))
