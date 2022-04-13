library(tidyverse)
library(tidyxl)
library(unpivotr)

formats <- xlsx_formats("./structure-june-eng-farmtypeseries-13dec18.xlsx")
bold <- formats$local$font$bold
font_size <- formats$local$font$size

tidy <-
  xlsx_cells("./structure-june-eng-farmtypeseries-13dec18.xlsx") %>%
  dplyr::filter(sheet != "Metadata ", !is_blank, row >= 2L) %>%
  nest(data = !c(sheet)) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("NNW", "header1") %>%
                      behead("N", "header2") %>%
                      behead_if(bold[local_format_id],
                                font_size[local_format_id] > 11,
                                direction = "WNW", name = "year") %>%
                      behead("W", "farm_type") %>%
                      mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
                      select(header1, header2, year, farm_type, value = numeric, suppressed)))

tidy
