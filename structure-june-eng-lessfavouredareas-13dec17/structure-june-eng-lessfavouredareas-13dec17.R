library(tidyverse)
library(tidyxl)
library(unpivotr)

cells <-
  xlsx_cells("./structure-june-eng-lessfavouredareas-13dec17.xlsx") %>%
  dplyr::filter(!is_blank, row >= 2L)

formats <- xlsx_formats("./structure-june-eng-lessfavouredareas-13dec17.xlsx")
font_size <- formats$local$font$size

livestock <-
  cells %>%
  dplyr::filter(sheet %in% c("Livestock (animals)", "Livestock (holdings)")) %>%
  nest(-sheet) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("NNW", "header1") %>%
                      behead("NNW", "header2") %>%
                      behead("N", "header3") %>%
                      behead_if(font_size[local_format_id] > 11, direction = "WNW", name = "year") %>%
                      behead("W", "areas") %>%
                      mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
                      select(header1, header2, header3, year, areas, value = numeric, suppressed))) %>%
  unnest()

other <-
  cells %>%
  dplyr::filter(!(sheet %in% c("Introduction",
                               "Livestock (animals)",
                               "Livestock (holdings)",
                               "Metadata "))) %>%
  nest(-sheet) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("NNW", "header1") %>%
                      behead("N", "header2") %>%
                      behead_if(font_size[local_format_id] > 11, direction = "WNW", name = "year") %>%
                      behead("W", "areas") %>%
                      mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
                      select(header1, header2, year, areas, value = numeric, suppressed)))

other$data[[6]] %>%
  dplyr::filter(is.na(value))
