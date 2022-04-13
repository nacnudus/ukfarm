library(tidyverse)
library(tidyxl)
library(unpivotr)

cells <-
  xlsx_cells("./structure-june-eng-localauthority-09jan18.xlsx") %>%
  dplyr::filter(!is_blank, row >= 2L)

formats <- xlsx_formats("./structure-june-eng-localauthority-09jan18.xlsx")
bold <- formats$local$font$bold

land_livestock <-
  cells %>%
  dplyr::filter(str_detect(sheet, "Land")) %>%
  nest(data = !c(sheet)) %>%
  mutate(sheet = str_trim(sheet),
         data = map(data,
                    ~ .x %>%
                      # Going west first beheads the "Local Authority" column
                      # header before it interferes with column 2 data in NNW.
                      behead_if(bold[local_format_id], direction = "WSW", name = "region") %>%
                      behead("W", "local_authority") %>%
                      behead("NNW", "header1") %>%
                      behead("NNW", "header2") %>%
                      behead("N", "year"))) %>%
  unnest(cols = c(data)) %>%
  mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
  select(sheet, row, col, header1, header2, year, region, local_authority, value = numeric, suppressed)

labour <-
  cells %>%
  dplyr::filter(str_detect(sheet, "Labour")) %>%
  nest(data = !c(sheet)) %>%
  mutate(sheet = str_trim(sheet),
         data = map(data,
                    ~ .x %>%
                      behead("NNW", "header1") %>%
                      behead("NNW", "header2") %>%
                      behead("NNW", "header3") %>%
                      behead("N", "year") %>%
                      behead_if(bold[local_format_id], direction = "WSW", name = "region") %>%
                      behead("W", "local_authority"))) %>%
  unnest(cols = c(data)) %>%
  mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
  select(header1, header2, header3, year, region, local_authority, value = numeric, suppressed)
