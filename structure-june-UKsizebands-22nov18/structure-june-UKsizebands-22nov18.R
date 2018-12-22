library(tidyverse)
library(tidyxl)
library(unpivotr)

cells <-
  xlsx_cells("./structure-june-UKsizebands-22nov18.xlsx") %>%
  dplyr::filter(!is_blank)

formats <- xlsx_formats("./structure-june-UKsizebands-22nov18.xlsx")
green_cell <- dplyr::filter(cells, address == "A6")
green_rgb <- formats$local$fill$patternFill$fgColor$rgb[green_cell$local_format_id]
green_fill <- formats$local$fill$patternFill$fgColor$rgb == "FF008000"

corners <-
  cells %>%
  dplyr::filter(col == 1, green_fill[local_format_id]) %>%
  select(row, col, farm_type = character)

tidy <-
  cells %>%
  partition(corners) %>%
  select(farm_type, cells) %>%
  mutate(cells = map(cells,
                     ~ .x %>%
                       behead("NNW", "year") %>%
                       behead("N", "metric") %>%
                       behead("N", "unit") %>%
                       behead("WNW", "produce") %>%
                       behead("W", "size_statistic"))) %>%
  unnest() %>%
  select(farm_type, produce, size_statistic, year, metric, unit)
