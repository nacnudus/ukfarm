library(tidyverse)
library(tidyxl)
library(unpivotr)

tidy_sheet <- function(cells) {
  cattle_col <- dplyr::filter(cells, row == 10L, str_detect(character, "Cattle"))$col
  corners <- tibble(row = 7L, col = c(1L, cattle_col))
  cells %>%
  mutate(col = if_else(row >= 10L &
                         !is.na(character) &
                         !bold[local_format_id] &
                         !is.na(fill[local_format_id]) &
                         fill[local_format_id] == green,
                       col + 4L, # Move the poorly placed unit cells
                       col)) %>%
  partition(corners, strict = FALSE) %>%
  select(cells) %>%
  mutate(cells = map(cells,
                     ~ .x %>%
                       behead("NNW", "holding_type") %>%
                       behead("N", "year") %>%
                       behead("ENE", "unit") %>%
                       behead_if(bold[local_format_id],
                                 fill[local_format_id] == green,
                                 direction = "WNW", name = "metric") %>%
                       behead("W", "subgroup") %>%
                       mutate(year = as.character(year),
                              suppressed = if_else(is.na(character),
                                                   FALSE,
                                                   character == "#")) %>%
                       select(holding_type, metric, subgroup, unit, year,
                              value = numeric, suppressed))) %>%
  unnest(cols = c(cells))
}

formats <- xlsx_formats("./file.xlsx")
bold <- formats$local$font$bold
fill <- formats$local$fill$patternFill$fgColor$rgb
green <- "FF008000"

tidy <-
  xlsx_cells("./file.xlsx") %>%
  dplyr::filter(!(sheet %in% c("Map", "Metadata")),
                !is_blank,
                row >= 7L) %>%
  nest(data = !c(sheet)) %>%
  mutate(data = map(data, tidy_sheet)) %>%
  unnest(cols = c(data)) %>%
  mutate_if(is.character, str_trim)
