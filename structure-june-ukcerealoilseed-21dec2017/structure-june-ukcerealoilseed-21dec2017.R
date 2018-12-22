library(tidyverse)
library(tidyxl)
library(unpivotr)

tidy_sheet <- function(cells) {
  corners <-
    cells %>%
    dplyr::filter(col == 1,
                  centred[local_format_id]) %>%
    mutate(row = row - 1L) %>%
    select(row, col, table = character)
  cells %>%
    partition(corners, strict = FALSE) %>%
    select(table, cells) %>%
    mutate(cells = map(cells,
                       ~ .x %>%
                         behead("NNE", "unit") %>%
                         behead("N", "year") %>%
                         behead("WNW", "nation") %>%
                         behead("WNW", "country") %>%
                         behead("WNW", "region") %>%
                         mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
                         select(unit, year, nation, country, region, value = numeric, suppressed))) %>%
    unnest()
}

formats <- xlsx_formats("./structure-june-ukcerealoilseed-21dec2017.xlsx")
bold <- formats$local$font$bold
centred <- formats$local$alignment$horizontal == "center"

cells <- # Sheet "UK cereal yields summary" is trivial.
  xlsx_cells("./structure-june-ukcerealoilseed-21dec2017.xlsx") %>%
  dplyr::filter(str_detect(sheet, "^Regional"),
                !is_blank,
                row >= 2L) %>%
  nest(-sheet) %>%
  mutate(data = map(data, tidy_sheet)) %>%
  unnest()
