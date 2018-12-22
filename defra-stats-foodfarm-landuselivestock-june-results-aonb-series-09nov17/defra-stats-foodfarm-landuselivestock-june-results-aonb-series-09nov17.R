library(tidyverse)
library(tidyxl)
library(unpivotr)

# 2007 only --------------------------------------------------------------------

tidy_2007 <-
  xlsx_cells("./defra-stats-foodfarm-landuselivestock-june-results-aonb-series-09nov17.xlsx", "2007") %>%
  dplyr::filter(!is_blank, row >= 6) %>%
  behead("NNW", "topic") %>%
  behead("NNW", "subtopic") %>%
  behead("N", "unit") %>%
  behead("W", "aonb_id") %>%
  behead("W", "aonb") %>%
  mutate(suppressed = if_else(is.na(character), FALSE, character == "#")) %>%
  select(row, col, topic, subtopic, unit, aonb_id, aonb, value = numeric, suppressed)

# 2008 onwards -----------------------------------------------------------------

cells <-
  xlsx_cells("./defra-stats-foodfarm-landuselivestock-june-results-aonb-series-09nov17.xlsx", "2008")

tidy_sheet <- function(cells) {
  cells %>%
    dplyr::filter(!is_blank, row >= 6) %>%
    behead("NNW", "topic") %>%
    behead("NNW", "subtopic") %>%
    behead("W", "aonb_id") %>%
    behead("W", "aonb") %>%
    mutate(suppressed = if_else(is.na(character), FALSE, character == "#"),
           aonb_id = as.character(aonb_id)) %>%
    select(topic, subtopic, aonb_id, aonb, value = numeric, suppressed)
}

tidy_2008_2016 <-
  xlsx_cells("./defra-stats-foodfarm-landuselivestock-june-results-aonb-series-09nov17.xlsx") %>%
  dplyr::filter(sheet != "Metadata") %>%
  dplyr::filter(sheet != "2007") %>% # 2007 has a third row of column headers
  nest(-sheet) %>%
  mutate(data = map(data, tidy_sheet)) %>%
  unnest()
