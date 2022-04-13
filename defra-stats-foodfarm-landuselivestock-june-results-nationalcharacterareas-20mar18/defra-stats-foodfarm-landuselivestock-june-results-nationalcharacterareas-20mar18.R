library(tidyverse)
library(tidyxl)
library(unpivotr)

tidy <-
  xlsx_cells("./file.xlsx") %>% # Excel can't handle long filenames
  dplyr::filter(!is_blank, sheet != "Metadata", row >= 2L) %>%
  nest(data = !c(sheet)) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("NNW", "year") %>%
                      behead("NNW", "metric") %>%
                      behead("N", "subgroup") %>%
                      behead("W", "nca") %>%
                      select(nca, metric, subgroup, year))) %>%
  unnest(cols = c(data))
