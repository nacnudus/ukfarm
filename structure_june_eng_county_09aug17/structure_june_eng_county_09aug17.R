# Sheet 1985 Doesn't even have column headers, just describes what "left column"
# and "right column" represent

library(tidyverse)
library(tidyxl)
library(unpivotr)

cells <-
  xlsx_cells("./file.xlsx") %>%
  dplyr::filter(!is_blank)

formats <- xlsx_formats("./file.xlsx")
bold <- formats$local$font$bold

# 1905--1975 -------------------------------------------------------------------

cells_1905_1975 <-
  cells %>%
  dplyr::filter(sheet >= "1905", sheet <= "1985") %>%
  nest(-sheet) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("N", "county") %>%
                      behead("WNW", "topic") %>%
                      behead("W", "subtopic") %>%
                      mutate(suppressed  = if_else(is.na(character), FALSE, character == "#")) %>%
                      select(county, topic, subtopic, amount = numeric, suppressed))) %>%
  unnest()

# 1985 -------------------------------------------------------------------------

cells_1985 <-
  cells %>%
  dplyr::filter(sheet == "1985") %>%
  behead("NNW", "county") %>%
  behead("WNW", "topic") %>%
  behead("W", "subtopic") %>%
  group_by(county) %>%
  mutate(unit = if_else(col == min(col), "holdings", "hectares/livestock")) %>%
  ungroup() %>%
  mutate(suppressed  = if_else(is.na(character), FALSE, character == "#")) %>%
  select(county, topic, subtopic, unit, amount = numeric, suppressed)

# 2010,2013 --------------------------------------------------------------------

cells_2010_2013 <-
  cells %>%
  dplyr::filter(sheet %in% c("2010", "2013")) %>%
  nest(-sheet) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("NNW", "table") %>%
                      behead("NNW", "header1") %>%
                      behead("NNW", "header2") %>%
                      behead("NNW", "unit") %>%
                      behead("N", "year") %>%
                      behead_if(bold[local_format_id], direction = "WSW", name ="region") %>%
                      behead("W", "la") %>%
                      dplyr::filter(!is.na(region)) %>% # extraneous cells at bottom of 2010
                      mutate(suppressed  = if_else(is.na(character), FALSE, character == "#")) %>%
                      select(table, header1, header2, unit, year, region, la, amount = numeric, suppressed))) %>%
  unnest()

# 2011-2,2014-5 ----------------------------------------------------------

cells_2011_2_2014_5 <-
  cells %>%
  dplyr::filter(sheet %in% c("2011", "2012", "2014", "2015")) %>%
  dplyr::filter(row >= 2) %>%
  nest(-sheet) %>%
  mutate(data = map(data,
                    ~ .x %>%
                      behead("NNW", "table") %>%
                      behead("NNW", "header1") %>%
                      behead("N", "header2") %>%
                      behead("W", "region") %>%
                      mutate(suppressed  = if_else(is.na(character), FALSE, character == "#")) %>%
                      select(table, header1, header2, region, amount = numeric, suppressed))) %>%
  unnest()

# 2016 -------------------------------------------------------------------------

cells_2016 <-
  cells %>%
  dplyr::filter(sheet == "2016",
                row >= 2L) %>%
  behead("NNW", "table") %>%
  behead("NNW", "header1") %>%
  behead("N", "header2") %>%
  behead_if(bold[local_format_id], direction = "WSW", name ="region") %>%
  behead("W", "la") %>%
  mutate(suppressed  = if_else(is.na(character), FALSE, character == "#")) %>%
  select(table, header1, header2, region, la, amount = numeric, suppressed)
