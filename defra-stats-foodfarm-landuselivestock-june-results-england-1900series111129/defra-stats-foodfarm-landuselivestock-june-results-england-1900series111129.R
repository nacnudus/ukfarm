library(tidyverse)
library(tidyxl)
library(unpivotr)

cells <-
  xlsx_cells("file.xlsx") %>% # Excel can't read files with long names
  dplyr::filter(!is_blank, row >= 9)

formats <- xlsx_formats("file.xlsx")

bold <- formats$local$font$bold
size <- formats$local$font$size

corners <-
  cells %>%
  dplyr::filter(!is.na(character),
                bold[local_format_id],
                size[local_format_id] == 11) %>%
  select(row, col, partition = character)

partitions <-
  cells %>%
  partition(corners, nest = FALSE) %>%
  select(-corner_row, -corner_col)

years <-
  cells %>%
  dplyr::filter(row == 9) %>%
  select(row, col, year = numeric)

crops <-
  partitions %>%
  dplyr::filter(partition == "Crops and fallow (thousand hectares)") %>%
  behead("W", "metric") %>%
  enhead(years, "N") %>%
  dplyr::filter(!is.na(metric)) %>% # cell containing "see note 3"
  select(metric, year, value = numeric)

livestock <-
  partitions %>%
  dplyr::filter(partition == "Livestock (numbers in thousands)") %>%
  behead_if(bold[local_format_id], direction = "WNW", name = "stock") %>%
  behead("W", "metric") %>%
  enhead(years, "N") %>%
  select(stock, metric, year, value = numeric)



