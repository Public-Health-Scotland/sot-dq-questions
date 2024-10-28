# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# create_excel.R
# Angus Morton
# 2024-09-18
# 
# Create excel template filled in with numbers
# 
# R version 4.1.2 (2021-11-01)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#### Step 0 : Housekeeping ----

library(readr)
library(dplyr)
library(lubridate)
library(tidylog)
library(openxlsx)

source(code/pull_numbers.R)

#### Step 1 : import out numbers ----

#### Step x : define styles ----

s_title <- createStyle(
  fontSize = 16,
  textDecoration = "bold",
)

s_subtitle <- createStyle(
  fgFill = "#B8CCE4",
  border = "bottom",
  borderStyle = "thin"
)

s_table <- createStyle(
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
  halign = "center"
)

s_table_header <- createStyle(
  fgFill = "#B8CCE4",
)

s_priority <- createStyle(
  textDecoration = "bold",
)

s_p_high <- createStyle(
  fgFill = "#808080",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
)

s_p_med <- createStyle(
  fgFill = "#BFBFBF",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
)

s_p_low <- createStyle(
  fgFill = "#F2F2F2",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
)

#### Step x : create xlsx file ----

wb <- createWorkbook()

modifyBaseFont(wb, fontSize = 12,
               fontName = "Arial")

addWorksheet(wb, "phs logo")
showGridLines(wb, "phs logo", showGridLines = FALSE)

insertImage(wb, "phs logo", "phs-logo.png",
            startRow = 2, startCol = 8)

addStyle(wb, "phs logo", s_title,
         rows = 2, cols = 2)
addStyle(wb, "phs logo", s_subtitle,
         rows = 4, cols = 2)
addStyle(wb, "phs logo", s_table_header,
         rows = 6, cols = 2:6, gridExpand = TRUE)
addStyle(wb, "phs logo", s_table,
         rows = 7, cols = 2:6, gridExpand = TRUE)
addStyle(wb, "phs logo", s_p_med,
         rows = 8, cols = 2)

#### Step x : write data to excel ----

saveWorkbook(wb, "output/template.xlsx", overwrite = TRUE)





