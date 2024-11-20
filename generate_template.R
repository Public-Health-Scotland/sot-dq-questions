# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# generate_template.R
# Angus Morton
# 2024-11-20
# 
# Once parameters have been selected by the shiny app this script
# generates the template and saves it
# 
# R version 4.1.2 (2021-11-01)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


#### Step 0 : Housekeeping ----

library(readr)
library(dplyr)
library(lubridate)
library(stringr)
library(purrr)
library(tidylog)
library(openxlsx)


paste_data <- function(q, figs) {
  
  title <- paste0("Table ", q-1, " - Data for question ", q)
  
  q_figs <- figs |> 
    filter(Question == q) |> 
    select(`Patient type` = Patient_Type,
           Specialty,
           Indicator,
           6:10)
  
  prev_rows <- figs |> 
    mutate(row_num = row_number()) |> 
    filter(Question == q) |> 
    select(row_num) |> 
    pull() |> 
    min()
  
  title_row <- 5 + (q-2)*4 + (prev_rows-1)
  data_row <- title_row+2
  
  writeData(wb, "SoT Data", title, startRow = title_row, startCol = 2)
  writeData(wb, "SoT Data", q_figs, startRow = data_row, startCol = 2)
  
  addStyle(wb, "SoT Data", style = s_data_table_title,
           rows = title_row, cols = 2)
  addStyle(wb, "SoT Data", s_data_table_header,
           rows = data_row, cols = 2:(ncol(q_figs)+1),
           gridExpand = TRUE)
  addStyle(wb, "SoT Data", s_table,
           rows = data_row:(data_row+nrow(q_figs)),
           cols = 2:(ncol(q_figs)+1),
           gridExpand = TRUE, stack = TRUE)
  
}


merge_column <- function(var, df, table_start) {
  
  df <- df |> 
    mutate(merge = data.table::rleid(.data[[var]]),
           row_num = row_number()) |> 
    group_by(merge) |> 
    mutate(start_row = min(row_num)+table_start,
           end_row = max(row_num)+table_start) |> 
    ungroup()
  
  print(paste("merge groups ", select(df, merge)))
  
  var_col <- which(colnames(df) == var)-1
  
  start_rows <- df |> 
    select(start_row) |> 
    distinct() |> 
    pull()
  
  end_rows <- df |> 
    select(end_row) |> 
    distinct() |> 
    pull()
  
  merge_groups <- map2(start_rows, end_rows, seq)
  
  print(paste(var, table_start))
  print(merge_groups)
  
  map(merge_groups, mergeCells, wb = wb,
      sheet = "SoT Data", cols = var_col)
  
}

merge_table <- function(question, table_start, df) {
  
  columns <- c("Patient_Type", "Specialty", "Indicator")
  
  df <- df |> 
    filter(Question == question)
  
  map(columns, merge_column, df = df, table_start = table_start)
  
}

#### Step 1 : source numbers ----

params <- read_rds('temp/params.rds')
data <- read_rds("data/data.rds")

figs <- params |> 
  filter(Question != 1) |> 
  inner_join(data, by = c("Patient_Type",
                          "NHS_Board_of_Treatment",
                          "Specialty",
                          "Indicator")) |>
  arrange(Date) |> 
  select(-Priority) |> 
  mutate(Date = format(Date, "%d/%m/%Y"),
         Specialty = str_to_title(Specialty)) |> 
  pivot_wider(names_from = "Date", values_from = "value") |> 
  mutate(across(where(is.numeric), ~ replace_na(.,0)))


#### Step 2 : define styles ----

s_title <- createStyle(
  fontSize = 16,
  textDecoration = "bold",
)

s_subtitle <- createStyle(
  fgFill = "#B8CCE4",
  textDecoration = "bold",
  border = "bottom",
  borderStyle = "thin"
)

s_table <- createStyle(
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
  halign = "center",
  valign = "center"
)

s_table_header <- createStyle(
  fgFill = "#B8CCE4",
  
)

s_question_text <- createStyle(
  wrapText = TRUE,
  halign = "left"
)

s_priority <- createStyle(
  textDecoration = "bold",
)

s_p_high <- createStyle(
  fgFill = "#808080",
  textDecoration = "bold",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
)

s_p_med <- createStyle(
  fgFill = "#BFBFBF",
  textDecoration = "bold",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
)

s_p_low <- createStyle(
  fgFill = "#F2F2F2",
  textDecoration = "bold",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "thin",
)

s_data_table_title <- createStyle(
  fontSize = 16,
  textDecoration = "bold",
)

s_data_table_header <- s_table_header <- createStyle(
  fgFill = "#B8CCE4",
  textDecoration = "bold"
)

#### Step 3 : Create workbook ----

board <- figs |> 
  select(NHS_Board_of_Treatment) |> 
  distinct() |> 
  pull()

qe <- figs |> 
  select(last_col()) |> 
  names() |> 
  dmy() |> 
  format("%d %B %Y")

## create workbook and sheets

wb <- createWorkbook()

modifyBaseFont(wb, fontSize = 12,
               fontName = "Arial")

addWorksheet(wb, "SoT")
addWorksheet(wb, "SoT Data")
showGridLines(wb, "SoT", showGridLines = FALSE)
showGridLines(wb, "SoT Data", showGridLines = FALSE)

## Insert common elements

# PHS logo
insertImage(wb, "SoT", "phs-logo.png",
            startRow = 1, startCol = 7,
            width = 2.29, height = 1)

insertImage(wb, "SoT Data", "phs-logo.png",
            startRow = 1, startCol = 8,
            width = 2.29, height = 1)

# Title
title <- paste0("Stage of Treatment - ",
                board,
                " - Quarter Ending ",
                qe)

writeData(wb, "SoT", title, startRow = 1, startCol = 2)
addStyle(wb, "SoT", s_title, rows = 1, cols = 2)

writeData(wb, "SoT Data", title, startRow = 1, startCol = 2)
addStyle(wb, "SoT Data", s_title, rows = 1, cols = 2)


#### Step 4 : Question sheet ----

setColWidths(wb, "SoT", cols = 1:7,
             widths = c(2.5, 11, 2, 10, 40, 35, 28))

setRowHeights(wb, "SoT", rows = 1:12,
              heights = c(30,31,25,15,15,46,28,16,16,16,16,34))

writeData(wb, "SoT", "Data Quality Questions",
          startRow = 4, startCol = 2)
addStyle(wb, "SoT", s_subtitle, rows = 4, cols = 2:7, gridExpand = TRUE)

description <- paste0("The following questions pinpoint areas", 
                      " where local insight will help enhance",
                      " our understanding of recent trends and",
                      " quality assure the data accordingly.")
writeData(wb, "SoT", description, startRow = 6, startCol = 2)

writeData(wb, "SoT", "Priority:", startRow = 8, startCol = 2)
addStyle(wb, "SoT", s_priority, rows = 8, cols = 2)

writeData(wb, "SoT", "High", startRow = 8, startCol = 4)
addStyle(wb, "SoT", s_p_high, rows = 8, cols = 4)

writeData(wb, "SoT", "Medium", startRow = 9, startCol = 4)
addStyle(wb, "SoT", s_p_med, rows = 9, cols = 4)

writeData(wb, "SoT", "Low", startRow = 10, startCol = 4)
addStyle(wb, "SoT", s_p_low, rows = 10, cols = 4)

## Question table

n_q <- figs |> 
  select(Question) |> 
  max()

q_table <- data.frame(Number = c(1:n_q)) |> 
  mutate(` ` = "",
         Question = "",
         `blank` = "",
         Response = "")

q_table[1,3] <- paste0("Are there any new or ongoing data quality issues",
                       " that you wish to bring to our attention?")

for (i in 0:nrow(q_table)) {
  
  mergeCells(wb, "SoT", cols = 4:5, rows = 12+i)
  mergeCells(wb, "SoT", cols = 6:7, rows = 12+i)
  
}

setRowHeights(wb, "SoT", rows = 13:(12+n_q),
              heights = 34)

writeData(wb, "SoT", q_table, startRow = 12, startCol = 2)
addStyle(wb, "SoT", s_table_header, rows = 12, cols = 2:7,
         gridExpand = TRUE)
addStyle(wb, "SoT", s_table, rows = 12:(12+nrow(q_table)), cols = 2:7,
         gridExpand = TRUE, stack = TRUE)
addStyle(wb, "SoT", s_question_text, rows = 13, cols = 4:5,
         stack = TRUE)

#### Step 5 : Data sheet ----

setColWidths(wb, "SoT Data", cols = 1:9,
             widths = c(2.5, 21, 24, 27, 13, 13, 13, 13, 13))

setRowHeights(wb, "SoT Data", rows = 1:2,
              heights = c(40,31))

writeData(wb, "SoT Data", "Accompanying Data",
          startRow = 3, startCol = 2)
addStyle(wb, "SoT Data", s_subtitle, rows = 3, cols = 2:9, gridExpand = TRUE)


# Paste data

map(2:max(figs$Question), paste_data, figs)


# Cell merging

prev_rows <- figs |> 
  mutate(row_num = row_number()) |> 
  group_by(Question) |> 
  summarise(prev_rows = min(row_num)-1) |> 
  ungroup() |> 
  select(prev_rows) |> 
  pull()

data_rows <- 7 + (2:max(figs$Question)-2)*4 + (prev_rows)

questions <- figs |> select(Question) |> distinct() |> pull()

pmap(list(question = questions, table_start = data_rows),
     merge_table, df = figs)


setColWidths(wb, "SoT Data", cols = 3, widths = "auto")


#### Step 6 : write to excel ----

qe <- figs |> 
  select(last_col()) |> 
  names() |> 
  dmy() |> 
  format("%b %Y")

fname <- paste0("output/","SoT Data Quality ", board ," - ", qe, ".xlsx")

saveWorkbook(wb, fname, overwrite = TRUE)

