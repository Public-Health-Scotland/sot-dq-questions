# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# pull_numbers.R
# Angus Morton
# 2024-10-22
# 
# Pull out the numbers for the post snapshot questions
# 
# R version 4.1.2 (2021-11-01)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#### Step 0 : Housekeeping ----

library(readr)
library(dplyr)
library(lubridate)
library(stringr)
library(openxlsx)

#### Step 0a : read in parameters from excel ----

params <- read.xlsx("question_parameters.xlsx")

#### Step 1 : import dashboard data ----

dashboard_fpath <- paste0("/PHI_conf/WaitingTimes/SoT/Projects/R Shiny DQ/",
                          "Snapshot BOXI/")

perf_nop <- read.xlsx(paste0(dashboard_fpath, "CO Quarterly.xlsx"),
                  sheet = "NewOP",
                          sep.names = "_") |> 
  mutate(Date = dmy(Date)) |> 
  filter(Date >= max(Date)-years(1)-days(1),
         month(Date) %in% c(3,6,9,12)) |> 
  rename(Indicator = `Ongoing/Completed`) |> 
  mutate(Indicator = if_else(Indicator == "Completed",
                             "Attendances", "Ongoing Waits"))

perf_ipdc <- read.xlsx(paste0(dashboard_fpath, "CO Quarterly.xlsx"),
                      sheet = "IPDC",
                      sep.names = "_") |> 
  mutate(Date = dmy(Date)) |> 
  filter(Date >= max(Date)-years(1)-days(1),
         month(Date) %in% c(3,6,9,12)) |> 
  rename(Indicator = `Ongoing/Completed`) |> 
  mutate(Indicator = if_else(Indicator == "Completed",
                             "Admissions", "Ongoing Waits"))

perf <- bind_rows(perf_nop, perf_ipdc) |> 
  rename(value = `Number_Seen/On_list`) |> 
  select(1:6)

rm(perf_nop, perf_ipdc)

rr_nop <- read.xlsx(paste0(dashboard_fpath, "RR Monthly.xlsx"),
                sheet = "ALL",
                sep.names = "_") |> 
  mutate(
    Date = rollforward(my(Date)),
    quarter = lubridate::quarter(Date, type = "date_last")) |> 
  select(-Date) |> 
  rename(Date = quarter) |> 
  filter(Date >= max(Date)-years(1)-days(1),
         month(Date) %in% c(3,6,9,12)) |> 
  group_by(Patient_Type, NHS_Board_of_Treatment, Specialty, Date) |> 
  summarise(
    across(where(is.numeric), sum)
  ) |> 
  ungroup() |> 
  pivot_longer(Additions_to_list:Other_reasons, names_to = "Indicator")

rr_ipdc <- read.xlsx(paste0(dashboard_fpath, "RR Monthly.xlsx"),
                sheet = "IPDC",
                sep.names = "_") |> 
  mutate(
    Date = rollforward(my(Date)),
    quarter = lubridate::quarter(Date, type = "date_last")) |> 
  select(-Date) |> 
  rename(Date = quarter) |> 
  filter(Date >= max(Date)-years(1)-days(1),
         month(Date) %in% c(3,6,9,12)) |> 
  group_by(Patient_Type, NHS_Board_of_Treatment, Specialty, Date) |> 
  summarise(
    across(where(is.numeric), sum)
  ) |> 
  ungroup() |> 
  pivot_longer(Additions_to_list:Other_reasons, names_to = "Indicator")

rr <- bind_rows(rr_nop, rr_ipdc)

rm(rr_nop, rr_ipdc)

data <- bind_rows(perf, rr) |> 
  mutate(Indicator = str_replace_all(Indicator, "_", " "),
         NHS_Board_of_Treatment = if_else(
           NHS_Board_of_Treatment == "Golden Jubilee National Hospital",
           "NHS Golden Jubilee",
           NHS_Board_of_Treatment
         ))

rm(perf, rr)

#### Step 2 : pull numbers ----

figs <- inner_join(params, data, by = c("Patient_Type",
                                        "NHS_Board_of_Treatment",
                                        "Specialty",
                                        "Indicator")) |>
  arrange(Date) |> 
  mutate(Date = format(Date, "%d/%m/%Y"),
         Specialty = str_to_title(Specialty)) |> 
  pivot_wider(names_from = "Date", values_from = "value") |> 
  mutate(across(where(is.numeric), ~ replace_na(.,0)))
