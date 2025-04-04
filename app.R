# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# app.R
# Angus Morton
# 2024-11-08
# 
# Shiny app for picking the parameters
# 
# R version 4.1.2 (2021-11-01)
# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

library(shiny)
library(readr)
library(dplyr)
library(lubridate)
library(janitor)

#source("functions.R")

prepost <- "Snapshot"

snapshot <- as_datetime(file.info(paste0("/PHI_conf/WaitingTimes/SoT/",
                                         "Projects/R Shiny DQ/Snapshot BOXI/",
                                         "CO Quarterly.xlsx"))$mtime)

data_created <- as_datetime(file.info("temp/data_snapshot.rds")$mtime)


if (snapshot > data_created | is.na(data_created)) {
  
  source("read_data.R", local = TRUE)
  
} else {
  
  data <- read_rds("temp/data_snapshot.rds")
  
}

boards <- data |> 
  select(NHS_Board_of_Treatment) |> 
  distinct() |> 
  pull()

ptypes <- data |> 
  select(Patient_Type) |> 
  distinct() |> 
  pull()

params <- data.frame(matrix(ncol = 3, nrow = 0))
x <- c("Patient_Type", "Specialty", "Indicator")
colnames(params) <- x

all_questions <- data.frame(Question = 1,
                            Priority = "Medium",
                            Patient_Type = "generic question",
                            Specialty = "generic question",
                            Indicator = "generic question")


ui <- fluidPage(
  
  fluidRow(
    radioButtons("prepost", label = "Live or Snapshot",
                 choices = c("Snapshot", "Live"), selected = "Snapshot")
  ),
  
  fluidRow(
    selectInput("board", "Select Board", boards)
  ),
  
  fluidRow(
    
    column(width = 3,
           selectInput("ptype", "Select Patient Type", ptypes)
    ),
    column(width = 3, 
           selectInput("spec", "Select Specialty", choices = NULL)
    ),
    column(width = 3, 
           selectInput("indicator", "Select Indicator", choices = NULL)
    ),
    column(width = 3,
           actionButton("add_row", "Add Row"))
  ),
  
  fluidRow(
    
    column(
      width = 6,
      textOutput("current_qnum"),
      tableOutput("current_question"),
      selectInput("priority", "Select Priority",
                  c("High", "Medium", "Low"),
                  selected = "Medium"),
      column(
        width = 3,
        actionButton("clear_question", "Clear Question")
      ),
      column(
        width = 3,
        actionButton("next_question", "Next Question")
      )
    ),
    
    column(
      width = 6,
      tableOutput("all_questions"),
      actionButton("generate", "Generate Template"),
      textOutput("file_exists"))
  )
  
)

server <- function(input, output, session) {
  
  observeEvent(input$prepost, {
    
    prepost <<- input$prepost
    
    if (input$prepost == "Live") {
      if (as.Date(file.info("temp/data_live.rds")$mtime) != Sys.Date() |
          is.na(as.Date(file.info("temp/data_live.rds")$mtime)))
        source("read_data.R", local = TRUE)
      data <<- read_rds("temp/data_live.rds")
    } else if (input$prepost == "Snapshot") {
      data <<- read_rds("temp/data_snapshot.rds")
    }
    
  })
  
  
  
  valid_choices <- reactive({
    filter(data,
           NHS_Board_of_Treatment == input$board,
           Patient_Type == input$ptype,
    )
  })
  
  observeEvent(valid_choices(), {
    current_choice <- input$spec
    choices <- unique(valid_choices()$Specialty)
    updateSelectInput(inputId = "spec", choices = unique(c("All Specialties",
                                                           sort(choices))))
    if (current_choice %in% choices) {
      updateSelectInput(inputId = "spec", selected = current_choice)
    }
  })
  
  observeEvent(valid_choices(), {
    current_choice <- input$indicator
    choices <- unique(valid_choices()$Indicator)
    updateSelectInput(inputId = "indicator", choices = choices)
    if (current_choice %in% choices) {
      updateSelectInput(inputId = "indicator", selected = current_choice)
    }
  })
  
  RV <- reactiveValues()
  RV$current_qnum <- 2
  #RV$prepost <- "Snapshot"
  RV$current_question <- params
  RV$all_questions <- all_questions
  
  observeEvent(input$add_row, {
    
    RV$current_question[nrow(RV$current_question) + 1,] = c(input$ptype,
                                                            input$spec,
                                                            input$indicator)
    
  })
  
  observeEvent(input$next_question, {
    
    RV$current_question <- RV$current_question |> 
      mutate(Question = RV$current_qnum,
             Priority = input$priority)
    
    RV$all_questions <- rbind(RV$all_questions, RV$current_question)
    
    if (nrow(RV$current_question) != 0) {
      
      RV$current_qnum <- RV$current_qnum+1
      
    } 
    
    RV$current_question <- slice(RV$current_question, 0) |> 
      select(-c(Question, Priority))
    
    updateSelectInput(inputId = "priority", selected = "Medium")
    
  })
  
  observeEvent(input$clear_question, {
    
    RV$current_question <- slice(RV$current_question, 0)
    
  })
  
  
  observeEvent(input$generate, {
    
    qe <- data |> 
      select(Date) |> 
      filter(Date == max(Date)) |> 
      distinct() |> 
      pull() |> 
      format("%b %Y")
    
    fpath <- paste0("output/",
                    "SoT Data Quality ", input$board ," - ", qe, ".xlsx")
    
    if (file.exists(fpath)) {
      
      output$file_exists <- renderText({
        
        paste0("File exists. Delete or move existing file to avoid overwriting")
        
      })
      
    } else {
      
      output$file_exists <- renderText("Generating")
      
      RV$all_questions <- RV$all_questions |> 
        mutate(NHS_Board_of_Treatment = input$board, .after = "Question")
      
      write_rds(RV$all_questions, 'temp/params.rds')
      
      source('generate_template.R', local = TRUE)
      
      output$saved <- renderText("Template Saved")
      stopApp()
      
    }
    
  })
  
  output$current_qnum <- renderText({
    paste0("Current Question : ", RV$current_qnum)
  })
  
  output$current_question <- renderTable(RV$current_question)
  
  output$all_questions <- renderTable(RV$all_questions,
                                      digits = 0)
  
}

shinyApp(ui, server)


