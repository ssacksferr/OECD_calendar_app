library(shiny)
library(tidyverse)
library(readxl)
library(writexl)

# Define UI
ui <- fluidPage(
  titlePanel("Generate Calendar (.vcs) File"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Upload Calendar from Template (.xlsx format)"),
      downloadButton("download_vcs", "Generate .vcs File"),
      downloadButton("download_file1", "Download Template (.xlsx)"),
      tags$div(
        id = "How to use",
        style = "border: 1px solid black; padding: 10px;",
        strong("How to use"),
        p("This app will read a .xslx file and convert it to an importable calendar for outlook"),
        p("Your file should follow the format of the template you can download in this app."),
        p("Only entries that are marked as", span('in_calendar = No', style = "color:blue"), "will be imported.")
      )
    ),
    mainPanel(
      tableOutput("data_table")
    )
  )
)

# Define server logic
server <- function(input, output) {
  
  # Read uploaded file
  data <- reactive({
    req(input$file)
    read_xlsx(input$file$datapath)
  })
  
  # Display uploaded data
  output$data_table <- renderTable({
    data()
  })
  
  # Generate .vcs file
  output$download_vcs <- downloadHandler(
    filename = function() {
      "calendar.vcs"
    },
    content = function(file) {
      unique_entries <- data() %>% distinct(country, date, calendar_description, .keep_all = T) %>% 
        filter(in_calendar == "No")
      
      vcs_content <- ""
      
      for (i in seq_len(nrow(unique_entries))) {
        entry <- unique_entries[i, ]
        event_content <- paste(
          "BEGIN:VCALENDAR",
          "VERSION:1.0",
          "BEGIN:VEVENT",
          paste("SUMMARY:", entry$calendar_title, sep = ""),
          paste("DESCRIPTION:", entry$calendar_description, sep = ""),
          paste("DTSTART:", entry$date, "T080000Z", sep = ""),
          paste("DTEND:", entry$date, "T083000Z", sep = ""),
          "END:VEVENT",
          "END:VCALENDAR",
          sep = "\n"
        )
        vcs_content <- paste(vcs_content, event_content, sep = "\n")
      }
      
      # Write calendar to .vcs file
      writeLines(vcs_content, file)
    }
  )
  
  # Generate File 1 (.xlsx)
  output$download_file1 <- downloadHandler(
    filename = function() {
      "Template.xlsx"
    },
    content = function(file) {
      # Generate the data for File 1
      file1_data <- read_xlsx("calendar_test.xlsx")
      
      # Write data to .xlsx file
      write_xlsx(file1_data, file)
    }
  )
  
}

# Run the Shiny app
shinyApp(ui = ui, server = server)