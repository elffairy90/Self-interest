# Load/install packages using pacman
if (!require("pacman")) install.packages("pacman")
pacman::p_load(shiny, tidyverse, lubridate)
rm(list = ls())

# Define file path at the start
file_path <- "spending_wide.csv"

# Step 1: Create initial file if needed
if (!file.exists(file_path)) {
  spending_wide <- tibble(
    date = as.Date("2025-06-15"),
    Income = 0,
    `Food & Drink` = 2.00,
    Rent = NA_real_,
    Recreation = NA_real_,
    Fashion = 0.00
  )
  write_csv(spending_wide, file_path)
  message("Created initial spending_wide.csv")
}

# Step 2: Improved add_spending function
add_spending <- function(date = Sys.Date(), income = 0, food = 0, rent = 0, recreation = 0, fashion = 0) {
  tryCatch({
    # Validate and format date
    clean_date <- as.Date(date)
    if (is.na(clean_date)) stop("Invalid date format")
    
    new_row <- tibble(
      date = clean_date,
      Income = income,
      `Food & Drink` = food,
      Rent = rent,
      Recreation = recreation,
      Fashion = fashion
    )
    
    # Read existing data with explicit column types
    if (file.exists(file_path)) {
      existing <- read_csv(file_path, col_types = cols(
        date = col_date(),
        Income = col_double(),
        `Food & Drink` = col_double(),
        Rent = col_double(),
        Recreation = col_double(),
        Fashion = col_double()
      ))
    } else {
      existing <- tibble(
        date = as.Date(character()),
        Income = numeric(),
        `Food & Drink` = numeric(),
        Rent = numeric(),
        Recreation = numeric(),
        Fashion = numeric()
      )
    }
    
    # Append and save
    updated <- bind_rows(existing, new_row) %>% 
      arrange(desc(date))  # Sort by date
    
    write_csv(updated, file_path)
    message("✅ Added entry for ", format(clean_date, "%Y-%m-%d"))
    return(TRUE)
    
  }, error = function(e) {
    message("❌ Error: ", e$message)
    return(FALSE)
  })
}

# Step 3: Shiny App with fixed date display
ui <- fluidPage(
  titlePanel("Spending Tracker"),
  sidebarLayout(
    sidebarPanel(
      dateRangeInput("dateRange", "Select Date Range:",
                     start = Sys.Date() - 30, end = Sys.Date()),
      selectInput("category", "Select Category:",
                  choices = c("All", "Income", "Food & Drink", "Rent", "Recreation", "Fashion"),
                  selected = "All"),
      actionButton("refresh", "Refresh Data"),
      hr(),
      h4("Add New Entry"),
      dateInput("new_date", "Date:", value = Sys.Date()),
      numericInput("new_income", "Income:", 0),
      numericInput("new_food", "Food & Drink:", 0),
      numericInput("new_rent", "Rent:", 0),
      numericInput("new_rec", "Recreation:", 0),
      numericInput("new_fashion", "Fashion:", 0),
      actionButton("add_entry", "Add Entry")
    ),
    mainPanel(
      tableOutput("filtered_data"),
      plotOutput("spending_plot")
    )
  )
)

server <- function(input, output, session) {
  # Reactive data
  spending_data <- reactive({
    input$refresh
    read_csv(file_path, col_types = cols(date = col_date()))
  })
  
  # Filtered data
  filtered_data <- reactive({
    df <- spending_data() %>%
      filter(date >= input$dateRange[1], date <= input$dateRange[2]) %>%
      arrange(desc(date))
    
    if (input$category != "All") {
      df <- df %>% select(date, !!input$category)
      colnames(df) <- c("Date", "Amount")
      df <- df %>% filter(!is.na(Amount) & Amount != 0)
    }
    df
  })
  
  # Formatted table output
  output$filtered_data <- renderTable({
    df <- filtered_data()
    
    # Format dates
    if ("date" %in% names(df)) {
      df <- df %>% mutate(date = format(date, "%Y-%m-%d"))
    }
    if ("Date" %in% names(df)) {
      df <- df %>% mutate(Date = format(Date, "%Y-%m-%d"))
    }
    df
  })
  
  # Plot output
  output$spending_plot <- renderPlot({
    df <- filtered_data()
    
    if (input$category == "All") {
      df_long <- df %>%
        pivot_longer(-date, names_to = "Category", values_to = "Amount") %>%
        filter(Category != "Income", !is.na(Amount))
      
      ggplot(df_long, aes(x = date, y = Amount, fill = Category)) +
        geom_col() +
        labs(title = "Spending by Category", x = "Date", y = "Amount") +
        theme_minimal()
      
    } else if (input$category == "Income") {
      ggplot(df, aes(x = date, y = Income)) +
        geom_col(fill = "darkgreen") +
        labs(title = "Income", x = "Date", y = "Amount") +
        theme_minimal()
      
    } else {
      ggplot(df, aes(x = date, y = .data[[input$category]])) +
        geom_col(fill = "orange") +
        labs(title = input$category, x = "Date", y = "Amount") +
        theme_minimal()
    }
  })
  
  # Add new entry
  observeEvent(input$add_entry, {
    add_spending(
      date = input$new_date,
      income = input$new_income,
      food = input$new_food,
      rent = input$new_rent,
      recreation = input$new_rec,
      fashion = input$new_fashion
    )
    updateNumericInput(session, "new_income", value = 0)
    updateNumericInput(session, "new_food", value = 0)
    updateNumericInput(session, "new_rent", value = 0)
    updateNumericInput(session, "new_rec", value = 0)
    updateNumericInput(session, "new_fashion", value = 0)
  })
}

shinyApp(ui, server)