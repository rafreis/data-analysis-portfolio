ui <- fluidPage(
  titlePanel("Interactive Data Visualization"),
  sidebarLayout(
    sidebarPanel(
      selectInput("category", "Select Category:", choices = colnames(data)[startsWith(colnames(data), "CAT_")]),
      selectInput("job_category", "Select Job Category:", choices = c("All", unique(data$CAT__What_best_describes_your_job_role__Recoded))),
      selectInput("comparison", "Select Comparison:", choices = colnames(data)[startsWith(colnames(data), "COMP_")])
      
    ),
    mainPanel(
      plotOutput("dynamicPlot"),
      tableOutput("summaryTable")
    )
  )
)