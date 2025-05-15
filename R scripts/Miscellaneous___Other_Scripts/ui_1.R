ui <- fluidPage(
  titlePanel("Interactive Data Visualization"),
  sidebarLayout(
    sidebarPanel(
      selectInput("category", "Select Category:", choices = colnames(data)[startsWith(colnames(data), "CAT_")]),
      selectInput("comparison", "Select Comparison:", choices = colnames(data)[startsWith(colnames(data), "COMP_")])
    ),
    mainPanel(
      plotOutput("dynamicPlot"),
      tableOutput("summaryTable")
    )
  )
)