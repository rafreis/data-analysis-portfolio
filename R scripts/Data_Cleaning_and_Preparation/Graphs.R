# Set working directory and load required packages
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/librarylyns")

library(readxl)
library(shiny)
library(ggplot2)
library(dplyr)
library(writexl)

# Read the data
data <- read_excel("Data_ToR.xlsx")

# Replace spaces with underscores and remove special characters for Shiny UI
ui_colnames <- gsub("[[:space:][:punct:]]", "_", colnames(data))
colnames(data) <- ui_colnames

# Define the UI
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



# Define the server function
server <- function(input, output, session) {
  
  library(readxl)
  library(shiny)
  library(ggplot2)
  library(dplyr)
  library(writexl)
  
  output$dynamicPlot <- renderPlot({
    
    comp_var <- input$comparison
    cat_var <- input$category
    
    if (is.numeric(data[[comp_var]]) || is.character(data[[comp_var]])) {
      
      if (is.numeric(data[[comp_var]])) {
        
        # Clean names for display
        cleanCatName <- sub("^CAT_", "", cat_var)
        cleanCompName <- sub("^COMP_", "", comp_var)
        cleanCatName <- gsub("_", " ", cleanCatName)
        cleanCompName <- gsub("_", " ", cleanCompName)
        
        ggplot(data, aes(x = .data[[cat_var]], y = .data[[comp_var]])) +
          geom_boxplot() +
          stat_summary(fun = mean, geom = "point", shape = 23, size = 3, fill = "red") +
          theme_minimal() +
          theme(legend.position = "none") +
          ggtitle(paste("Boxplot of", cleanCompName, "by", cleanCatName)) +
          xlab(cleanCatName) +
          ylab(cleanCompName)
        
        # Create dataframe for boxplot background data
        boxplot_data <- data %>% 
          select(.data[[cat_var]], .data[[comp_var]]) %>% 
          rename(Category = .data[[cat_var]], Comparison = .data[[comp_var]])
        
        # Write to Excel
        write_xlsx(boxplot_data, "boxplot_data.xlsx")
        
      } else {
        
        count_data <- as.data.frame(table(data[[cat_var]], data[[comp_var]]))
        colnames(count_data) <- c("category", "comparison", "n")
        
        # Calculate proportions within each category
        count_data <- count_data %>% 
          group_by(category) %>% 
          mutate(prop = n / sum(n)) %>% 
          ungroup()
        
        # Clean names for display
        cleanCatName <- sub("^CAT_", "", cat_var)
        cleanCompName <- sub("^COMP_", "", comp_var)
        cleanCatName <- gsub("_", " ", cleanCatName)
        cleanCompName <- gsub("_", " ", cleanCompName)
        
        # Create dataframe for barplot background data (count_data already available)
        # Write to Excel
        write_xlsx(count_data, "barplot_data.xlsx")
        
        plot <- ggplot(count_data, aes(x = category, fill = comparison)) +
          geom_bar(aes(y = prop * 100), stat="identity", position="dodge") +
          geom_text(aes(y = (prop * 100), label = sprintf("%.1f%%", prop * 100)), 
                    vjust = -0.5, position = position_dodge(0.9)) +
          theme_minimal() +
          labs(fill = "Response") +  # Update legend title
          ggtitle(paste("Frequency Bar Chart of", cleanCompName, "by", cleanCatName)) +
          xlab(cleanCatName) +
          ylab("Proportion (%)")
        
        
        return(plot)
        
        
      }
    }
  })
}




# Run the app
shinyApp(ui = ui, server = server)


library(rsconnect)

rsconnect::setAccountInfo(name='rafreis', token='26C55640A264230BE3DBCA8EB5E79DD5', secret='zcA+N/puQqeHdaKOcxfNeuFKmwDIMO1yp3lvdhvD')

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/librarylyns/App")

deployApp(appName = "Lynsey_App")


## CREATE DATASET

# All possible 'Category' and 'Comparison' columns
category_columns <- colnames(data)[startsWith(colnames(data), "CAT_")]
comparison_columns <- colnames(data)[startsWith(colnames(data), "COMP_")]

# Initialize an empty list to store the dataframes
all_data_frames <- list()

# Initialize an empty data frame to store the combined data
combined_data_frame <- data.frame()

# Loop through all combinations
for(cat_var in category_columns){
  for(comp_var in comparison_columns){
    
    identifier <- paste(cat_var, comp_var, sep = " & ")
    
    if (is.numeric(data[[comp_var]])) {
      # For Numeric Variables: Boxplot
      boxplot_data <- data %>% 
        select(.data[[cat_var]], .data[[comp_var]]) %>% 
        rename(Category = .data[[cat_var]], Comparison = .data[[comp_var]])
      
      # Add identifier
      boxplot_data$Identifier <- identifier
      
      # Add to combined data frame
      combined_data_frame <- rbind(combined_data_frame, boxplot_data)
      
    } else {
      # For Character Variables: Barplot
      count_data <- as.data.frame(table(data[[cat_var]], data[[comp_var]]))
      colnames(count_data) <- c("Category", "Comparison", "Count")
      
      # Add identifier
      count_data$Identifier <- identifier
      
      # Add to combined data frame
      combined_data_frame <- rbind(combined_data_frame, count_data)
    }
  }
}

# Write the combined data frame to Excel
write.xlsx(combined_data_frame, "Combined_Data.xlsx")