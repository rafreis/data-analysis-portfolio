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
      selectInput("job_category", "Select Job Category:", choices = c("All", unique(data$CAT__What_best_describes_your_job_role__Recoded))),
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
    job_category_selected <- input$job_category
    
    # Filter data based on selected Job Category
    # Filter data based on selected Job Category
    if(job_category_selected != "All") {
      filtered_data <- data %>% filter(CAT__What_best_describes_your_job_role__Recoded == job_category_selected)
    } else {
      filtered_data <- data
    }
    
    
    if (is.numeric(filtered_data[[comp_var]]) || is.character(filtered_data[[comp_var]])) {
      
      cleanCatName <- sub("^CAT_", "", cat_var)
      cleanCompName <- sub("^COMP_", "", comp_var)
      cleanCatName <- gsub("_", " ", cleanCatName)
      cleanCompName <- gsub("_", " ", cleanCompName)
      
      temp_data <- filtered_data %>% select(all_of(cat_var), all_of(comp_var))
      renamed_data <- temp_data %>% rename(Category = all_of(cat_var), Comparison = all_of(comp_var))
      
      if (is.numeric(filtered_data[[comp_var]])) {
        
        boxplot <- ggplot(renamed_data, aes(x = Category, y = Comparison)) + 
          geom_boxplot() +
          stat_summary(fun = mean, geom = "point", shape = 23, size = 3, fill = "red") +
          theme_minimal() +
          theme(legend.position = "none") +
          ggtitle(paste("Boxplot of", cleanCompName, "by", cleanCatName)) +
          xlab(cleanCatName) +
          ylab(cleanCompName)
        
        return(boxplot)
        
      } else {
        
        count_filtered_data <- as.data.frame(table(filtered_data[[cat_var]], filtered_data[[comp_var]]))
        colnames(count_filtered_data) <- c("category", "comparison", "n")
        
        # Calculate proportions within each category
        count_filtered_data <- count_filtered_data %>% 
          group_by(category) %>% 
          mutate(prop = n / sum(n)) %>% 
          ungroup()
        
        # Clean names for display
        cleanCatName <- sub("^CAT_", "", cat_var)
        cleanCompName <- sub("^COMP_", "", comp_var)
        cleanCatName <- gsub("_", " ", cleanCatName)
        cleanCompName <- gsub("_", " ", cleanCompName)
        
        print("Before select")
        print(head(filtered_data))
        
        temp_data <- filtered_data %>% select(all_of(cat_var), all_of(comp_var))
        
        print("After select")
        print(head(temp_data))
        
        boxplot_filtered_data <- temp_data %>% rename(Category = all_of(cat_var), Comparison = all_of(comp_var))
        
        print("After rename")
        print(head(boxplot_filtered_data))
        
        plot <- ggplot(count_filtered_data, aes(x = category, fill = comparison)) +
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

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/librarylyns/App_v2")

deployApp(appName = "Lynsey_Appv2")


## CREATE DATASET

# Identify the numeric 'COMP_' columns
numeric_comp_cols <- names(data)[sapply(data, is.numeric) & grepl("^COMP_", names(data))]

# Identify the 'CAT_' columns
cat_cols <- grep("^CAT_", names(data), value = TRUE)

# Initialize an empty data frame to store the results
final_df <- data.frame()

# Loop through each 'CAT_' column and calculate means for numeric 'COMP_' columns
for (cat in cat_cols) {
  
  temp_df <- data %>%
    select(all_of(cat), all_of(numeric_comp_cols)) %>%
    group_by(across(all_of(cat))) %>%
    summarise(across(all_of(numeric_comp_cols), mean, na.rm = TRUE), .groups = 'drop') 
  
  # Add the CAT variable and category as columns
  temp_df$Variable <- cat
  temp_df <- temp_df %>% rename(Category = !!cat)
  
  final_df <- bind_rows(final_df, temp_df)
}

# Reorder columns to have 'Variable' and 'Category' first
final_df <- final_df %>% select(Variable, Category, everything())


# Identify the character 'COMP_' columns
char_comp_cols <- names(data)[sapply(data, is.character) & grepl("^COMP_", names(data))]

# Initialize an empty data frame to store the frequency percentages
final_freq_df <- data.frame()

# Loop through each 'CAT_' column and calculate frequency percentages for character 'COMP_' columns
for (cat in cat_cols) {
  
  temp_df <- data %>%
    select(all_of(cat), all_of(char_comp_cols)) %>%
    gather(key = "comp_var", value = "comp_value", -all_of(cat)) %>%
    group_by(across(all_of(cat)), comp_var, comp_value) %>%
    summarise(count = n(), .groups = 'drop') %>%
    group_by(across(all_of(cat)), comp_var) %>%
    mutate(freq_percent = (count / sum(count)) * 100, .groups = 'drop')
  
  # Add the CAT variable and category as columns
  temp_df$Variable <- cat
  temp_df <- temp_df %>% rename(Category = !!cat)
  
  final_freq_df <- bind_rows(final_freq_df, temp_df)
}

# Reorder columns to have 'Variable' and 'Category' first
final_freq_df <- final_freq_df %>% select(Variable, Category, comp_var, comp_value, freq_percent)


# Export both dataframes into a single Excel file with different sheets
write_xlsx(list("Mean_Values" = final_df, "Frequency_Percentages" = final_freq_df), "Summary_Results.xlsx")
