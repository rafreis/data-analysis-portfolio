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
