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