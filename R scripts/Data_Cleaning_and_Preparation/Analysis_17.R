setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/getsellgo/Yougov")

library(readxl)
df <- read_xlsx('CleanData.xlsx')


library(tidyverse)

long_df <- df %>%
  pivot_longer(cols = Total:Midlands, names_to = "Group", values_to = "Percentage") %>%
  arrange(Year, Question)

long_df$Year <- factor(long_df$Year)

# Categorize Groups

long_df <- long_df %>%
  mutate(Category = case_when(
    Group %in% c("Male", "Female") ~ "Gender",
    Group %in% c("Con", "Lab", "Lib Dem") ~ "Vote in 2019 GE",
    Group %in% c("Remain", "Leave") ~ "EU Ref 2016",
    Group %in% c("18-24", "25-49", "50-64", "65+") ~ "Age",
    Group %in% c("ABC1", "C2DE") ~ "Social Grade",
    Group %in% c("North", "Midlands / Wales", "Rest of South", "Scotland", "London") ~ "Region",
    TRUE ~ "Other"
  ))



# Line Plots

plot_group_trends <- function(data, question, answer, groups) {
  # Filter the data and remove rows with NA in Percentage
  filtered_data <- data %>%
    filter(Question == question, Answer == answer, Group %in% groups, !is.na(Percentage))
  
  # Create the line graph with data labels
  ggplot(filtered_data, aes(x = Year, y = Percentage, color = Group, group = Group)) +
    geom_line() +
    geom_text(aes(label = Percentage), vjust = -0.5, size = 3) +
    labs(title = paste("Trends for:", question, "-", answer),
         x = "Year",
         y = "Percentage",
         color = "Group") +
    theme_minimal()
}



# List of unique categories
categories <- unique(long_df$Category)

# Loop through each category and plot
for (cat in categories) {
  # Filter the data for the current category and 'Total'
  cat_groups <- unique(long_df$Group[long_df$Category == cat])
  plot_groups <- c("Total", cat_groups)
  
  # Plotting function for the current category
  plot <- plot_group_trends(data = long_df, 
                    question = "Which, if any, of the following types of oil have you heard of? \r\nPlease tick all that apply.", 
                    answer = "Palm oil", 
                    groups = plot_groups)
  print(plot)
  #saving the plot
  ggsave(paste0("plot_", cat, ".png"), width = 10, height = 6)
}

# Loop through each category and plot
for (cat in categories) {
  # Filter the data for the current category and 'Total'
  cat_groups <- unique(long_df$Group[long_df$Category == cat])
  plot_groups <- c("Total", cat_groups)
  
  # Plotting function for the current category
  plot <- plot_group_trends(data = long_df, 
                            question = "And from what you know, do you think the following is generally produced in a way that is friendly or unfriendly to the environment?\r\nPalm Oil is generally produced in a way that is...", 
                            answer = "Environmentally unfriendly", 
                            groups = plot_groups)
  print(plot)
  #saving the plot
  ggsave(paste0("plot_", cat, ".png"), width = 10, height = 6)
}

# Loop through each category and plot
for (cat in categories) {
  # Filter the data for the current category and 'Total'
  cat_groups <- unique(long_df$Group[long_df$Category == cat])
  plot_groups <- c("Total", cat_groups)
  
  # Plotting function for the current category
  plot <- plot_group_trends(data = long_df, 
                            question = "And from what you know, do you think the following is generally produced in a way that is friendly or unfriendly to the environment?\r\nPalm Oil is generally produced in a way that is...", 
                            answer = "Environmentally friendly", 
                            groups = plot_groups)
  print(plot)
  #saving the plot
  ggsave(paste0("plot_", cat, ".png"), width = 10, height = 6)
}

unique(long_df$Question)
unique(long_df$Answer)

# Integrated CrossTab

# Assuming long_df is your original data frame and it has a 'Category' column
new_df <- long_df %>%
  select(Year, Question, Answer, Group, Category, Percentage) %>%
  pivot_wider(names_from = Year, values_from = Percentage) %>%
  mutate(Change_2023_2021 = (`2023` - `2021`) / `2021`,
         Change_2023_2016 = (`2023` - `2016`) / `2016`) %>%
  arrange(Category, Question, Answer, Group)



## Export Results

library(openxlsx)

save_apa_formatted_excel <- function(data_list, filename) {
  wb <- createWorkbook()  # Create a new workbook
  
  for (i in seq_along(data_list)) {
    # Define the sheet name
    sheet_name <- names(data_list)[i]
    if (is.null(sheet_name)) sheet_name <- paste("Sheet", i)
    addWorksheet(wb, sheet_name)  # Add a sheet to the workbook
    
    # Convert matrices to data frames, if necessary
    data_to_write <- if (is.matrix(data_list[[i]])) as.data.frame(data_list[[i]]) else data_list[[i]]
    
    # Include row names as a separate column, if they exist
    if (!is.null(row.names(data_to_write))) {
      data_to_write <- cbind("Index" = row.names(data_to_write), data_to_write)
    }
    
    # Write the data to the sheet
    writeData(wb, sheet_name, data_to_write)
    
    # Define styles
    header_style <- createStyle(textDecoration = "bold", border = "TopBottom", borderColour = "black", borderStyle = "thin")
    bottom_border_style <- createStyle(border = "bottom", borderColour = "black", borderStyle = "thin")
    
    # Apply styles
    addStyle(wb, sheet_name, style = header_style, rows = 1, cols = 1:ncol(data_to_write), gridExpand = TRUE)
    addStyle(wb, sheet_name, style = bottom_border_style, rows = nrow(data_to_write) + 1, cols = 1:ncol(data_to_write), gridExpand = TRUE)
    
    # Set column widths to auto
    setColWidths(wb, sheet_name, cols = 1:ncol(data_to_write), widths = "auto")
  }
  
  # Save the workbook
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Example usage
data_list <- list(
  "DataFrame" = new_df)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

# Input values
sample_proportion <- 0.41  # Replace with your sample proportion
confidence_level <- 0.95  # Replace with your desired confidence level (e.g., 95%)

# Calculate standard error
standard_error <- sqrt((sample_proportion * (1 - sample_proportion)) / 1720)  # n is the sample size

# Calculate critical value from the z-table
z_score <- qnorm((1 + confidence_level) / 2)

# Calculate margin of error
margin_of_error <- z_score * standard_error

# Print the margin of error
cat("Margin of Error:", margin_of_error, "\n")




# Load ggplot2 package
library(ggplot2)

# Create a data frame with your data
data <- data.frame(
  Year = c(2016, 2021, 2022, 2023),
  Percentage = c(41, 71, 72, 68),
  Error = c(2, 2, 2, 2)  # Margin of error
)

# Plot the line graph with error bars
ggplot(data, aes(x = Year, y = Percentage)) + 
  geom_line(color = "red") +  # Red line connecting the points
  geom_point() +  # Points for each year
  geom_errorbar(aes(ymin = Percentage - Error, ymax = Percentage + Error), width = 0.1) +  # Error bars
  geom_text(aes(label = Percentage), nudge_y = 4, color = "black") +  # Adjusted data labels
  scale_x_continuous(breaks = data$Year) +  # Ensure all years are shown on x-axis
  ggtitle("Perception of Palm Oil Production Being Environmentally Unfriendly") +
  xlab("Year") +
  ylab("Percentage (%)") +
  theme_minimal() +
  theme(
    panel.background = element_rect(fill = "white", colour = "white"),  # Set background to white
    panel.grid.major.x = element_blank(),  # Remove vertical major grid lines
    panel.grid.minor.x = element_blank(),  # Remove vertical minor grid lines
    plot.background = element_rect(fill = "white", colour = "white")  # Set plot background to white
  )