# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/danielle_adair")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Clean_Data.xlsx")

# Get rid of special characters

names(df) <- gsub(" ", "_", trimws(names(df)))
names(df) <- gsub("\\s+", "_", trimws(names(df), whitespace = "[\\h\\v\\s]+"))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))
names(df) <- gsub("\\$", "", names(df))
names(df) <- gsub("\\+", "_", names(df))

# Trim all values

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))

colnames(df)

vars <- c("PR1_Brown_1"       ,       "PR2_Brown_1"     ,         "PA1_Brown_1"   ,          
          "PA2_Brown_1"        ,      "PR1_Brown_2"      ,        "PR2_Brown_2"    ,          "PA1_Brown_2" ,            
          "PA2_Brown_2"         ,     "PR1_Blue_1"        ,       "PR2_Blue_1"      ,         "PA1_Blue_1"   ,           
           "PA2_Blue_1"          ,     "PR1_Blue_2"        ,       "PR2_Blue_2"      ,         "PA1_Blue_2"   ,           
           "PA2_Blue_2"           ,    "PR1_Green_1"        ,      "PR2_Green_1"      ,        "PA1_Green_1"   ,          
           "PA2_Green_1"           ,   "PR1_Green_2"         ,     "PR2_Green_2"       ,       "PA1_Green_2"    ,         
           "PA2_Green_2"            ,  "Similarity_Brown_1"   ,    "Similarity_Blue_1"  ,      "Similarity_Green_1" ,     
           "Similarity_Brown_2"      , "Similarity_Blue_2"     ,   "Similarity_Green_2"  ,     "EasyToOrder_1"  ,         
           "EasyToOrder_2")

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    Median = numeric(),
    SEM = numeric(),  # Standard Error of the Mean
    SD = numeric(),   # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    median_val <- median(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      Median = median_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

# Example usage with your dataframe and list of variables
df_descriptive_stats <- calculate_descriptive_stats(df, vars)

str(df)
print(head(df))

library(dplyr)
library(tidyr)
library(rstatix)  # for convenient statistical testing


# Corrected data transformation function
transform_data_long <- function(data) {
  # Define columns that are part of the experimental data (excluding metadata and unrelated columns)
  rank_columns <- grep("PR1|PR2|PA1|PA2", names(data), value = TRUE)
  
  # Transform the data to long format
  long_data <- data %>%
    pivot_longer(
      cols = rank_columns,
      names_to = c("Method", "Unit", "Color", "Week"),
      names_pattern = "(PR|PA)(\\d)_(\\w+)_(\\d)",
      values_to = "Rank"
    ) %>%
    mutate(
      Method = ifelse(grepl("PR", Method), "Printed", "Painted"),
      Color = tolower(Color),
      Week = paste("Week", Week)
    )
  
  return(long_data)
}

# Apply the transformation
df_long <- transform_data_long(df)

# Check the structure of the transformed data
print(head(df_long))

# Aggregate the ranks by sum within each method, color, and week
df_aggregated <- df_long %>%
  group_by(ID, Method, Color, Week) %>%
  summarize(Sum_Rank = sum(Rank, na.rm = TRUE), .groups = 'drop')

# Run Friedman test and handle results correctly
run_friedman_tests <- function(data) {
  results <- data %>%
    group_by(Color, Week) %>%
    friedman_test(Sum_Rank ~ Method | ID) %>%
    ungroup() %>%
    select(Color, Week, statistic, p, method)  # Corrected from p.value to p
  
  return(results)
}

# Apply the function
df_friedman_results <- run_friedman_tests(df_aggregated)

# Print the results to check
print(df_friedman_results)

library(dplyr)
library(moments)

# Function to calculate descriptive statistics for grouped data
calculate_grouped_descriptive_stats <- function(data, group_vars, desc_var) {
  results <- data %>%
    group_by(across(all_of(group_vars))) %>%
    summarise(
      Mean = mean({{ desc_var }}, na.rm = TRUE),
      Median = median({{ desc_var }}, na.rm = TRUE),
      SEM = sd({{ desc_var }}, na.rm = TRUE) / sqrt(n()),
      SD = sd({{ desc_var }}, na.rm = TRUE),
      Skewness = skewness({{ desc_var }}, na.rm = TRUE),
      Kurtosis = kurtosis({{ desc_var }}, na.rm = TRUE),
      .groups = 'drop'
    )
  
  return(results)
}

# Apply the function to your aggregated data
# Assuming df_aggregated contains Sum_Rank, Method, Color, and Week
group_vars <- c("Method", "Color", "Week")
df_groupeddesc_stats_results <- calculate_grouped_descriptive_stats(df_aggregated, group_vars, Sum_Rank)

# Print the results to check
print(df_groupeddesc_stats_results)


library(dplyr)
library(rstatix)  # for Wilcoxon test

# Define column pairs manually
column_pairs <- list(
  PR1_Brown = c("PR1_Brown_1", "PR1_Brown_2"),
  PR2_Brown = c("PR2_Brown_1", "PR2_Brown_2"),
  PA1_Brown = c("PA1_Brown_1", "PA1_Brown_2"),
  PA2_Brown = c("PA2_Brown_1", "PA2_Brown_2"),
  PR1_Blue = c("PR1_Blue_1", "PR1_Blue_2"),
  PR2_Blue = c("PR2_Blue_1", "PR2_Blue_2"),
  PA1_Blue = c("PA1_Blue_1", "PA1_Blue_2"),
  PA2_Blue = c("PA2_Blue_1", "PA2_Blue_2"),
  PR1_Green = c("PR1_Green_1", "PR1_Green_2"),
  PR2_Green = c("PR2_Green_1", "PR2_Green_2"),
  PA1_Green = c("PA1_Green_1", "PA1_Green_2"),
  PA2_Green = c("PA2_Green_1", "PA2_Green_2")
)

# Prepare an empty data frame to store results
df_results_wilcoxon <- data.frame()

# Iterate through each pair and perform the Wilcoxon signed-rank test
for (key in names(column_pairs)) {
  col_week1 <- column_pairs[[key]][1]
  col_week2 <- column_pairs[[key]][2]
  
  # Extract the relevant columns and perform the Wilcoxon test
  test_data <- df %>%
    select(ID, Week_1 = .data[[col_week1]], Week_2 = .data[[col_week2]])
  
  # Check for length equality before testing
  if (length(test_data$Week_1) == length(test_data$Week_2)) {
    # Perform Wilcoxon signed-rank test
    test_result <- wilcox.test(test_data$Week_1, test_data$Week_2, paired = TRUE)
    
    # Append the results
    df_results_wilcoxon <- rbind(df_results_wilcoxon, data.frame(
      Group = key,
      Statistic = test_result$statistic,
      P_Value = test_result$p.value,
      Method = "Wilcoxon signed-rank test"
    ))
  } else {
    # Append a notice if the lengths do not match
    df_results_wilcoxon <- rbind(results_df, data.frame(
      Group = key,
      Statistic = NA,
      P_Value = NA,
      Method = "Mismatch in data lengths"
    ))
  }
}

# Print the results
print(df_results_wilcoxon)


library(dplyr)
library(tidyr)

# Create a dataframe with just the relevant columns
relevant_columns <- c("PR1_Brown_1", "PR1_Brown_2", "PR2_Brown_1", "PR2_Brown_2",
                      "PA1_Brown_1", "PA1_Brown_2", "PA2_Brown_1", "PA2_Brown_2",
                      "PR1_Blue_1", "PR1_Blue_2", "PR2_Blue_1", "PR2_Blue_2",
                      "PA1_Blue_1", "PA1_Blue_2", "PA2_Blue_1", "PA2_Blue_2",
                      "PR1_Green_1", "PR1_Green_2", "PR2_Green_1", "PR2_Green_2",
                      "PA1_Green_1", "PA1_Green_2", "PA2_Green_1", "PA2_Green_2")

df_wilcoxon_stats_df <- df %>%
  select(all_of(relevant_columns)) %>%
  pivot_longer(cols = everything(), names_to = "variable", values_to = "rank") %>%
  separate(variable, into = c("method", "color", "week"), sep = "_") %>%
  group_by(method, color, week) %>%
  summarize(
    mean_rank = mean(rank, na.rm = TRUE),
    median_rank = median(rank, na.rm = TRUE),
    sd_rank = sd(rank, na.rm = TRUE),
    min_rank = min(rank, na.rm = TRUE),
    max_rank = max(rank, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  pivot_wider(names_from = week, values_from = c(mean_rank, median_rank, sd_rank, min_rank, max_rank))

# Print the dataframe with descriptive statistics
print(df_wilcoxon_stats_df)





library(ggplot2)
library(dplyr)

# Function to plot comparison between methods for each week
plot_method_comparison <- function(data) {
  # Split the data by week for separate plotting
  week1_data <- data %>% filter(Week == "Week 1")
  week2_data <- data %>% filter(Week == "Week 2")
  
  # Create plot for Week 1
  plot_week1 <- ggplot(week1_data, aes(x = Color, y = Mean, fill = Method)) +
    geom_bar(stat = "identity", position = position_dodge(), width = 0.7) +
    geom_errorbar(aes(ymin = Mean - SEM, ymax = Mean + SEM), width = 0.25, position = position_dodge(0.7)) +
    labs(title = "Comparison of Mean Ranks between Methods for Week 1",
         x = "Colour",
         y = "Mean Rank",
         fill = "Method") +
    theme_minimal()
  
  # Create plot for Week 2
  plot_week2 <- ggplot(week2_data, aes(x = Color, y = Mean, fill = Method)) +
    geom_bar(stat = "identity", position = position_dodge(), width = 0.7) +
    geom_errorbar(aes(ymin = Mean - SEM, ymax = Mean + SEM), width = 0.25, position = position_dodge(0.7)) +
    labs(title = "Comparison of Mean Ranks between Methods for Week 2",
         x = "Colour",
         y = "Mean Rank",
         fill = "Method") +
    theme_minimal()
  
  return(list(Week1 = plot_week1, Week2 = plot_week2))
}

# Run the function with your descriptive statistics data
method_comparison_plots <- plot_method_comparison(df_groupeddesc_stats_results)

# Print the plots to check
print(method_comparison_plots$Week1)
print(method_comparison_plots$Week2)

# Optionally save the plots to files
ggsave("method_comparison_week1.png", method_comparison_plots$Week1, width = 8, height = 6, dpi = 300)
ggsave("method_comparison_week2.png", method_comparison_plots$Week2, width = 8, height = 6, dpi = 300)



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
  
  "Descriptive Stats" = df_descriptive_stats, 
  "Descriptive Stats - Friedman" = df_groupeddesc_stats_results, 
  "Descriptive Stats - Wilc" = df_wilcoxon_stats_df,
  "Friedman" = df_friedman_results,
  "Wilcoxon" = df_results_wilcoxon
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")



# COmbined tests

library(dplyr)
library(tidyr)
library(rstatix)

# Combine PA1 with PA2 and PR1 with PR2 within each week
df_combined <- df %>%
  mutate(
    PA_Brown_1 = PA1_Brown_1 + PA2_Brown_1,
    PA_Brown_2 = PA1_Brown_2 + PA2_Brown_2,
    PR_Brown_1 = PR1_Brown_1 + PR2_Brown_1,
    PR_Brown_2 = PR1_Brown_2 + PR2_Brown_2,
    
    PA_Blue_1 = PA1_Blue_1 + PA2_Blue_1,
    PA_Blue_2 = PA1_Blue_2 + PA2_Blue_2,
    PR_Blue_1 = PR1_Blue_1 + PR2_Blue_1,
    PR_Blue_2 = PR1_Blue_2 + PR2_Blue_2,
    
    PA_Green_1 = PA1_Green_1 + PA2_Green_1,
    PA_Green_2 = PA1_Green_2 + PA2_Green_2,
    PR_Green_1 = PR1_Green_1 + PR2_Green_1,
    PR_Green_2 = PR1_Green_2 + PR2_Green_2
  ) %>%
  select(ID, starts_with("PA_"), starts_with("PR_"))  # Keep only combined columns

# Transform data to long format for analysis
df_long_combined <- df_combined %>%
  pivot_longer(
    cols = -ID,
    names_to = c("Method", "Color", "Week"),
    names_pattern = "(PA|PR)_(\\w+)_(\\d)",
    values_to = "Rank"
  ) %>%
  mutate(
    Method = ifelse(Method == "PA", "Painted", "Printed"),
    Color = tolower(Color),
    Week = paste("Week", Week)
  )

# Sum ranks for each method across weeks
df_aggregated_combined <- df_long_combined %>%
  group_by(ID, Method, Color) %>%
  summarize(Sum_Rank = sum(Rank, na.rm = TRUE), .groups = 'drop')

# Run Friedman test on the aggregated combined data
friedman_results_combined <- df_aggregated_combined %>%
  group_by(Color) %>%
  friedman_test(Sum_Rank ~ Method | ID) %>%
  ungroup() %>%
  select(Color, statistic, p, method)

# Print Friedman test results
print(friedman_results_combined)

# Calculate descriptive statistics for combined data
df_descriptive_combined <- df_aggregated_combined %>%
  group_by(Method, Color) %>%
  summarise(
    Mean = mean(Sum_Rank, na.rm = TRUE),
    Median = median(Sum_Rank, na.rm = TRUE),
    SEM = sd(Sum_Rank, na.rm = TRUE) / sqrt(n()),
    SD = sd(Sum_Rank, na.rm = TRUE),
    Skewness = skewness(Sum_Rank, na.rm = TRUE),
    Kurtosis = kurtosis(Sum_Rank, na.rm = TRUE),
    .groups = 'drop'
  )

# Print descriptive statistics
print(df_descriptive_combined)

# Export results to Excel
data_list_combined <- list(
  "Friedman Combined" = friedman_results_combined,
  "Descriptive Combined" = df_descriptive_combined
)

save_apa_formatted_excel(data_list_combined, "Combined_Results_APA.xlsx")



library(dplyr)
library(moments)

# Combine PA1 and PA2, PR1 and PR2 for each color and week
df_combined <- df %>%
  mutate(
    PA_Brown_1 = PA1_Brown_1 + PA2_Brown_1,
    PA_Brown_2 = PA1_Brown_2 + PA2_Brown_2,
    PR_Brown_1 = PR1_Brown_1 + PR2_Brown_1,
    PR_Brown_2 = PR1_Brown_2 + PR2_Brown_2,
    
    PA_Blue_1 = PA1_Blue_1 + PA2_Blue_1,
    PA_Blue_2 = PA1_Blue_2 + PA2_Blue_2,
    PR_Blue_1 = PR1_Blue_1 + PR2_Blue_1,
    PR_Blue_2 = PR1_Blue_2 + PR2_Blue_2,
    
    PA_Green_1 = PA1_Green_1 + PA2_Green_1,
    PA_Green_2 = PA1_Green_2 + PA2_Green_2,
    PR_Green_1 = PR1_Green_1 + PR2_Green_1,
    PR_Green_2 = PR1_Green_2 + PR2_Green_2
  ) %>%
  select(starts_with("PA_"), starts_with("PR_"))  # Keep only the combined columns

# Define the new combined variable names for descriptive statistics calculation
combined_vars <- colnames(df_combined)

# Function to calculate descriptive statistics for combined variables
calculate_combined_stats <- function(data, desc_vars) {
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    Median = numeric(),
    SEM = numeric(),
    SD = numeric(),
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    median_val <- median(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      Median = median_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

# Calculate descriptive statistics for the combined data
df_combined_stats <- calculate_combined_stats(df_combined, combined_vars)

# Print the results
print(df_combined_stats)


# Export results to Excel
data_list_combined <- list(
  
  "Descriptive Combined" = df_combined_stats
)

save_apa_formatted_excel(data_list_combined, "Combined_Descriptives.xlsx")