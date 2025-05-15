setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/patriarchz/Second Order")

library(readxl)
df <- read_xlsx('ANE 7600 Pharmacology I Spaced-Repetition Study Program Pre-Implementation  (Responses) (1).xlsx', sheet = 'Pre-Implementation')


# Data Recoding

library(dplyr)

recode_all_columns <- function(df, recoding_scheme, columns = NULL) {
  if (is.null(columns)) {
    columns <- names(df)
  }
  
  df %>%
    mutate(across(all_of(columns), function(x) {
      sapply(x, function(value) {
        if (as.character(value) %in% names(recoding_scheme)) {
          return(recoding_scheme[[as.character(value)]])
        } else {
          return(value)
        }
      })
    }))
}


# Recoding scheme for agreement scale
numeric_scheme <- list(
  'Strongly Disagree' = 1,
  'Disagree' = 2,
  'Neutral' = 3,
  'Agree' = 4,
  'Strongly Agree' = 5
)

# Recoding the specified columns in the sample dataframe
recoded_df <- recode_all_columns(df, numeric_scheme)



# Comparative Bar Plots

library(ggplot2)
library(dplyr)
library(tidyr)
library(stringr)

generate_comparative_bar_plot <- function(df, vars, factor_column, factor_levels = NULL, font_size = 12, output_filename = "comparative_bar_plot.png") {
  # Convert the specified columns to numeric
  df[vars] <- lapply(df[vars], function(x) as.numeric(as.character(x)))
  
  # Reshape data to long format, excluding the factor column
  long_df <- df %>%
    pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "value") %>%
    drop_na(value)
  
  # Set the factor levels if provided, otherwise factorize as is
  if (!is.null(factor_levels)) {
    long_df[[factor_column]] <- factor(long_df[[factor_column]], levels = factor_levels)
  } else {
    long_df[[factor_column]] <- factor(long_df[[factor_column]])
  }
  
  # Generate the plot
  p <- ggplot(long_df, aes(x = variable, y = value, fill = !!sym(factor_column))) +
    geom_bar(stat = "summary", fun.y = "mean", position = position_dodge()) +
    geom_text(aes(label = sprintf("%.3f", ..y..)), stat = "summary", fun.y = "mean", position = position_dodge(width = 0.9), vjust = -0.3, size = 3.5) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 20)) + # Wrap labels for long names
    scale_fill_brewer(palette = "Pastel1", name = factor_column) + # Use a color palette for the factor levels
    theme_minimal(base_size = font_size) +
    theme(axis.text.x = element_text(angle = 0, hjust = 0.5, size = font_size), # Centered X-axis text
          axis.text.y = element_text(size = font_size)) +
    labs(y = "Mean", x = "") +
    ggtitle("Comparative Bar Plot") +
    theme(plot.background = element_rect(fill = "white"),
          legend.position = "bottom")
  
  # Print the plot to display
  print(p)
  
  # Save the plot
  ggsave(filename = output_filename, plot = p, width = 10, height = 6, bg = "white")
}


vars1 <- c("I have trouble sleeping when there is an upcoming exam",
          "I usually feel well rested before an exam",
          "I study class material every day",
          "I find myself cramming before exams") 

factor_col <- "TimePoint"

factor_levels <- c("Pre-Implementation", "Post-Implementation")

# Call the function
generate_comparative_bar_plot(recoded_df, vars1, factor_col, factor_levels, font_size = 10.5, output_filename = "comparative_bar_plot1.png")

vars2 <- c("I feel comfortable before taking an exam",
           "I feel anxious before taking an exam",
           "I feel prepared before taking an exam",
           "I feel like I could have prepared better for my exams") 

generate_comparative_bar_plot(recoded_df, vars2, factor_col, factor_levels, font_size = 10.5, output_filename = "comparative_bar_plot2.png")

# Wilcoxon Signed-Rank Test with rank biserial correlation effect size

run_paired_wilcox_tests <- function(data, factor_col, subject_col, measurement_cols) {
  # Ensure factor_col is a factor and get its levels
  data[[factor_col]] <- as.factor(data[[factor_col]])
  factor_levels <- levels(data[[factor_col]])
  
  # Check if there are at least two levels
  if (length(factor_levels) < 2) {
    stop("Error: factor column must have at least two levels")
  }
  
  # Filter out rows where subject is NA
  data <- data %>% filter(!is.na(!!as.name(subject_col)))
  
  # Initialize a dataframe to store results
  results_df <- data.frame(Measurement = character(), Mean_Group1 = numeric(), Mean_Group2 = numeric(),
                           W_Statistic = numeric(), P_Value = numeric(), Effect_Size = numeric())
  
  # Iterate over each measurement column
  for (measurement_col in measurement_cols) {
    # Filter data for the two conditions, handling NA in measurement columns
    data1 <- data %>% filter(!!as.name(factor_col) == factor_levels[1], !is.na(!!as.name(measurement_col)))
    data2 <- data %>% filter(!!as.name(factor_col) == factor_levels[2], !is.na(!!as.name(measurement_col)))
    
    # Merge datasets by subject
    paired_data <- merge(data1, data2, by = subject_col)
    
    # Calculate means
    mean1 <- mean(paired_data[[paste0(measurement_col, ".x")]], na.rm = TRUE)
    mean2 <- mean(paired_data[[paste0(measurement_col, ".y")]], na.rm = TRUE)
    
    # Run the Wilcoxon signed-rank test
    wilcox_result <- wilcox.test(paired_data[[paste0(measurement_col, ".x")]], paired_data[[paste0(measurement_col, ".y")]], paired = TRUE, exact = FALSE)
    
    # Compute effect size (rank biserial correlation)
    n_pairs <- nrow(paired_data)
    effect_size <- wilcox_result$statistic / (n_pairs * (n_pairs + 1) / 2)
    
    # Naming columns with factor level names
    mean_col_name1 <- paste("Mean", factor_levels[1], sep = "_")
    mean_col_name2 <- paste("Mean", factor_levels[2], sep = "_")
    
    # Create a new data frame for this measurement
    new_df <- data.frame(Measurement = measurement_col, Mean_Group1 = mean1, Mean_Group2 = mean2,
                         W_Statistic = wilcox_result$statistic, P_Value = wilcox_result$p.value, Effect_Size = effect_size)
    
    # Dynamically rename the mean columns
    names(new_df)[names(new_df) == "Mean_Group1"] <- mean_col_name1
    names(new_df)[names(new_df) == "Mean_Group2"] <- mean_col_name2
    
    # Append results to the results dataframe
    results_df <- rbind(results_df, new_df)
  }
  
  # Return the results dataframe
  return(results_df)
}

vars <- c("I have trouble sleeping when there is an upcoming exam",
          "I usually feel well rested before an exam",
          "I study class material every day",
          "I find myself cramming before exams",
          "I feel comfortable before taking an exam",
          "I feel anxious before taking an exam",
          "I feel prepared before taking an exam",
          "I feel like I could have prepared better for my exams")
  
results <- run_paired_wilcox_tests(recoded_df, "TimePoint", "ID", vars)
print(results)

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
  "Numeric Dataframe" = recoded_df,
  "Wilcoxon Test Results" = results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")