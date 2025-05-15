setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/hugorondan")

library(readxl)
df <- read_xlsx('Colonies per site Pre and Post treatment..xlsx', sheet = 'CleanData')

# Get rid of special characters

names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("___", "_", names(df))
names(df) <- gsub("/", "_", names(df))

# Wilcoxon Signed-Rank Test
library(dplyr)

# Revised Function for Paired Wilcoxon Tests with Additional Statistics
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
  results_df <- data.frame(Measurement = character(), Mean_Group1 = numeric(), Mean_Group2 = numeric(), W_Statistic = numeric(), P_Value = numeric())
  
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
    
    # Naming columns with factor level names
    mean_col_name1 <- paste("Mean", factor_levels[1], sep = "_")
    mean_col_name2 <- paste("Mean", factor_levels[2], sep = "_")
    
    # Create a new data frame for this measurement
    new_df <- data.frame(Measurement = measurement_col, Mean_Group1 = mean1, Mean_Group2 = mean2, W_Statistic = wilcox_result$statistic, P_Value = wilcox_result$p.value)
    
    # Dynamically rename the mean columns
    names(new_df)[names(new_df) == "Mean_Group1"] <- mean_col_name1
    names(new_df)[names(new_df) == "Mean_Group2"] <- mean_col_name2
    
    # Append results to the results dataframe
    results_df <- rbind(results_df, new_df)
  }
  
  # Return the results dataframe
  return(results_df)
}


results <- run_paired_wilcox_tests(df, "Product", "Horse", c("wall_1_10_dilution", "wall_1_1000_dilution","wall_1_100000","sole_1_1000_dilution","sole_1_10000_dilution","frog_1_1000_dilution","frog_1_10000_dilution"))
print(results)


## DESCRIPTIVE STATS BY FACTORS

calculate_means_and_sds_by_factors <- function(data, variables, factors) {
  # Create an empty list to store intermediate results
  results_list <- list()
  
  # Iterate over each variable
  for (var in variables) {
    # Create a temporary data frame to store results for this variable
    temp_results <- data.frame(Variable = var)
    
    # Iterate over each factor
    for (factor in factors) {
      # Aggregate data by factor
      agg_data <- aggregate(data[[var]], by = list(data[[factor]]), 
                            FUN = function(x) c(Mean = mean(x, na.rm = TRUE), SD = sd(x, na.rm = TRUE)))
      
      # Create columns for each level of the factor
      for (level in unique(data[[factor]])) {
        level_agg_data <- agg_data[agg_data[, 1] == level, ]
        
        if (nrow(level_agg_data) > 0) {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- level_agg_data$x[1, "Mean"]
          temp_results[[paste0(factor, "_", level, "_SD")]] <- level_agg_data$x[1, "SD"]
        } else {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- NA
          temp_results[[paste0(factor, "_", level, "_SD")]] <- NA
        }
      }
    }
    
    # Add the results for this variable to the list
    results_list[[var]] <- temp_results
  }
  
  # Combine all the results into a single dataframe
  descriptive_stats_bygroup <- do.call(rbind, results_list)
  return(descriptive_stats_bygroup)
}

vars <- c("wall_1_10_dilution", "wall_1_1000_dilution","wall_1_100000","sole_1_1000_dilution","sole_1_10000_dilution","frog_1_1000_dilution","frog_1_10000_dilution")
factors <- c("TimePoint", "Product")
calculate_means_and_sds_by_factors(df, vars, factors)
descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df, vars, factors)

# REDUCTION STATISTICS

library(dplyr)

# Assuming df is your data frame

# Calculate mean PreWash measurements
prewash_means <- df %>% 
  filter(TimePoint == "PreWash") %>% 
  summarise(across(starts_with("wall_"), mean, na.rm = TRUE),
            across(starts_with("sole_"), mean, na.rm = TRUE),
            across(starts_with("frog_"), mean, na.rm = TRUE))

# Calculate mean PostWash measurements for each product
postwash_means_cleantrax <- df %>% 
  filter(TimePoint == "PostWash", Product == "CleanTrax") %>% 
  summarise(across(starts_with("wall_"), mean, na.rm = TRUE),
            across(starts_with("sole_"), mean, na.rm = TRUE),
            across(starts_with("frog_"), mean, na.rm = TRUE))

postwash_means_iodine <- df %>% 
  filter(TimePoint == "PostWash", Product == "2%_iodine_Tincture") %>% 
  summarise(across(starts_with("wall_"), mean, na.rm = TRUE),
            across(starts_with("sole_"), mean, na.rm = TRUE),
            across(starts_with("frog_"), mean, na.rm = TRUE))

# Calculate percentage reductions
percentage_reduction_cleantrax <- (prewash_means - postwash_means_cleantrax) / prewash_means * 100
percentage_reduction_iodine <- (prewash_means - postwash_means_iodine) / prewash_means * 100

# Create a new data frame
reduction_df <- data.frame(
  Measurement = names(prewash_means),
  Percentage_Reduction_CleanTrax = as.numeric(percentage_reduction_cleantrax[1,]),
  Percentage_Reduction_2_Iodine = as.numeric(percentage_reduction_iodine[1,])
)

# View the result
print(reduction_df)


# Compare levels of contamination

# Assuming df is your data frame
df_pre_wash <- df %>% filter(TimePoint == "PreWash")

library(dplyr)
library(broom)

# Function to perform pairwise comparisons for multiple independent samples
compare_pairwise_samples <- function(data, columns) {
  # Create an empty data frame to store results
  results <- data.frame(
    Comparison = character(),
    Mean1 = numeric(),
    Mean2 = numeric(),
    W_Statistic = numeric(),
    P_Value = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Loop over all combinations of columns for pairwise comparison
  for (i in 1:(length(columns) - 1)) {
    for (j in (i + 1):length(columns)) {
      col1 <- columns[i]
      col2 <- columns[j]
      
      # Calculate means
      mean1 <- mean(data[[col1]], na.rm = TRUE)
      mean2 <- mean(data[[col2]], na.rm = TRUE)
      
      # Perform the Wilcoxon test
      test_result <- wilcox.test(data[[col1]], data[[col2]], paired = FALSE)
      
      # Tidy the test result and add to the results data frame
      tidy_result <- tidy(test_result)
      results <- rbind(results, c(paste(col1, "vs", col2), mean1, mean2, tidy_result$statistic, tidy_result$p.value))
    }
  }
  
  # Naming the columns of the result data frame
  names(results) <- c("Comparison", "Mean1", "Mean2", "W_Statistic", "P_Value")
  
  return(results)
}

columns_to_compare <- c("wall_1_100000", "sole_1_10000_dilution", "frog_1_10000_dilution")
result_df_precontamination <- compare_pairwise_samples(df_pre_wash, columns_to_compare)

# Print the result
print(result_df_precontamination)

# Power Analysis

library(pwr)

effect_size <- 0.8  
n <- 10             
sig_level <- 0.05  

# Calculate power for a two-sample t-test as an approximation
power_result <- pwr.t.test(d = effect_size, n = n, sig.level = sig_level, type = "paired", alternative = "two.sided")

print(power_result)


# Assuming the same effect size and significance level
d <- 0.5          # Effect size (Cohen's d)
sig_level <- 0.05 # Significance level
power <- 0.8      # Desired power

# Calculate required sample size for a one-tailed test
sample_size_result <- pwr.t.test(d = d, sig.level = sig_level, power = power, type = "paired", alternative = "two.sided")

print(sample_size_result)

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
  "Wilcoxon Results" = results,
  "Descriptive Stats" = descriptive_stats_bygroup,
  "% Reductions" = reduction_df,
  "Precontamination Comparison" = result_df_precontamination
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")