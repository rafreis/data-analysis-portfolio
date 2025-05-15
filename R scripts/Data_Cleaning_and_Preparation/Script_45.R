# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/opalskies")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df_montreal <- read.xlsx("Haitian Informants Montreal-Final  (Responses) (1).xlsx")
df_tijuana <- read.xlsx("Haitian Informants Tijuana-Final  (Responses) (1).xlsx")


# Identify column names in both dataframes
cols_montreal <- colnames(df_montreal)
cols_tijuana <- colnames(df_tijuana)

# Find common and different columns
common_columns <- intersect(cols_montreal, cols_tijuana)
different_columns_montreal <- setdiff(cols_montreal, cols_tijuana)
different_columns_tijuana <- setdiff(cols_tijuana, cols_montreal)

# Print results
cat("### Common Columns ###\n")
print(common_columns)

cat("\n### Columns Unique to Montreal ###\n")
print(different_columns_montreal)

cat("\n### Columns Unique to Tijuana ###\n")
print(different_columns_tijuana)

# Add a 'Country' column to each dataframe
df_montreal$City <- "Montreal"
df_tijuana$City <- "Tijuana"

# Ensure both dataframes have the same column names structure
common_cols <- union(colnames(df_montreal), colnames(df_tijuana))

# Add any missing columns in each dataframe with NA
for (col in common_cols) {
  if (!(col %in% colnames(df_montreal))) {
    df_montreal[[col]] <- NA
  }
  if (!(col %in% colnames(df_tijuana))) {
    df_tijuana[[col]] <- NA
  }
}

# Combine the two dataframes using rbind
df <- rbind(df_montreal, df_tijuana)

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

str(df)

# Rename the column
colnames(df)[colnames(df) == "5D..Education.level"] <- "Education.level"

# Standardize responses and apply them directly to the dataframe
df$Education.level <- sapply(df$Education.level, function(x) {
  x <- tolower(x)  # Convert to lowercase
  if (grepl("no school|none", x)) {
    return("No School")
  } else if (grepl("primary|up to 2nd grade", x)) {
    return("Primary")
  } else if (grepl("secondary", x)) {
    return("Secondary")
  } else if (grepl("high school|finished hs", x)) {
    return("High School")
  } else if (grepl("bachelor|degree in", x)) {
    return("Bachelor Degree")
  } else if (grepl("graduate degree|phd|master|doctor", x)) {
    return("Graduate Degree")
  } else if (grepl("college", x)) {
    return("Some College")
  } else {
    return("Other")
  }
})

# Convert back to a clean factor for ordered analysis
df$Education.level <- factor(df$Education.level, 
                             levels = c("No School", "Primary", "Secondary", 
                                        "High School", "Some College", 
                                        "Bachelor Degree", "Graduate Degree", "Other"))

# View the updated distribution
print(table(df$Education.level))


# Merge categories with fewer than 5 observations
df$Education.level <- as.character(df$Education.level)  # Convert to character

df$Education.level[df$Education.level %in% c("No School", "Some College", "Other")] <- "Other"
df$Education.level[df$Education.level %in% c("Bachelor Degree", "Graduate Degree")] <- "Higher Education"

# Convert back to factor with updated levels
df$Education.level <- factor(df$Education.level, 
                             levels = c("Primary", "Secondary", "High School", 
                                        "Higher Education", "Other"))

# View the updated distribution
print(table(df$Education.level))


# Racism column
df$Feel.Racism <- ifelse(!is.na(df$`16MXa..On.a.scale.of.1._.5_.with.1.being.most_.how.much.racism.do.you.feel.in.Mexico`), 
                    df$`16MXa..On.a.scale.of.1._.5_.with.1.being.most_.how.much.racism.do.you.feel.in.Mexico`, 
                    df$`16CAa..On.a.scale.of.1._.5_.with.1.being.most_.how.much.racism.do.you.feel.in.Canada`)

# Comfort/Acceptance column
df$Comfort.Acceptance <- ifelse(!is.na(df$`16MXb..On.a.scale.of.1._.5_.with.1.being.least_.how.comfortable_accepted.do.you.feel.in.Mexico`), 
                                df$`16MXb..On.a.scale.of.1._.5_.with.1.being.least_.how.comfortable_accepted.do.you.feel.in.Mexico`, 
                                df$`16CAb..On.a.scale.of.1._.5_.with.1.being.least_.how.comfortable_accepted.do.you.feel.in.Canada`)

# Gender Equality column
df$Gender.Equality <- ifelse(!is.na(df$`16MXc..On.a.scale.of.1._.5_.which.1.being.lowest_.how.much.would.you.rank.gender.equality.in.Mexico`), 
                             df$`16MXc..On.a.scale.of.1._.5_.which.1.being.lowest_.how.much.would.you.rank.gender.equality.in.Mexico`, 
                             df$`16CAc..On.a.scale.of.1._.5_.which.1.being.lowest_.how.much.would.you.rank.gender.equality.in.Canada`)


vars_likert <- c("Feel.Racism", 
                 "Comfort.Acceptance", "Gender.Equality")


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
df_descriptive_stats <- calculate_descriptive_stats(df, vars_likert)


calculate_means_sds_medians_by_factors <- function(data, variables, factors) {
  # Create an empty list to store intermediate results
  results_list <- list()
  
  # Iterate over each variable
  for (var in variables) {
    # Create a temporary data frame to store results for this variable
    temp_results <- data.frame(Variable = var)
    
    # Iterate over each factor
    for (factor in factors) {
      # Aggregate data by factor: Mean, SD, Median
      agg_data <- aggregate(data[[var]], 
                            by = list(data[[factor]]), 
                            FUN = function(x) c(Mean = mean(x, na.rm = TRUE),
                                                SD = sd(x, na.rm = TRUE),
                                                Median = median(x, na.rm = TRUE)))
      
      # Create columns for each level of the factor
      for (level in unique(data[[factor]])) {
        level_agg_data <- agg_data[agg_data[, 1] == level, ]
        
        if (nrow(level_agg_data) > 0) {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- level_agg_data$x[1, "Mean"]
          temp_results[[paste0(factor, "_", level, "_SD")]] <- level_agg_data$x[1, "SD"]
          temp_results[[paste0(factor, "_", level, "_Median")]] <- level_agg_data$x[1, "Median"]
        } else {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- NA
          temp_results[[paste0(factor, "_", level, "_SD")]] <- NA
          temp_results[[paste0(factor, "_", level, "_Median")]] <- NA
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


df_descriptive_stats_bygroup <- calculate_means_sds_medians_by_factors(df, vars_likert, "City")


# Mann-Whitney Test

library(dplyr)

# Function to calculate effect size for Mann-Whitney U tests
# Note: There's no direct equivalent to Cohen's d for Mann-Whitney, but we can use rank biserial correlation as an effect size measure.

run_mann_whitney_tests <- function(data, group_col, group1, group2, measurement_cols) {
  results <- data.frame(Variable = character(),
                        U_Value = numeric(),
                        P_Value = numeric(),
                        Effect_Size = numeric(),
                        Group1_Size = integer(),
                        Group2_Size = integer(),
                        stringsAsFactors = FALSE)
  
  # Iterate over each measurement column
  for (var in measurement_cols) {
    # Extract group data with proper filtering and subsetting
    group1_data <- data %>%
      filter(.[[group_col]] == group1) %>%
      select(var) %>%
      pull() %>%
      na.omit()
    
    group2_data <- data %>%
      filter(.[[group_col]] == group2) %>%
      select(var) %>%
      pull() %>%
      na.omit()
    
    group1_size <- length(group1_data)
    group2_size <- length(group2_data)
    
    if (group1_size > 0 && group2_size > 0) {
      # Perform Mann-Whitney U test
      mw_test_result <- wilcox.test(group1_data, group2_data, paired = FALSE, exact = FALSE)
      
      # Calculate effect size (rank biserial correlation) for Mann-Whitney U test
      effect_size <- (mw_test_result$statistic - (group1_size * group2_size / 2)) / sqrt(group1_size * group2_size * (group1_size + group2_size + 1) / 12)
      
      results <- rbind(results, data.frame(
        Variable = var,
        U_Value = mw_test_result$statistic,
        P_Value = mw_test_result$p.value,
        Effect_Size = effect_size,
        Group1_Size = group1_size,
        Group2_Size = group2_size
      ))
    }
  }
  
  return(results)
}

df_mannwhitney <- run_mann_whitney_tests(df, "City", "Montreal", "Tijuana", vars_likert)



# Percentages relative to the column
chi_square_analysis_multiple <- function(data, row_vars, col_var) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over row variables
  for (row_var in row_vars) {
    # Ensure that the variables are factors
    data[[row_var]] <- factor(data[[row_var]])
    data[[col_var]] <- factor(data[[col_var]])
    
    # Create a crosstab with percentages
    crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 2) * 100
    
    # Check if the crosstab is 2x2
    is_2x2_table <- all(dim(table(data[[row_var]], data[[col_var]])) == 2)
    
    # Perform chi-square test with correction for 2x2 tables
    chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]), correct = is_2x2_table)
    
    # Convert crosstab to a dataframe
    crosstab_df <- as.data.frame.matrix(crosstab)
    
    # Create a dataframe for this pair of variables
    for (level in levels(data[[row_var]])) {
      level_df <- data.frame(
        "Row_Variable" = row_var,
        "Row_Level" = level,
        "Column_Variable" = col_var,
        check.names = FALSE
      )
      
      level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
      level_df$Chi_Square <- chi_square_test$statistic
      level_df$P_Value <- chi_square_test$p.value
      
      # Add the result to the list
      results[[paste0(row_var, "_", level)]] <- level_df
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}

column_variable <- "City"  # Replace with your column variable name

df_chisquare <- chi_square_analysis_multiple(df, "Education.level", column_variable)


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


factor_col <- "City"

factor_levels <- c("Montreal", "Tijuana")

# Call the function
generate_comparative_bar_plot(df, vars_likert, factor_col, factor_levels, font_size = 9, output_filename = "/mnt/data/comparative_bar_plot.png")



## A LIST OF VARIABLES PLOTTED TOGETHER

library(ggplot2)
library(dplyr)
library(tidyr)

# Function to create dot-and-whisker plots with Standard Error
create_mean_se_plot <- function(data, variables, factor_column, save_path = NULL) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    group_by(!!sym(factor_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SE = sd(Value, na.rm = TRUE) / sqrt(sum(!is.na(Value))),  # Calculate Standard Error
      .groups = 'drop'
    )
  
  # Create the plot with mean points and error bars for SE
  p <- ggplot(long_data, aes(x = !!sym(factor_column), y = Mean)) +
    geom_point(aes(color = !!sym(factor_column)), size = 3) +
    geom_errorbar(aes(ymin = Mean - SE, ymax = Mean + SE, color = !!sym(factor_column)), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mean and Standard Error by City", 
      x = "City",  # Rename the x-axis title
      y = "Value",
      color = "City"  # Rename the legend title
    ) +
    scale_color_discrete(name = "City")  # Rename the legend title
  
  # Print the plot
  print(p)
  
  # Save the plot if a save path is provided
  if (!is.null(save_path)) {
    ggsave(filename = save_path, plot = p, width = 8, height = 6)
  }
  
  # Return the plot object
  return(p)
}

# Example usage
plot <- create_mean_se_plot(df, vars_likert, factor_column = "City", save_path = "mean_se_plot.png")


library(ggplot2)
library(dplyr)
library(tidyr)

# Function to create visualizations for chi-square results
visualize_chi_square_results <- function(data, row_var, col_var) {
  # Calculate percentages for the crosstab
  crosstab <- data %>%
    group_by(!!sym(col_var), !!sym(row_var)) %>%
    summarise(Count = n(), .groups = "drop") %>%
    group_by(!!sym(col_var)) %>%
    mutate(Percentage = Count / sum(Count) * 100)
  
  # Chi-square test
  chi_test <- chisq.test(table(data[[row_var]], data[[col_var]]))
  
  # Visualization
  ggplot(crosstab, aes(x = !!sym(col_var), y = Percentage, fill = !!sym(row_var))) +
    geom_bar(stat = "identity", position = "stack", color = "black") +
    geom_text(aes(label = paste0(round(Percentage, 1), "%")),
              position = position_stack(vjust = 0.5), size = 3) +
    labs(
      title = paste("Chi-Square Results for", row_var, "and", col_var),
      subtitle = paste("Chi-Square Statistic:", round(chi_test$statistic, 2), 
                       "| p-value:", signif(chi_test$p.value, 3)),
      x = col_var,
      y = "Percentage (%)",
      fill = row_var
    ) +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 1))
}

# Example usage
visualize_chi_square_results(df, row_var = "Education.level", col_var = "City")



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
  "Data" = df, 
  "Descriptive Stats" = df_descriptive_stats, 
  "Descriptive Stats by Group" = df_descriptive_stats_bygroup,
  "Mann-Shitney" = df_mannwhitney,
  "Chi-Square" = df_chisquare
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

colnames(df)
