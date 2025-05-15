setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/teacherdata")

library(openxlsx)
df <- read.xlsx("2021-22 2022-23 2023-24.xlsx")

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

## DESCRIPTIVE TABLES

## DESCRIPTIVE STATS BY FACTORS - ONLY ONE FACTOR


library(dplyr)

calculate_descriptive_stats_bygroups <- function(data, desc_vars, cat_column) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each category
  for (var in desc_vars) {
    # Group data by the categorical column
    grouped_data <- data %>%
      group_by(!!sym(cat_column)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE)
        
      ) %>%
      mutate(Variable = var)
    
    # Append the results for each variable and category to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}

df_descriptive_byperiod <- calculate_descriptive_stats_bygroups(df,"StdScore", "Period")
df_descriptive_byclass <- calculate_descriptive_stats_bygroups(df, "StdScore", "Class")
df_descriptive_bytest <- calculate_descriptive_stats_bygroups(df, "StdScore", "Test")


## DESCRIPTIVE STATS BY FACTORS - TWO FACTORS AS LEVELS OF COLUMNS

library(dplyr)

calculate_descriptive_stats_bygroups2 <- function(data, desc_vars, cat_column1, cat_column2) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category1 = character(),
    Category2 = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each combination of categories
  for (var in desc_vars) {
    # Group data by both categorical columns
    grouped_data <- data %>%
      group_by(!!sym(cat_column1), !!sym(cat_column2)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE),
        .groups = 'drop' # Drop grouping structure after summarising
      ) %>%
      mutate(Variable = var) %>%
      rename(Category1 = !!sym(cat_column1), Category2 = !!sym(cat_column2))
    
    # Append the results for each variable and category combination to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}

df_descriptive_byclassperiod <- calculate_descriptive_stats_bygroups2(df,"StdScore", "Class", "Period")
df_descriptive_byclasstest <- calculate_descriptive_stats_bygroups2(df,"StdScore", "Class", "Test")

calculate_descriptive_stats_bygroups3 <- function(data, desc_vars, cat_column1, cat_column2, cat_column3) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category1 = character(),
    Category2 = character(),
    Category3 = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each combination of categories
  for (var in desc_vars) {
    # Group data by all three categorical columns
    grouped_data <- data %>%
      group_by(!!sym(cat_column1), !!sym(cat_column2), !!sym(cat_column3)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE),
        .groups = 'drop' # Drop grouping structure after summarising
      ) %>%
      mutate(Variable = var) %>%
      rename(
        Category1 = !!sym(cat_column1), 
        Category2 = !!sym(cat_column2), 
        Category3 = !!sym(cat_column3)
      )
    
    # Append the results for each variable and category combination to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}

# Example usage with three categorical columns
df_descriptive_byclassperiodtest <- calculate_descriptive_stats_bygroups3(df, "StdScore", "Class", "Period", "Test")

#Visuals

library(ggplot2)
library(dplyr)
library(tidyr)

create_mean_ci_plot <- function(data, variables, factor_column, y_limits = NULL) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    group_by(!!sym(factor_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      N = n(),
      SD = sd(Value, na.rm = TRUE),
      SEM = SD / sqrt(N),
      .groups = 'drop'
    ) %>%
    mutate(
      CI = qt(0.975, df = N-1) * SEM,  # 95% confidence interval
      Lower = Mean - CI,
      Upper = Mean + CI
    )
  
  # Create the plot with mean points and error bars for 95% CI
  p <- ggplot(long_data, aes(x = !!sym(factor_column), y = Mean)) +
    geom_point(aes(color = !!sym(factor_column)), size = 3) +
    geom_errorbar(aes(ymin = Lower, ymax = Upper, color = !!sym(factor_column)), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      
      x = "Period",  # Rename the x-axis title
      y = "Value",
      color = factor_column  # Set the legend title to the factor column
    ) +
    scale_color_discrete(name = factor_column)  # Ensure this matches your intended legend title
  
  # Set y-axis limits if provided
  if (!is.null(y_limits)) {
    p <- p + ylim(y_limits)
  }
  
  # Print the plot
  return(p)
}

# Create a new column that concatenates Class and ExamType
df$ClassExamType <- paste(df$Class, df$ExamType, sep="_")

# Single categorical column plot
p1 <- create_mean_ci_plot(df, "StdScore", "Period", c(0.5,0.8))
print(p1)

create_mean_ci_plot_nested <- function(data, variables, factor_column1, factor_column2, y_limits = NULL) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    group_by(!!sym(factor_column1), !!sym(factor_column2), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      N = n(),
      SD = sd(Value, na.rm = TRUE),
      SEM = SD / sqrt(N),
      .groups = 'drop'
    ) %>%
    mutate(
      CI = qt(0.975, df = N-1) * SEM,  # 95% confidence interval
      Lower = Mean - CI,
      Upper = Mean + CI
    )
  
  # Create the plot with mean points and error bars for 95% CI
  p <- ggplot(long_data, aes(x = !!sym(factor_column2), y = Mean, color = !!sym(factor_column1))) +
    geom_point(position = position_dodge(width = 0.75), size = 3) +
    geom_errorbar(aes(ymin = Lower, ymax = Upper), position = position_dodge(width = 0.75), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      x = "Period",  # Use Year for x-axis title
      y = "Value",
      color = "Class - Test"  # Use Class - Exam for legend title
    ) +
    scale_color_discrete(name = "Class - Test", 
                         labels = c("Class 1 FC", "Class 1 T", "Class 2 FC", "Class 2 T", "Class 3 FC", "Class 3 T"))  # Ensure this matches your intended legend title and labels
  
  # Set y-axis limits if provided
  if (!is.null(y_limits)) {
    p <- p + ylim(y_limits)
  }
  
  # Print the plot
  return(p)
}


# Two nested categorical columns plot
p2 <- create_mean_ci_plot_nested(df, "StdScore", "ClassExamType", "Period")
print(p2)

# Model

library(lme4)
library(broom.mixed)
library(afex)
library(car)

# Extract ExamType from Class column and redefine Class column
df$ExamType <- ifelse(grepl("FC", df$Class), "FC", "T")
df$Class <- sub("FC|T", "", df$Class)

# Function to fit LMM with multiple random effects and return formatted results
fit_lmm_and_format <- function(data, save_plots = FALSE, plot_param = "") {
  # Create a unique identifier for each student within each class
  data$Class_Student <- interaction(data$Class, data$Student)
  
  # Define the formula
  formula <- StdScore ~ Period + (1 | Class) + (1 | Test) + (1 | Class_Student) + (1 | ExamType)
  
  # Fit the linear mixed-effects model
  lmer_model <- lmer(formula, data = data)
  
  # Print the summary of the model for fit statistics
  model_summary <- summary(lmer_model)
  print(model_summary)
  
  # Extract the tidy output for fixed effects
  lmm_results_fixed <- tidy(lmer_model, "fixed")
  
  # Extract random effects variances
  random_effects_variances <- as.data.frame(VarCorr(lmer_model))
  random_effects_variances <- random_effects_variances %>%
    dplyr::rename(Term = grp, Variance = vcov) %>%
    dplyr::select(Term, Variance)
  
  # Optionally print the results
  print(lmm_results_fixed)
  print(random_effects_variances)
  
  
  
  # Generate and save residual plots
  if (save_plots) {
    residual_plot_filename <- paste0("Residuals_vs_Fitted_", plot_param, ".jpeg")
    qq_plot_filename <- paste0("QQ_Plot_", plot_param, ".jpeg")
    
    jpeg(residual_plot_filename)
    plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
    dev.off()
    
    jpeg(qq_plot_filename)
    qqnorm(residuals(lmer_model), main = "Q-Q Plot")
    qqline(residuals(lmer_model))
    dev.off()
  }
  
  # Combine fixed and random effects results into a single data frame
  lmm_results_fixed <- lmm_results_fixed %>%
    dplyr::mutate(Type = "Fixed")
  
  random_effects_variances <- random_effects_variances %>%
    dplyr::mutate(Type = "Random")
  
  combined_results <- dplyr::bind_rows(lmm_results_fixed, random_effects_variances)
  
  # Return the results as a data frame
  return(combined_results)
}

# Fit the model and get results
results <- fit_lmm_and_format(df, save_plots = TRUE, plot_param = "Period")


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
  "Cleaned Data" = df,
  "by Class" = df_descriptive_byclass,
  "by Period" = df_descriptive_byperiod,
  "by Test" = df_descriptive_bytest,
  "by Class and Period" = df_descriptive_byclassperiod,
  "by Class and Test" = df_descriptive_byclasstest,
  "by Class and Period and Test" = df_descriptive_byclassperiodtest,
  "LMM" = results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")
