setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/patriarchz")

library(readxl)
df <- read_xlsx('FinalData.xlsx')

# Get rid of spaces in column names

names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "", names(df))
names(df) <- gsub("\\)", "", names(df))

vars <- c("Exam_Score")

## DESCRIPTIVE STATS BY EXAMS

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

vars <- c("Exam_Score", "Days_studied_of_31", "Total_reviews","Average_for_days_studied","Average_total","Card_New" ,                "Card_Learning",           
          "Card_Relearning"    ,      "Card_Young"     ,          "Card_Mature" ,            
          "Card_Suspended")
factors <- c("Exam")
calculate_means_and_sds_by_factors(df, vars, factors)
descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df, vars, factors)


## BOXPLOTS

library(ggplot2)

# Specify the factor as a column name
factor_var <- "Exam"  # replace with your actual factor column name

# Loop for generating and saving boxplots
for (var in vars) {
  # Create boxplot for each variable using ggplot2
  p <- ggplot(df, aes_string(x = factor_var, y = var)) +
    geom_boxplot() +
    stat_summary(fun.y = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = paste("Boxplot of", var, "by", factor_var),
         x = factor_var, y = var)
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = paste("boxplot_", var, ".png", sep = ""), plot = p, width = 10, height = 6)
}

# SCATTERPLOTS

library(ggplot2)

# Function to generate scatterplots
generate_scatterplots <- function(data, baseline_var, vars) {
  # Ensure baseline_var is a character string
  baseline_var <- as.character(baseline_var)
  
  # Loop over the variables
  for (var in vars) {
    # Create scatterplot
    p <- ggplot(data, aes_string(x = baseline_var, y = var)) +
      geom_point() +
      labs(title = paste("Scatterplot of", var, "vs", baseline_var),
           x = baseline_var, y = var)
    
    # Print the plot
    print(p)
    
    # Save the plot
    ggsave(filename = paste("scatterplot_", baseline_var, "_vs_", var, ".png", sep = ""), plot = p, width = 10, height = 6)
  }
}

baseline_var = "Exam_Score"
vars_to_plot = c("Days_studied_of_31", "Total_reviews","Average_for_days_studied","Average_total","Card_New" ,        
                  "Card_Young")
generate_scatterplots(df, baseline_var, vars_to_plot)

## CORRELATION ANALYSIS

calculate_correlation_matrix <- function(data, variables, method = "pearson") {
  # Ensure the method is valid
  if (!method %in% c("pearson", "spearman")) {
    stop("Method must be either 'pearson' or 'spearman'")
  }
  
  # Subset the data for the specified variables
  data_subset <- data[variables]
  
  # Calculate the correlation matrix
  corr_matrix <- cor(data_subset, method = method, use = "complete.obs")
  
  # Initialize a matrix to hold formatted correlation values
  corr_matrix_formatted <- matrix(nrow = nrow(corr_matrix), ncol = ncol(corr_matrix))
  rownames(corr_matrix_formatted) <- colnames(corr_matrix)
  colnames(corr_matrix_formatted) <- colnames(corr_matrix)
  
  # Calculate p-values and apply formatting with stars
  for (i in 1:nrow(corr_matrix)) {
    for (j in 1:ncol(corr_matrix)) {
      if (i == j) {  # Diagonal elements (correlation with itself)
        corr_matrix_formatted[i, j] <- format(round(corr_matrix[i, j], 3), nsmall = 3)
      } else {
        p_value <- cor.test(data_subset[[i]], data_subset[[j]], method = method)$p.value
        stars <- ifelse(p_value < 0.01, "***", 
                        ifelse(p_value < 0.05, "**", 
                               ifelse(p_value < 0.1, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}

correlation_matrix <- calculate_correlation_matrix(df, vars, method = "pearson")

# Linear Mixed Model

library(lme4)
library(broom.mixed)
library(afex)

# Function to fit LMM with multiple random effects and return formatted results
fit_lmm_and_format <- function(formula, data, save_plots = FALSE) {
  # Fit the linear mixed-effects model
  lmer_model <- lmer(formula, data = data)
  
  # Print the summary of the model for fit statistics
  print(summary(lmer_model))
  
  # Extract the tidy output and assign it to lmm_results
  lmm_results <- tidy(lmer_model,"fixed")
  
  # Optionally print the tidy output
  print(lmm_results)
  
  # Generate and save residual plots
  if (save_plots) {
    jpeg("Residuals_vs_Fitted.jpeg")
    plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
    dev.off()
    
    jpeg("QQ_Plot.jpeg")
    qqnorm(residuals(lmer_model), main = "Q-Q Plot")
    qqline(residuals(lmer_model))
    dev.off()
  }
  
  return(lmm_results)
}

lmm_results <- fit_lmm_and_format("Exam_Score ~ Total_reviews + (1|Exam) + (1|Student)", df, save_plots = TRUE)

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
  "Descriptive Stats" = descriptive_stats_bygroup, 
  "Correlation" = correlation_matrix, 
  "Model Results" = lmm_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")



## Indicators above and below 90%

# Create subsets for students who used the app (Days_Studied > 0) and those who did not (Days_Studied = 0)
used_app <- df[df$Days_studied_of_31 > 0, ]
did_not_use_app <- df[df$Days_studied_of_31 == 0, ]

# Create subsets after removing students 4 and 10
used_app_no_outliers <- used_app[!(used_app$Student %in% c(4, 10)), ]
did_not_use_app_no_outliers <- did_not_use_app[!(did_not_use_app$Student %in% c(4, 10)), ]

# Calculate the number of exams below and above 90% for each subset
exams_below_90_used <- sum(used_app$Exam_Score < 0.9)
exams_above_90_used <- sum(used_app$Exam_Score >= 0.9)

exams_below_90_did_not_use <- sum(did_not_use_app$Exam_Score < 0.9)
exams_above_90_did_not_use <- sum(did_not_use_app$Exam_Score >= 0.9)

exams_below_90_used_no_outliers <- sum(used_app_no_outliers$Exam_Score < 0.9)
exams_above_90_used_no_outliers <- sum(used_app_no_outliers$Exam_Score >= 0.9)

exams_below_90_did_not_use_no_outliers <- sum(did_not_use_app_no_outliers$Exam_Score < 0.9)
exams_above_90_did_not_use_no_outliers <- sum(did_not_use_app_no_outliers$Exam_Score >= 0.9)

# Calculate the average number of days studied for scores below 90% for each subset
avg_days_below_90_used <- mean(used_app$Days_studied_of_31[used_app$Exam_Score < 0.9], na.rm = TRUE)
avg_days_below_90_did_not_use <- mean(did_not_use_app$Days_studied_of_31[did_not_use_app$Exam_Score < 0.9], na.rm = TRUE)
avg_days_below_90_used_no_outliers <- mean(used_app_no_outliers$Days_studied_of_31[used_app_no_outliers$Exam_Score < 0.9], na.rm = TRUE)
avg_days_below_90_did_not_use_no_outliers <- mean(did_not_use_app_no_outliers$Days_studied_of_31[did_not_use_app_no_outliers$Exam_Score < 0.9], na.rm = TRUE)

# Calculate the average number of days studied for scores at and above 90% for each subset
avg_days_above_90_used <- mean(used_app$Days_studied_of_31[used_app$Exam_Score >= 0.9], na.rm = TRUE)
avg_days_above_90_did_not_use <- mean(did_not_use_app$Days_studied_of_31[did_not_use_app$Exam_Score >= 0.9], na.rm = TRUE)
avg_days_above_90_used_no_outliers <- mean(used_app_no_outliers$Days_studied_of_31[used_app_no_outliers$Exam_Score >= 0.9], na.rm = TRUE)
avg_days_above_90_did_not_use_no_outliers <- mean(did_not_use_app_no_outliers$Days_studied_of_31[did_not_use_app_no_outliers$Exam_Score >= 0.9], na.rm = TRUE)

# Print the results for each subset
cat("Results for Students Who Used the App:\n")
cat("Number of Exams Below 90%:", exams_below_90_used, "\n")
cat("Number of Exams At or Above 90%:", exams_above_90_used, "\n")
cat("Average Days Studied for Scores Below 90%:", avg_days_below_90_used, "\n")
cat("Average Days Studied for Scores At or Above 90%:", avg_days_above_90_used, "\n")

cat("\nResults for Students Who Did Not Use the App:\n")
cat("Number of Exams Below 90%:", exams_below_90_did_not_use, "\n")
cat("Number of Exams At or Above 90%:", exams_above_90_did_not_use, "\n")
cat("Average Days Studied for Scores Below 90%:", avg_days_below_90_did_not_use, "\n")
cat("Average Days Studied for Scores At or Above 90%:", avg_days_above_90_did_not_use, "\n")

cat("\nResults for Students Who Used the App (No Outliers - Students 4 and 10):\n")
cat("Number of Exams Below 90%:", exams_below_90_used_no_outliers, "\n")
cat("Number of Exams At or Above 90%:", exams_above_90_used_no_outliers, "\n")
cat("Average Days Studied for Scores Below 90%:", avg_days_below_90_used_no_outliers, "\n")
cat("Average Days Studied for Scores At or Above 90%:", avg_days_above_90_used_no_outliers, "\n")

cat("\nResults for Students Who Did Not Use the App (No Outliers - Students 4 and 10):\n")
cat("Number of Exams Below 90%:", exams_below_90_did_not_use_no_outliers, "\n")
cat("Number of Exams At or Above 90%:", exams_above_90_did_not_use_no_outliers, "\n")
cat("Average Days Studied for Scores Below 90%:", avg_days_below_90_did_not_use_no_outliers, "\n")
cat("Average Days Studied for Scores At or Above 90%:", avg_days_above_90_did_not_use_no_outliers, "\n")
