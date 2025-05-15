library(openxlsx)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/drpmouser")

df <- read.xlsx("Complete Geltor Clinical dataset in one spreadsheet for stats software.xlsx")

## DATA PROCESSING

library(tidyverse)

df <- df %>%
  pivot_longer(
    cols = c("T0", "T42", "T84"),
    names_to = "Timepoint",
    values_to = "Measurement"
  )

df <- df %>%
  pivot_wider(
    names_from = Parameter, 
    values_from = Measurement,
    id_cols = c(Group, Group.Name, Subject, Timepoint)
  )

# Declare Dependent Variables

dvs <- c(
  "Corneometer forearm",
  "Elasticity Forearm",
  "Elasticity Face",
  "Firmness Forearm",
  "Firmness Face",
  "Wrinkles",
  "Smoothness Sesm Forearm",
  "VFV Forearm",
  "Smoothness Sesm Face",
  "VFV Face"
)

# Update column names to syntactically valid names
names(df) <- make.names(names(df))

# Also update the names in the 'dvs' vector
dvs <- make.names(dvs)

## OUTLIER INSPECTION

# Function to calculate Z-scores and identify outliers
find_outliers <- function(data, vars, threshold = 3) {
  outlier_df <- data.frame(Variable = character(), Outlier_Count = integer())
  outlier_values_df <- data.frame()
  
  for (var in vars) {
    x <- data[[var]]
    z_scores <- (x - mean(x, na.rm = TRUE)) / sd(x, na.rm = TRUE)
    
    # Identify outliers based on the Z-score threshold
    outliers <- which(abs(z_scores) > threshold)
    
    # Count outliers
    count_outliers <- length(outliers)
    
    # Store results in a summary dataframe
    outlier_df <- rbind(outlier_df, data.frame(Variable = var, Outlier_Count = count_outliers))
    
    # Store outlier values and Z-scores in a separate dataframe
    if (count_outliers > 0) {
      tmp_df <- data.frame(data[outliers, ], Z_Score = z_scores[outliers])
      tmp_df$Variable <- var
      outlier_values_df <- rbind(outlier_values_df, tmp_df)
      
      print(paste("Variable:", var, "has", count_outliers, "outliers with Z-scores greater than", threshold))
    }
  }
  
  return(list(summary = outlier_df, values = outlier_values_df))
}

# Execute the function
outlier_summary <- find_outliers(df, dvs, threshold = 3)

# Splitting the list into two separate dataframes
outlier_count <- outlier_summary$summary
outlier_values_df <- outlier_summary$values
  
# Check Boxplots
library(ggplot2)

# Generate separate boxplots for each dependent variable
for (dv in dvs) {
  plot <- ggplot(df, aes(x = Timepoint, y = !!sym(dv), fill = Group.Name)) +
    geom_boxplot() +
    labs(
      title = paste("Boxplot of", dv),
      x = "Timepoint",
      y = "Measurement Value"
    ) +
    theme_minimal()
  
  # Show plot
  print(plot)
  
  # Save plot to a file if needed
  ggsave(filename = paste0("boxplot_", dv, ".png"), plot = plot)
}

## DATA DISTRIBUTION INSPECTION

# Function to calculate normality statistics for each variable
library(moments)
library(stats)

calc_descriptive_stats_for_dvs <- function(data, dvs) {
  stats_df <- data.frame(Variable = character(), Skewness = numeric(), 
                         Kurtosis = numeric(), W = numeric(), P_Value = numeric())
  
  for (var in dvs) {
    x <- data[[var]]
    x <- x[!is.na(x)]  # Remove NA values if present
    
    if (length(x) < 3) {
      next  # Skip if less than 3 observations
    }
    
    skewness_value <- skewness(x)
    kurtosis_value <- kurtosis(x)
    shapiro_test <- shapiro.test(x)
    W_value <- shapiro_test$statistic
    p_value <- shapiro_test$p.value
    
    stats_df <- rbind(stats_df, data.frame(Variable = var, Skewness = skewness_value, 
                                           Kurtosis = kurtosis_value, W = W_value, P_Value = p_value))
  }
  
  return(stats_df)
}

# Use function to get stats for all dvs
stats_df <- calc_descriptive_stats_for_dvs(df, dvs)

# CHECK HOMOGENEITY OF VARIANCES

# Levene's test
library(car)

# Declare the factor variable
factor_var <- "Group.Name"

# Initialize an empty dataframe to store the results
levenes_results <- data.frame(
  Farm = character(),
  Dependent_Variable = character(),
  Factor = character(),
  F_value = numeric(),
  DF1 = numeric(),
  DF2 = numeric(),
  pValue = numeric(),
  stringsAsFactors = FALSE
)

# Loop through each dependent variable and apply Levene's test
for (var in dvs) {
  
  # Conduct Levene's Test
  levenes_test <- leveneTest(as.formula(paste(var, "~", factor_var)), data = df)
  
  # Extract relevant statistics
  F_value <- levenes_test[1, "F value"]
  DF1 <- levenes_test[1, "Df"]
  DF2 <- levenes_test[2, "Df"]
  pValue <- levenes_test[1, "Pr(>F)"]
  
  # Store the results in the dataframe
  levenes_results <- rbind(levenes_results, data.frame(
    Dependent_Variable = var,
    Factor = factor_var,
    F_value = F_value,
    DF1 = DF1,
    DF2 = DF2,
    pValue = pValue,
    stringsAsFactors = FALSE
  ))
}

## DESCRIPTIVE STATISTICS

within_subject_factor <- "Timepoint"
between_subject_factor <- "Group.Name"

library(dplyr)

# Loop through each dependent variable
for (var in dvs) {
  
  # Calculate mean scores for each group and timepoint
  temp_mean_df <- df %>%
    group_by(!!sym(between_subject_factor), !!sym(within_subject_factor)) %>%
    summarise(Mean = mean(!!sym(var), na.rm = TRUE))
  
  # Create the line plot
  p <- ggplot(temp_mean_df, aes_string(x = within_subject_factor, y = "Mean", group = between_subject_factor, color = between_subject_factor)) +
    geom_line() +
    geom_point(size = 4) +
    labs(
      title = paste("Line Plot of", var),
      x = within_subject_factor,
      y = "Mean Value"
    )
  
  # Display the plot
  print(p)
}

# Initialize an empty data frame to store the mean values
mean_values_df <- data.frame()

# Loop through each dependent variable
for (var in dvs) {
  
  # Calculate mean scores for each group and timepoint
  temp_df <- df %>%
    group_by(!!sym(between_subject_factor), !!sym(within_subject_factor)) %>%
    summarise(Mean_Value = mean(!!sym(var), na.rm = TRUE)) %>%
    spread(!!sym(within_subject_factor), Mean_Value) %>% # convert to wide format with timepoints as columns
    add_column(Variable = var, .before = 1) # add the dependent variable name
  
  # Add the mean values to the final data frame
  mean_values_df <- bind_rows(mean_values_df, temp_df)
}

# Display the mean_values_df
print(mean_values_df)


## LINEAR MIXED MODEL
library(nlme)
library(car)  # for Anova()

# Initialize an empty dataframe to store the results
results_df <- data.frame()

# Loop through each dependent variable in 'dvs'
for (var in dvs) {
  # Define the mixed ANOVA model formula
  formula_str <- as.formula(paste(var, "~", between_subject_factor, "*", within_subject_factor))
  
  # Run mixed ANOVA
  mixed_anova_model <- lme(formula_str, random = ~1 | Subject, data = df, method = "REML")
  
  # Run type-III ANOVA
  anova_results <- Anova(mixed_anova_model, type = "III")
  
  # Extract relevant statistics
  model_results <- data.frame(
    Dependent_Variable = var,
    Effect = row.names(anova_results),
    F_value = anova_results$`F value`,
    DF = anova_results$`Df`,
    pValue = anova_results$`Pr(>F)`
  )
  
  # Append the results to the overall results dataframe
  results_df <- rbind(results_df, model_results)
}

# View the results dataframe
print(results_df)