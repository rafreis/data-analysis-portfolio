library(openxlsx)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/toeberto")

df <- read.xlsx("Data.xlsx")

## OUTLIER INSPECTION

dvs <- c("%.imported.material.of.roof",
         "%.imported.floor.material",
         "%.imported.material.of.wall",
         "%.subsistence.type.of.lighting",
         "%.Department.Born.not.from.Gracias.a.Dios",
         "%.Department.lived.5.years.ago.not.from.Gracias.a.Dios",
         "%.of.Expensive.Type.of.Bathroom",
         "%.No.Land.Change.Job",
         "%.expensive.type.of.lighting.(Generator,Solar.Panel)",
         "Own.Vehicle.No%",
         "Own.Refrigerator.No%",
         "Own.Television.No%",
         "In.home.Internet.Service.No%",
         "Own.Video.Games.(Playstation,.Wii).No%")
         

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
      tmp_df <- data.frame(Departamento = data[outliers, 'Departamento'],
                           Municipio = data[outliers, 'Municipio'],
                           Year = data[outliers, 'Year'],
                           Group = data[outliers, 'Group'],
                           DV_Value = data[outliers, var],
                           Z_Score = z_scores[outliers],
                           Variable = var)
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

# Check for duplicate column names
duplicate_columns <- names(df)[duplicated(names(df))]
if (length(duplicate_columns) > 0) {
  print(paste("Duplicate columns found:", paste(duplicate_columns, collapse = ", ")))
} else {
  print("No duplicate columns found.")
}

# Make column names unique
names(df) <- make.unique(names(df))

df$Year <- as.factor(df$Year)

# Check Boxplots
library(ggplot2)

# Convert Year and Group to factors
df$Year <- as.factor(df$Year)
df$Group <- as.factor(df$Group)

# Generate separate boxplots for each dependent variable
for (dv in dvs) {
  # Remove rows where the dependent variable is NA
  df_filtered <- df[!is.na(df[[dv]]), ]
  
  # Generate the boxplot
  plot <- ggplot(df_filtered, aes(x = Year, y = !!sym(dv), fill = Group)) +
    geom_boxplot() +
    labs(
      title = paste("Boxplot of", dv),
      x = "Year",
      y = "Measurement Value"
    ) +
    theme_minimal()
  
  # Show plot
  print(plot)
  
  # Save plot to a file
  ggsave(filename = paste0("boxplot_", dv, ".png"), plot = plot)
}


## DESCRIPTIVE STATISTICS

within_subject_factor <- "Year"
between_subject_factor <- "Group"

library(dplyr)
library(tidyverse)

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
  
  # Save plot to a file
  ggsave(filename = paste0("lineplot_", var, ".png"), plot = p)
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


## DATA DISTRIBUTION CHECK
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

normality_stats_df <- calc_descriptive_stats_for_dvs(df, dvs)

# Log-transform the dependent variables and handle zero or negative values
for (var in dvs) {
  new_var_name <- paste0("log_", var)
  df[[new_var_name]] <- ifelse(df[[var]] > 0, log(df[[var]]), NA)
}

# Update the list of dependent variables to include the log-transformed variables
log_dvs <- paste0("log_", dvs)

# Calculate normality statistics for the log-transformed variables
normality_stats_df_log <- calc_descriptive_stats_for_dvs(df, log_dvs)


## FINAL SET OF DVS

# Identify outliers for '% subsistence type of lighting'
outlier_info_subsistence_lighting <- find_outliers(df, c("%.subsistence.type.of.lighting"), threshold = 3)

# If there are outliers, remove them
if(nrow(outlier_info_subsistence_lighting$values) > 0){
  outlier_rows <- rownames(outlier_info_subsistence_lighting$values)
  df <- df[-as.numeric(outlier_rows), ]
}

calc_descriptive_stats_for_dvs(df, "%.subsistence.type.of.lighting")

# Create a list of final DVs for mixed models
final_dvs <- setdiff(log_dvs, "log_%.subsistence.type.of.lighting")
final_dvs <- c(final_dvs, "%.subsistence.type.of.lighting")

## LINEAR MIXED MODEL
# Include necessary libraries
library(nlme)
library(dplyr)

# Initialize an empty dataframe to store model summary statistics
model_summary_df <- data.frame()

# Get unique group names
group_names <- unique(df$Group)

# Loop through each dependent variable
for (var in final_dvs) {
  
  new_var_name <- paste0("TempDV_", gsub("[%.]", "_", var))  # Replace special characters
  names(df)[names(df) == var] <- new_var_name
  
  # Loop through each unique pair of groups
  for (i in 1:(length(group_names) - 1)) {
    for (j in (i + 1):length(group_names)) {
      
      # Filter data for the current pair of groups and remove NAs
      df_pair <- df %>% 
        filter(Group %in% c(group_names[i], group_names[j])) %>%
        filter(!is.na(!!sym(new_var_name)))
      
      # Fit the linear mixed-effects model using nlme
      model_formula <- reformulate(termlabels = "Group * Year", response = new_var_name)
      model <- lme(model_formula, random = ~1 | Departamento/Municipio, data = df_pair, method = "REML")
      
      # Get ANOVA table
      anova_table <- anova(model)
      
      # Create a temporary dataframe to store the results
      model_summary_temp <- data.frame(
        Dependent_Variable = rep(var, nrow(anova_table)),
        Group_Pair = paste(group_names[i], "-", group_names[j]),
        Effect = row.names(anova_table),
        numDF = anova_table$numDF,
        denDF = anova_table$denDF,
        F_value = anova_table$`F-value`,
        pValue = anova_table$`p-value`,
        stringsAsFactors = FALSE
      )
      
      # Append the results to the main summary dataframe
      model_summary_df <- rbind(model_summary_df, model_summary_temp)
    }
  }
  
  names(df)[names(df) == new_var_name] <- var  # Revert to the original variable name
}

# Export to Excel
library(writexl)

normality_stats_df_post <- calc_descriptive_stats_for_dvs(df, final_dvs)

# Create a list of data frames to write to Excel
list_of_dfs <- list("Mean Values" = mean_values_df, "Model Summary" = model_summary_df, "Normality Assessment - Initial" = normality_stats_df, "Normality Assessment - Post" = normality_stats_df_post)


# Write to Excel
write_xlsx(list_of_dfs, "Results.xlsx")


