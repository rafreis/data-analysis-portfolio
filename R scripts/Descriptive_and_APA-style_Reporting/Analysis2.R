library(openxlsx)
library(ggplot2)
library(dplyr)
library(tidyr)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/nicolecow")

# Reading the 1st sheet
data <- read.xlsx("Trial data for analysis.xlsx", sheet = 1)






## INITIAL NORMALITY CHECK
library(moments)

# Function to calculate normality statistics for each variable
calc_descriptive_stats <- function(x) {
  skewness_value <- skewness(x, na.rm = TRUE)
  kurtosis_value <- kurtosis(x, na.rm = TRUE)
  shapiro_test <- shapiro.test(x)
  W_value <- shapiro_test$statistic
  p_value <- shapiro_test$p.value
  
  return(data.frame(Skewness = skewness_value, Kurtosis = kurtosis_value, W = W_value, P_Value = p_value))
}

# Apply function

milk_stats <- calc_descriptive_stats(data$Milk.Kg)
fat_stats <- calc_descriptive_stats(data$Fat.Kg)
protein_stats <- calc_descriptive_stats(data$Protein.Kg)
scc_stats <- calc_descriptive_stats(data$SCC)

# Combine the results into a single data frame
normality_df <- rbind(milk_stats, fat_stats, protein_stats, scc_stats)
rownames(normality_df) <- c("Milk.Kg", "Fat.Kg", "Protein.Kg", "SCC")

# Histograms

# Histogram for Milk.Kg
ggplot(data, aes(x = Milk.Kg)) +
  geom_histogram(binwidth = 1, fill = "blue", alpha = 0.7) +
  ggtitle("Histogram of Milk.Kg") +
  xlab("Milk.Kg") +
  ylab("Frequency")

# Histogram for Fat.Kg
ggplot(data, aes(x = Fat.Kg)) +
  geom_histogram(binwidth = 0.1, fill = "green", alpha = 0.7) +
  ggtitle("Histogram of Fat.Kg") +
  xlab("Fat.Kg") +
  ylab("Frequency")

# Histogram for Protein.Kg
ggplot(data, aes(x = Protein.Kg)) +
  geom_histogram(binwidth = 0.05, fill = "red", alpha = 0.7) +
  ggtitle("Histogram of Protein.Kg") +
  xlab("Protein.Kg") +
  ylab("Frequency")

# Histogram for SCC
ggplot(data, aes(x = SCC)) +
  geom_histogram(binwidth = 0.05, fill = "red", alpha = 0.7) +
  ggtitle("Histogram of SCC") +
  xlab("SCC") +
  ylab("Frequency")


## OUTLIER EXAMINATIon

# Create boxplots for the variables
ggplot(data, aes(x = "", y = Milk.Kg)) + geom_boxplot() + ggtitle("Boxplot of Milk.Kg")
ggplot(data, aes(x = "", y = Fat.Kg)) + geom_boxplot() + ggtitle("Boxplot of Fat.Kg")
ggplot(data, aes(x = "", y = Protein.Kg)) + geom_boxplot() + ggtitle("Boxplot of Protein.Kg")
ggplot(data, aes(x = "", y = SCC)) + geom_boxplot() + ggtitle("Boxplot of SCC")

# Function to calculate Z-scores
calc_zscore <- function(x) {
  (x - mean(x, na.rm = TRUE)) / sd(x, na.rm = TRUE)
}

# Calculate Z-scores for the variables
data$Milk.Kg_Zscore <- calc_zscore(data$Milk.Kg)
data$Fat.Kg_Zscore <- calc_zscore(data$Fat.Kg)
data$Protein.Kg_Zscore <- calc_zscore(data$Protein.Kg)
data$SCC_Zscore <- calc_zscore(data$SCC)

# Filter data to only include rows where the Z-score is greater than 3 or smaller than -3
outliers <- data[abs(data$Milk.Kg_Zscore) > 3 | abs(data$Fat.Kg_Zscore) > 3 | abs(data$Protein.Kg_Zscore) > 3 | abs(data$SCC_Zscore) > 3, ]

## OUTLIER HANDLING
# Log-transform the SCC variable
data$SCC_log <- log(data$SCC)

# Create boxplot for the log-transformed SCC variable
ggplot(data, aes(x = "", y = SCC_log)) + geom_boxplot() + ggtitle("Boxplot of Log-Transformed SCC")

# Calculate Z-scores for the log-transformed SCC
data$SCC_log_Zscore <- calc_zscore(data$SCC_log)

# Evaluate normality of log-transformed variable
scc_stats <- calc_descriptive_stats(data$SCC_log)

# Histogram for log-transformed SCC
ggplot(data, aes(x = SCC_log)) +
  geom_histogram(binwidth = 0.1, fill = "purple", alpha = 0.7) +
  ggtitle("Histogram of Log-transformed SCC") +
  xlab("Log-transformed SCC") +
  ylab("Frequency")

## DESCRIPTIVE STATISTICS - CROSSTABULATIONS

# Make 'Time' a factor with ordered levels
data$Time <- factor(data$Time, levels = c("Pre_Treatment", "Mid_Lactation", "Early_Lactation"), ordered = TRUE)

# Function to create crosstabulation tables for each Farm
create_crosstab_per_farm <- function(data, variable_name) {
  
  for (farm in unique(data$Farm)) {
    temp_data <- dplyr::filter(data, Farm == farm)
    
    temp_table <- temp_data %>%
      dplyr::group_by(Time, Population) %>%
      dplyr::summarise(mean_value = round(mean(!!rlang::sym(variable_name), na.rm = TRUE), 3)) %>%
      tidyr::spread(key = Population, value = mean_value)
    
    # Create a unique name for the dataframe for each farm and variable
    df_name <- paste0(variable_name, "_", farm, "_crosstab")
    
    # Assign the data frame to a variable in the global environment
    assign(df_name, temp_table, envir = .GlobalEnv)
  }
}

# Create crosstabs for each of the four variables and each Farm
create_crosstab_per_farm(data, "Milk.Kg")
create_crosstab_per_farm(data, "Fat.Kg")
create_crosstab_per_farm(data, "Protein.Kg")
create_crosstab_per_farm(data, "SCC_log")

# Function to create a crosstabulation table for all farms
create_crosstab_all_farms <- function(data, variable_name) {
  
  temp_table <- data %>%
    dplyr::group_by(Time, Population) %>%
    dplyr::summarise(mean_value = round(mean(!!rlang::sym(variable_name), na.rm = TRUE), 3)) %>%
    tidyr::spread(key = Population, value = mean_value)
  
  # Create a unique name for the dataframe for the variable
  df_name <- paste0(variable_name, "_all_farms_crosstab")
  
  # Assign the data frame to a variable in the global environment
  assign(df_name, temp_table, envir = .GlobalEnv)
}

# Create crosstabs for each of the four variables for all farms together
create_crosstab_all_farms(data, "Milk.Kg")
create_crosstab_all_farms(data, "Fat.Kg")
create_crosstab_all_farms(data, "Protein.Kg")
create_crosstab_all_farms(data, "SCC_log")


# LEVENE'S TEST - FARMS
library(car)

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

# Dependent variables
dependent_vars <- c("Milk.Kg", "Fat.Kg", "Protein.Kg", "SCC")

# Farms
farms <- c("A", "B", "C", "D")

# Factors to test
factors <- c("Time", "Population")

# Function to run Levene's test for a given dependent variable, farm, and factor
run_levenes_test <- function(data, farm, dependent_var, factor) {
  # Filter data for the given farm
  data_farm <- subset(data, Farm == farm)
  
  # Remove rows where the dependent variable is NA
  data_farm_clean <- data_farm[!is.na(data_farm[[dependent_var]]),]
  
  # Run Levene's Test
  formula_str <- paste(dependent_var, "~", factor)
  levene_result <- leveneTest(as.formula(formula_str), data = data_farm_clean)
  
  return(levene_result)
}

# Function to run Levene's test for a given dependent variable and factor
run_levenes_test_overall <- function(data, dependent_var, factor) {
  # Remove rows where the dependent variable is NA
  data_clean <- data[!is.na(data[[dependent_var]]),]
  
  # Run Levene's Test
  formula_str <- paste(dependent_var, "~", factor)
  levene_result <- leveneTest(as.formula(formula_str), data = data_clean)
  
  return(levene_result)
}

# Loop through each farm, dependent variable, and factor to run Levene's Test
for (farm in farms) {
  for (dv in dependent_vars) {
    for (factor in factors) {
      result <- run_levenes_test(data, farm, dv, factor)
      
      new_row <- data.frame(
        Farm = farm,
        Dependent_Variable = dv,
        Factor = factor,
        F_value = as.numeric(result[1, "F value"]),
        DF1 = as.numeric(result[1, "Df"]),
        DF2 = as.numeric(result[2, "Df"]),
        pValue = as.numeric(result[1, "Pr(>F)"]),
        stringsAsFactors = FALSE
      )
      levenes_results <- rbind(levenes_results, new_row)
    }
  }
}

# Loop through each dependent variable and factor to run Levene's Test for all farms
for (dv in dependent_vars) {
  for (factor in factors) {
    result <- run_levenes_test_overall(data, dv, factor)
    
    new_row <- data.frame(
      Farm = "Overall",
      Dependent_Variable = dv,
      Factor = factor,
      F_value = as.numeric(result[1, "F value"]),
      DF1 = as.numeric(result[1, "Df"]),
      DF2 = as.numeric(result[2, "Df"]),
      pValue = as.numeric(result[1, "Pr(>F)"]),
      stringsAsFactors = FALSE
    )
    levenes_results <- rbind(levenes_results, new_row)
  }
}

## MIXED ANOVA

# Install and load the packages
library(nlme)

# Initialize an empty dataframe to store the results
results_df <- data.frame()

# Function to run mixed ANOVA for a given dependent variable and farm
run_mixed_anova <- function(data, farm, dependent_var) {
  
  # Filter data for the given farm
  data_farm <- subset(data, Farm == farm)
  
  # Remove rows where the dependent variable is NA
  data_farm_clean <- na.omit(data_farm[, c("Cow.ID", "Time", "Population", dependent_var)])
  
  # Run mixed ANOVA
  formula_str <- as.formula(paste(dependent_var, "~ Time * Population"))
  mixed_anova_model <- lme(formula_str, random = ~1 | Cow.ID/Time, data = data_farm_clean, method = "REML")
  
  # Summary of the model
  model_summary <- summary(mixed_anova_model)
  
  # Extract relevant statistics
  fixed_effects <- model_summary$tTable
  model_results <- data.frame(
    Farm = farm,
    Dependent_Variable = dependent_var,
    Effect = rownames(fixed_effects),
    Value = fixed_effects[, "Value"],
    Std.Error = fixed_effects[, "Std.Error"],
    DF = fixed_effects[, "DF"],
    tValue = fixed_effects[, "t-value"],
    pValue = fixed_effects[, "p-value"]
  )
  
  return(model_results)
}

# Dependent variables
dependent_vars <- c("Milk.Kg", "Fat.Kg", "Protein.Kg", "SCC_log")

# Farms
farms <- c("A", "B", "C", "D")

# Loop through each farm and each dependent variable to run mixed ANOVA
for (farm in farms) {
  for (dv in dependent_vars) {
    model_results <- run_mixed_anova(data, farm, dv)
    results_df <- rbind(results_df, model_results)
  }
}

# Now for all farms combined
for (dv in dependent_vars) {
  data_clean <- na.omit(data[, c("Cow.ID", "Time", "Population", dv)])
  formula_str <- as.formula(paste(dv, "~ Time * Population"))
  mixed_anova_model_all_farms <- lme(formula_str, random = ~1 | Cow.ID/Time, data = data_clean, method = "REML")
  model_summary_all_farms <- summary(mixed_anova_model_all_farms)
  
  fixed_effects <- model_summary_all_farms$tTable
  model_results_all_farms <- data.frame(
    Farm = "All",
    Dependent_Variable = dv,
    Effect = rownames(fixed_effects),
    Value = fixed_effects[, "Value"],
    Std.Error = fixed_effects[, "Std.Error"],
    DF = fixed_effects[, "DF"],
    tValue = fixed_effects[, "t-value"],
    pValue = fixed_effects[, "p-value"]
  )
  
  results_df <- rbind(results_df, model_results_all_farms)
}

# View the results dataframe
print(results_df)




## POST-HOC TESTS

# Load necessary packages
library(multcomp)
library(nparcomp)

# Initialize an empty dataframe to store the post-hoc results
posthoc_results_df <- data.frame(Farm = character(),
                                 Dependent_Variable = character(),
                                 Factor = character(),
                                 Comparison = character(),
                                 p_value = numeric())


# Function to run appropriate post-hoc tests
run_posthoc_test <- function(data, farm, dv, factor, levenes_results, anova_df, posthoc_results_df) {
  print(paste("Running for farm: ", farm, ", DV: ", dv, ", Factor: ", factor))
  
  levenes_row <- levenes_results %>% filter(Farm == farm, Dependent_Variable == dv, Factor == factor)
  
  anova_row <- anova_df %>% filter(Farm == farm, Dependent_Variable == dv, startsWith(Effect, factor))
  
  if(nrow(anova_row) == 0) {
    print("No anova row for this combination")
    return(posthoc_results_df)
  }
  
  anova_significant <- any(anova_row$pValue < 0.05)
  
  if(anova_significant) {
    levene_significant <- any(levenes_row$pValue < 0.05)
    if(levene_significant) {
      posthoc_result <- pairwise.t.test(data[[dv]], data[[factor]], p.adjust.method = "none", pool.sd = FALSE)
    } else {
      posthoc_result <- pairwise.t.test(data[[dv]], data[[factor]], p.adjust.method = "none", pool.sd = TRUE)
    }
    
    # Extract comparisons and p-values
    comps <- rownames(posthoc_result$p.value)
    pvals <- as.vector(posthoc_result$p.value)
    
    # Create a data frame for these results
    new_rows <- data.frame(Farm = rep(farm, length(comps)),
                           Dependent_Variable = rep(dv, length(comps)),
                           Factor = rep(factor, length(comps)),
                           Comparison = comps,
                           p_value = pvals)
    
    # Append to the existing dataframe
    posthoc_results_df <- rbind(posthoc_results_df, new_rows)
  }
  return(posthoc_results_df)
}

# Loop through the unique farms, dependent variables, and factors
for (farm in unique(data$Farm)) {
  for (dv in dependent_vars) {
    for (factor in factors) {
      posthoc_results_df <- run_posthoc_test(data, farm, dv, factor, levenes_results, results_df, posthoc_results_df)
    }
  }
}

# For all farms together
for (dv in dependent_vars) {
  for (factor in factors) {
    posthoc_results_df <- run_posthoc_test(data, "All", dv, factor, levenes_results, results_df, posthoc_results_df)
  }
}

## INTERACTION GRAPHS

# Loop through the unique farms and dependent variables
for (dv in dependent_vars) {
  print(paste("Processing dependent variable:", dv))
  
  # Calculate the summary statistics for the current dependent variable
  summary_data <- data %>%
    group_by(Farm, Time, Population) %>%
    summarise(
      mean_value = mean(!!sym(dv), na.rm = TRUE),
      se = sd(!!sym(dv), na.rm = TRUE) / sqrt(n())
    )
  
  print("Summary data for this DV:")
  print(head(summary_data))
  
  for (farm in unique(summary_data$Farm)) {
    print(paste("Processing farm:", farm))
    
    # Subset the data for the farm of interest
    sub_data <- subset(summary_data, Farm == farm)
    
    print("Subset data for this farm:")
    print(head(sub_data))
    
    # Create the interaction plot using ggplot2
    p <- ggplot(sub_data, aes(x=factor(Time, levels=c("Pre_Treatment", "Early_Lactation", "Mid_Lactation")), y=mean_value, color=Population, linetype=Population)) +
      geom_errorbar(aes(ymin=mean_value-se, ymax=mean_value+se), width=0.2) +
      geom_point() +
      geom_line() +
      labs(
        title = paste("Interaction Plot for Farm", farm, "and DV", dv),
        x = "Time",
        y = paste("Mean", dv)
      ) +
      theme_minimal() +
      theme(
        text = element_text(family = "Times New Roman", size = 12),
        plot.title = element_text(hjust = 0.5, size = 14, family = "Times New Roman"),
        legend.position = "bottom"
      )
    
    print(p)
  }
}

# Calculate the summary statistics for the entire dataset (across all farms)
summary_data_all_farms <- data %>%
  group_by(Time, Population) %>%
  summarise(
    mean_value = mean(!!sym(dv), na.rm = TRUE),
    se = sd(!!sym(dv), na.rm = TRUE) / sqrt(n())
  )

print("Summary data for all farms:")
print(head(summary_data_all_farms))

# Loop through the unique dependent variables
for (dv in dependent_vars) {
  print(paste("Processing dependent variable for all farms:", dv))
  
  # Create the interaction plot using ggplot2 for all farms together
  p <- ggplot(summary_data_all_farms, aes(x=factor(Time, levels=c("Pre_Treatment", "Early_Lactation", "Mid_Lactation")), y=mean_value, color=Population, linetype=Population)) +
    geom_errorbar(aes(ymin=mean_value-se, ymax=mean_value+se), width=0.2) +
    geom_point() +
    geom_line() +
    labs(
      title = paste("Interaction Plot for All Farms and DV", dv),
      x = "Time",
      y = paste("Mean", dv)
    ) +
    theme_minimal() +
    theme(
      text = element_text(family = "Times New Roman", size = 12),
      plot.title = element_text(hjust = 0.5, size = 14, family = "Times New Roman"),
      legend.position = "bottom"
    )
  
  print(p)
}


## EXPORT DATAFRAMES

library(openxlsx)

# Create a new workbook
wb <- createWorkbook()

# Add sheets to the workbook for each data frame
addWorksheet(wb, "levenes_results")
writeData(wb, "levenes_results", levenes_results)

addWorksheet(wb, "normality_df")
writeData(wb, "normality_df", normality_df)

addWorksheet(wb, "results_df")
writeData(wb, "results_df", results_df)

addWorksheet(wb, "posthoc_results_df")
writeData(wb, "posthoc_results_df", posthoc_results_df)

# Save the workbook to disk
saveWorkbook(wb, "All_Data_Frames.xlsx", overwrite = TRUE)

# Create a new workbook
wb_crosstab <- createWorkbook()

# List of data frame names to be saved
crosstab_names <- c(
  "Fat.Kg_A_crosstab", "Fat.Kg_B_crosstab", "Fat.Kg_C_crosstab", "Fat.Kg_D_crosstab", "Fat.Kg_all_farms_crosstab",
  "Milk.Kg_A_crosstab", "Milk.Kg_B_crosstab", "Milk.Kg_C_crosstab", "Milk.Kg_D_crosstab","Milk.Kg_all_farms_crosstab",
  "Protein.Kg_A_crosstab", "Protein.Kg_B_crosstab", "Protein.Kg_C_crosstab", "Protein.Kg_D_crosstab","Protein.Kg_all_farms_crosstab",
  "SCC_log_A_crosstab", "SCC_log_B_crosstab", "SCC_log_C_crosstab", "SCC_log_D_crosstab", "SCC_log_all_farms_crosstab"
)

# Loop through each data frame name and add it as a new sheet
for(df_name in crosstab_names) {
  addWorksheet(wb_crosstab, df_name)
  writeData(wb_crosstab, df_name, get(df_name))
}

# Save the workbook to disk
saveWorkbook(wb_crosstab, "Crosstab_Data_Frames.xlsx", overwrite = TRUE)



## EXPORT INTERACTION PLOTS'

# Initialize an empty vector to store filenames
all_png_files <- c()

# Loop through the unique farms and dependent variables
for (dv in dependent_vars) {
  
  # Calculate the summary statistics for the current dependent variable
  summary_data <- data %>%
    group_by(Farm, Time, Population) %>%
    summarise(
      mean_value = mean(!!sym(dv), na.rm = TRUE),
      se = sd(!!sym(dv), na.rm = TRUE) / sqrt(n())
    )
  
  for (farm in unique(summary_data$Farm)) {
    
    # Subset the data for the farm of interest
    sub_data <- subset(summary_data, Farm == farm)
    
    # Create the interaction plot using ggplot2
    p <- ggplot(sub_data, aes(x=factor(Time, levels=c("Pre_Treatment", "Early_Lactation", "Mid_Lactation")), y=mean_value, color=Population, linetype=Population)) +
      geom_errorbar(aes(ymin=mean_value-se, ymax=mean_value+se), width=0.2) + # error bars
      geom_point() +  # Show mean data points
      geom_line() +   # Show mean lines
      labs(
        title = paste("Interaction Plot for Farm", farm, "and DV", dv),
        x = "Time",
        y = paste("Mean", dv)
      ) +
      theme_minimal() +
      theme(
        text = element_text(family = "Times New Roman", size = 12),
        plot.title = element_text(hjust = 0.5, size = 14, family = "Times New Roman"),
        legend.position = "bottom"
      )
    
    # Show the plot
    print(p)
    
    # Define the filename for individual farms
    farm_filename <- paste0("Interaction_Plot_Farm_", farm, "_DV_", dv, ".png")
    
    # Save the plot
    ggsave(filename = farm_filename, plot = p)
    
    # Add this filename to the vector
    all_png_files <- c(all_png_files, farm_filename)
  }
}

# Loop through the unique dependent variables
for (dv in dependent_vars) {
  
  # Calculate the summary statistics for the entire dataset (across all farms)
  summary_data_all_farms <- data %>%
    group_by(Time, Population) %>%
    summarise(
      mean_value = mean(!!sym(dv), na.rm = TRUE),
      se = sd(!!sym(dv), na.rm = TRUE) / sqrt(n())
    )
  
  # Create the interaction plot using ggplot2 for all farms together
  p <- ggplot(summary_data_all_farms, aes(x=factor(Time, levels=c("Pre_Treatment", "Early_Lactation", "Mid_Lactation")), y=mean_value, color=Population, linetype=Population)) +
    geom_errorbar(aes(ymin=mean_value-se, ymax=mean_value+se), width=0.2) + # error bars
    geom_point() +  # Show mean data points
    geom_line() +   # Show mean lines
    labs(
      title = paste("Interaction Plot for All Farms and DV", dv),
      x = "Time",
      y = paste("Mean", dv)
    ) +
    theme_minimal() +
    theme(
      text = element_text(family = "Times New Roman", size = 12),
      plot.title = element_text(hjust = 0.5, size = 14, family = "Times New Roman"),
      legend.position = "bottom"
    )
  
  # Show the plot
  print(p)
  
  # Define the filename
  all_farms_filename <- paste0("Interaction_Plot_All_Farms_", dv, ".png")
  
  # Save the plot
  ggsave(filename = all_farms_filename, plot = p)
  
  # Add this filename to the vector
  all_png_files <- c(all_png_files, all_farms_filename)
}

# Zip all the PNG files

library(zip)
zip(zipfile = "All_Interaction_Plots.zip", files = all_png_files)


### NEW DESCRIPTIVE TABLE - BETWEEN-GROUPS

# Generate a dataframe with mean, min, max, and SD scores for each farm and each population
# For individual farms
mean_scores_df <- data %>%
  group_by(Farm, Population) %>%
  summarise(across(starts_with(c("Milk", "Fat", "Protein", "SCC_log")), 
                   list(mean = mean, min = min, max = max, sd = sd), 
                   na.rm = TRUE))

mean_scores_df_wide <- mean_scores_df %>% 
  pivot_wider(names_from = Population, 
              names_glue = "{Population}_{.value}",
              values_from = starts_with(c("Milk", "Fat", "Protein", "SCC_log")))

# For "All" farms
mean_scores_all_farms <- data %>%
  group_by(Population) %>%
  summarise(across(starts_with(c("Milk", "Fat", "Protein", "SCC_log")), 
                   list(mean = mean, min = min, max = max, sd = sd), 
                   na.rm = TRUE))

mean_scores_all_farms_wide <- mean_scores_all_farms %>% 
  pivot_wider(names_from = Population, 
              names_glue = "{Population}_{.value}", 
              values_from = starts_with(c("Milk", "Fat", "Protein", "SCC_log")))

# Add a Farm column indicating this is for "All" farms
mean_scores_all_farms_wide$Farm <- "All"

# Combine with the original mean_scores_df_wide
mean_scores_df_wide <- bind_rows(mean_scores_df_wide, mean_scores_all_farms_wide)

print("Mean, min, max, and SD scores by farm and population:")
print(mean_scores_df_wide)


# Function to retrieve the ANOVA results for between-subjects and interactions.
get_anova_results <- function(farm, dv) {
  farm_results <- subset(results_df, Farm == farm & Dependent_Variable == dv)
  
  between_subjects_pvalue <- subset(farm_results, Effect == "PopulationNAP")$pValue
  between_subjects_fvalue <- subset(farm_results, Effect == "PopulationNAP")$Value
  interaction_pvalue <- subset(farm_results, Effect == "Time.Q:PopulationNAP")$pValue
  interaction_fvalue <- subset(farm_results, Effect == "Time.Q:PopulationNAP")$Value
  
  return(c(
    Between_Subjects_pValue = ifelse(length(between_subjects_pvalue) > 0, between_subjects_pvalue, NA),
    Between_Subjects_FValue = ifelse(length(between_subjects_fvalue) > 0, between_subjects_fvalue, NA),
    Interaction_pValue = ifelse(length(interaction_pvalue) > 0, interaction_pvalue, NA),
    Interaction_FValue = ifelse(length(interaction_fvalue) > 0, interaction_fvalue, NA)
  ))
}

# List of all farms including "All"
all_farms <- unique(c(results_df$Farm, "All"))

# Add the ANOVA results to mean_scores_df_wide
for (farm in all_farms) {
  for (dv in dependent_vars) {
    anova_results <- get_anova_results(farm, dv)
    for (effect_name in names(anova_results)) {
      col_name <- paste(dv, effect_name, sep = "_")
      mean_scores_df_wide[mean_scores_df_wide$Farm == farm, col_name] <- anova_results[effect_name]
    }
  }
}

# Print the updated dataframe
print(mean_scores_df_wide)

# Export to Excel

library(openxlsx)

# Write the dataframe to an Excel file
write.xlsx(mean_scores_df_wide, "Mean_Scores_By_Farm_And_Population.xlsx")

