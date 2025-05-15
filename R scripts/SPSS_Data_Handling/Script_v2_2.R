setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elodie_claire/information for rafael/datasets/english")

library(haven)
library(dplyr)

df1 <- read_sav("dataset t1.sav")
df2 <- read_sav("dataset t2.sav")
df3 <- read_sav("dataset t3.sav")

# Get rid of special characters
names(df3) <- gsub(" ", "_", trimws(names(df3)))
names(df3) <- gsub("\\s+", "_", trimws(names(df3), whitespace = "[\\h\\v\\s]+"))
names(df3) <- gsub("\\(", "_", names(df3))
names(df3) <- gsub("\\)", "_", names(df3))
names(df3) <- gsub("\\-", "_", names(df3))
names(df3) <- gsub("/", "_", names(df3))
names(df3) <- gsub("\\\\", "_", names(df3)) 
names(df3) <- gsub("\\?", "", names(df3))
names(df3) <- gsub("\\'", "", names(df3))
names(df3) <- gsub("\\,", "_", names(df3))
names(df3) <- gsub("\\$", "", names(df3))
names(df3) <- gsub("\\+", "_", names(df3))

# Checking for unique values in Q25
unique_values_df1 <- sort(unique(df1$Q25))
unique_values_df2 <- sort(unique(df2$Q25))
unique_values_df3 <- sort(unique(df3$Q25))

# Finding the maximum length for even padding
max_length <- max(length(unique_values_df1), length(unique_values_df2), length(unique_values_df3))

# Padding shorter lists with NA to align all vectors
unique_values_df1 <- c(unique_values_df1, rep(NA, max_length - length(unique_values_df1)))
unique_values_df2 <- c(unique_values_df2, rep(NA, max_length - length(unique_values_df2)))
unique_values_df3 <- c(unique_values_df3, rep(NA, max_length - length(unique_values_df3)))

# Creating a data frame for side-by-side comparison
comparison_table <- data.frame(
  Dataset_1 = unique_values_df1,
  Dataset_2 = unique_values_df2,
  Dataset_3 = unique_values_df3
)

# Print the comparison table
print(comparison_table)


# Trim all values

# Loop over each column in the dataframe
df3 <- data.frame(lapply(df3, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))

colnames(df1)
colnames(df2)
colnames(df3)


# First, reverse the scores for the specified items
df2 <- df2 %>%
  mutate(
    # Reversing negative affect measures from 1 to 7 (assuming a scale of 1 to 7)
    Affect_2 = 8 - Affect_2,
    Affect_3 = 8 - Affect_3,
    Affect_5 = 8 - Affect_5,
    Affect_8 = 8 - Affect_8,
    Affect_9 = 8 - Affect_9,
    
    # Reversing Need Satisfaction items as specified
    Q15___Need_Satisfact_4 = 8 - Q15___Need_Satisfact_4,
    Q15___Need_Satisfact_7 = 8 - Q15___Need_Satisfact_7,
    Q15___Need_Satisfact_6 = 8 - Q15___Need_Satisfact_6,
    Q15___Need_Satisfact_12 = 8 - Q15___Need_Satisfact_12,
    Q15___Need_Satisfact_2 = 8 - Q15___Need_Satisfact_2,
    Q15___Need_Satisfact_11 = 8 - Q15___Need_Satisfact_11
  )

df3 <- df3 %>%
  mutate(
    # Reversing negative affect measures from 1 to 7 (assuming a scale of 1 to 7)
    Affect_2 = 8 - Affect_2,
    Affect_3 = 8 - Affect_3,
    Affect_5 = 8 - Affect_5,
    Affect_8 = 8 - Affect_8,
    Affect_9 = 8 - Affect_9,
    
    # Reversing Need Satisfaction items as specified
    Q15___Need_Satisfact_4 = 8 - Q15___Need_Satisfact_4,
    Q15___Need_Satisfact_7 = 8 - Q15___Need_Satisfact_7,
    Q15___Need_Satisfact_6 = 8 - Q15___Need_Satisfact_6,
    Q15___Need_Satisfact_12 = 8 - Q15___Need_Satisfact_12,
    Q15___Need_Satisfact_2 = 8 - Q15___Need_Satisfact_2,
    Q15___Need_Satisfact_11 = 8 - Q15___Need_Satisfact_11
  )


# Autonomy Support at Time 1 with extended variables
autonomy_support_T1 <- c(
  "Q95___FamAutSupp2_1", "Q95___FamAutSupp2_2", "Q95___FamAutSupp2_3", "Q95___FamAutSupp2_4",
  "Q101___FrieAutSupp1_1", "Q101___FrieAutSupp1_2", "Q101___FrieAutSupp1_3", "Q101___FrieAutSupp1_4"
)

# Need Satisfaction 
need_satisfaction <- c(
  "Q15___Need_Satisfact_1", "Q15___Need_Satisfact_10", "Q15___Need_Satisfact_3", "Q15___Need_Satisfact_9",
  "Q15___Need_Satisfact_5", "Q15___Need_Satisfact_8",
  # Reverse scored
  "Q15___Need_Satisfact_4", "Q15___Need_Satisfact_7", "Q15___Need_Satisfact_6", "Q15___Need_Satisfact_12",
  "Q15___Need_Satisfact_2", "Q15___Need_Satisfact_11"
)


# Define Positive Affect, Negative Affect (reversed), and Life Satisfaction
positive_affect <- c("Affect_1", "Affect_4", "Affect_6", "Affect_7")
negative_affect_reversed <- c("Affect_2", "Affect_3", "Affect_5", "Affect_8", "Affect_9")
life_satisfaction <- c("SWLS_1", "SWLS_2", "SWLS_3", "SWLS_4", "SWLS_5")

# Subjective Well-Being at Time 3
subjective_wellbeing_T3 <- c(
  "Affect_1", "Affect_4", "Affect_6", "Affect_7", 
  # Reverse scored negative affect
  "Affect_2", "Affect_3", "Affect_5", "Affect_8", "Affect_9",
  "SWLS_1", "SWLS_2", "SWLS_3", "SWLS_4", "SWLS_5"
)

# Emotion Regulation at Time 3
emotion_regulation_T3 <- c(
  "Emotional_regulation_1", "Emotional_regulation_2", "Emotional_regulation_3",
  "Emotional_regulation_4", "Emotional_regulation_5", "Emotional_regulation_6",
  "Emotional_regulation_7", "Emotional_regulation_8", "Emotional_regulation_9",
  "Emotional_regulation_10"
)

integrative_regulation <- c("Emotional_regulation_8", "Emotional_regulation_9", "Emotional_regulation_10")
suppressive_regulation <- c("Emotional_regulation_5", "Emotional_regulation_6", "Emotional_regulation_7")
dysregulation <- c("Emotional_regulation_1", "Emotional_regulation_2", "Emotional_regulation_3", "Emotional_regulation_4")

posttraumatic_growth <- paste0("Q45_", 1:16)

# Aspirations at Time 3
aspirations_T3 <- c(
  "asp_1", "asp_2", "asp_3", "asp_4", "asp_5", "asp_6", "asp_7", "asp_8", "asp_9", 
  "asp_10", "asp_11", "asp_12"
)

intrinsic_aspirations <- c("asp_2", "asp_4", "asp_6", "asp_8", "asp_10", "asp_12")
extrinsic_aspirations <- c("asp_1", "asp_3", "asp_5", "asp_7", "asp_9", "asp_11")


# Scale Construction

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Added for item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val <- alpha_results$total$raw_alpha
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"] # Get ITC for the item
      item_mean <- mean(item_data, na.rm = TRUE)
      item_sem <- sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data)))
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = item_mean,
        SEM = item_sem,
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
        Alpha = NA
      ))
    }
    
    # Calculate the mean score for the scale and add as a new column
    scale_mean <- rowMeans(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_mean
    scale_mean_overall <- mean(scale_mean, na.rm = TRUE)
    scale_sem <- sd(scale_mean, na.rm = TRUE) / sqrt(sum(!is.na(scale_mean)))
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = scale_mean_overall,
      SEM = scale_sem,
      StDev = sd(scale_mean, na.rm = TRUE),
      ITC = NA,  # ITC is not applicable for the total scale
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}

scales1 <- list(
  "Autonomy Support T1" = c("Q95___FamAutSupp2_1", "Q95___FamAutSupp2_2", "Q95___FamAutSupp2_3", "Q95___FamAutSupp2_4",
                            "Q101___FrieAutSupp1_1", "Q101___FrieAutSupp1_2", "Q101___FrieAutSupp1_3", "Q101___FrieAutSupp1_4"))


alpha_results1 <- reliability_analysis(df1, scales1)

df1_recoded <- alpha_results1$data_with_scales
df1_descriptives <- alpha_results1$statistics


scales3 <- list(
  # Existing scales
  
  # Breakdown of Subjective Well-Being into subscales plus total
  "Positive Affect T3" = c("Affect_1", "Affect_4", "Affect_6", "Affect_7"),
  "Negative Affect T3" = c("Affect_2", "Affect_3", "Affect_5", "Affect_8", "Affect_9"),
  "Life Satisfaction T3" = c("SWLS_1", "SWLS_2", "SWLS_3", "SWLS_4", "SWLS_5"),
  "Total Subjective Well-Being T3" = c("Affect_1", "Affect_4", "Affect_6", "Affect_7",
                                    "Affect_2", "Affect_3", "Affect_5", "Affect_8", "Affect_9",
                                    "SWLS_1", "SWLS_2", "SWLS_3", "SWLS_4", "SWLS_5"),
  
  # Breakdown of Emotion Regulation into subscales
  "Integrative Regulation T3" = c("Emotional_regulation_8", "Emotional_regulation_9", "Emotional_regulation_10"),
  "Suppressive Regulation T3" = c("Emotional_regulation_5", "Emotional_regulation_6", "Emotional_regulation_7"),
  "Dysregulation T3" = c("Emotional_regulation_1", "Emotional_regulation_2", "Emotional_regulation_3", "Emotional_regulation_4"),
  "Total Emotion Regulation T3" = c("Emotional_regulation_1", "Emotional_regulation_2", "Emotional_regulation_3",
                                 "Emotional_regulation_4", "Emotional_regulation_5", "Emotional_regulation_6",
                                 "Emotional_regulation_7", "Emotional_regulation_8", "Emotional_regulation_9",
                                 "Emotional_regulation_10"),
  
  # Posttraumatic Growth
  "Posttraumatic Growth T3" = paste0("Q45_", 1:16),
  
  # Breakdown of Aspirations into subscales plus total
  "Intrinsic Aspirations T3" = c("asp_2", "asp_4", "asp_6", "asp_8", "asp_10", "asp_12"),
  "Extrinsic Aspirations T3" = c("asp_1", "asp_3", "asp_5", "asp_7", "asp_9", "asp_11"),
  "Total Aspirations T3" = c("asp_1", "asp_2", "asp_3", "asp_4", "asp_5", "asp_6", "asp_7", "asp_8", "asp_9", 
                          "asp_10", "asp_11", "asp_12")
)

alpha_results3 <- reliability_analysis(df3, scales3)

df3_recoded <- alpha_results3$data_with_scales
df3_descriptives <- alpha_results3$statistics

scales2 <- list(
  # Existing scales
  "Need Satisfaction T2" = c(
    "Q15___Need_Satisfact_1", "Q15___Need_Satisfact_2", "Q15___Need_Satisfact_3",
    "Q15___Need_Satisfact_4", "Q15___Need_Satisfact_5", "Q15___Need_Satisfact_6",
    "Q15___Need_Satisfact_7", "Q15___Need_Satisfact_8", "Q15___Need_Satisfact_9",
    "Q15___Need_Satisfact_10", "Q15___Need_Satisfact_11", "Q15___Need_Satisfact_12"
  ))

alpha_results2 <- reliability_analysis(df2, scales2)

df2_recoded <- alpha_results2$data_with_scales
df2_descriptives <- alpha_results2$statistics

scales2 <- list(
  # Existing scales
  "Need Satisfaction T2" = c(
    "Q15___Need_Satisfact_2", 
    "Q15___Need_Satisfact_4",  "Q15___Need_Satisfact_6",
    "Q15___Need_Satisfact_7",  
     "Q15___Need_Satisfact_11", "Q15___Need_Satisfact_12"
  ))

alpha_results2 <- reliability_analysis(df2, scales2)
df2_descriptives2 <- alpha_results2$statistics
df2_recoded <- alpha_results2$data_with_scales

scales3 <- list(
  # Existing scales
  
  # Breakdown of Subjective Well-Being into subscales plus total
  "Positive Affect T3" = c("Affect_1", "Affect_4", "Affect_6", "Affect_7"),
  "Negative Affect T3" = c("Affect_2", "Affect_3", "Affect_5", "Affect_8", "Affect_9"),
  "Life Satisfaction T3" = c("SWLS_1", "SWLS_2", "SWLS_3", "SWLS_4", "SWLS_5"),
  "Total Subjective Well-Being T3" = c("Affect_1", "Affect_4", "Affect_6", "Affect_7",
                                       "Affect_2", "Affect_3", "Affect_5", "Affect_8", "Affect_9",
                                       "SWLS_1", "SWLS_2", "SWLS_3", "SWLS_4", "SWLS_5"),
  
  # Breakdown of Emotion Regulation into subscales
  "Integrative Regulation T3" = c("Emotional_regulation_8", "Emotional_regulation_9", "Emotional_regulation_10"),
  "Suppressive Regulation T3" = c("Emotional_regulation_5", "Emotional_regulation_6", "Emotional_regulation_7"),
  "Dysregulation T3" = c("Emotional_regulation_1", "Emotional_regulation_2", "Emotional_regulation_3", "Emotional_regulation_4"),
  "Total Emotion Regulation T3" = c("Emotional_regulation_1", "Emotional_regulation_2", "Emotional_regulation_3",
                                    "Emotional_regulation_4", "Emotional_regulation_5", "Emotional_regulation_6",
                                    "Emotional_regulation_7", "Emotional_regulation_8", "Emotional_regulation_9",
                                    "Emotional_regulation_10"),
  
  # Posttraumatic Growth
  "Posttraumatic Growth T3" = paste0("Q45_", setdiff(1:16, c(14, 15))),
  
  # Breakdown of Aspirations into subscales plus total
  "Intrinsic Aspirations T3" = c("asp_2", "asp_4", "asp_6", "asp_8", "asp_10", "asp_12"),
  "Extrinsic Aspirations T3" = c("asp_1", "asp_3", "asp_5", "asp_7", "asp_9", "asp_11"),
  "Total Aspirations T3" = c("asp_1", "asp_2", "asp_3", "asp_4", "asp_5", "asp_6", "asp_7", "asp_8", "asp_9", 
                             "asp_10", "asp_11", "asp_12")
)

alpha_results3 <- reliability_analysis(df3, scales3)

df3_recoded <- alpha_results3$data_with_scales
df3_descriptives2 <- alpha_results3$statistics


colnames(df1_recoded)
colnames(df2_recoded)
colnames(df3_recoded)

# Join datasets by Q25

library(dplyr)

# Define a function to extract all variable names from the scales list
get_vars_from_scales <- function(scales_list) {
  unlist(scales_list, recursive = FALSE)
}

# Prepare df1_recoded
df1_recoded <- df1_recoded %>%
  mutate(Q25_normalized = tolower(Q25)) %>%
  filter(Q25_normalized != "") %>%
  select(Q25, Q25_normalized, get_vars_from_scales(scales1), "Autonomy Support T1", 
         c("Q64_3", "Q64_9", "Q64_4", "Q64_5", "Q64_6", "Q64_2", "Q64_1", "Q64_8", "Q18", "Q4", "Q6", "Q48.0", "Q47", "Q81", "Q60"))

# Prepare df2_recoded
df2_recoded <- df2_recoded %>%
  mutate(Q25_normalized = tolower(Q25)) %>%
  filter(Q25_normalized != "") %>%
  select(c("Q25", Q25_normalized), get_vars_from_scales(scales2), "Need Satisfaction T2")

# Prepare df3_recoded
df3_recoded <- df3_recoded %>%
  mutate(Q25_normalized = tolower(Q25)) %>%
  filter(Q25_normalized != "") %>%
  select(c("Q25", Q25_normalized), get_vars_from_scales(scales3), 
         "Positive Affect T3", "Negative Affect T3", "Life Satisfaction T3",
         "Total Subjective Well-Being T3", "Integrative Regulation T3", "Suppressive Regulation T3",
         "Dysregulation T3", "Total Emotion Regulation T3", "Posttraumatic Growth T3",
         "Intrinsic Aspirations T3", "Extrinsic Aspirations T3", "Total Aspirations T3")

# Join df1_recoded and df2_recoded
df1_2_joined <- df1_recoded %>%
  left_join(df2_recoded, by = "Q25_normalized")

# Join the result with df3_recoded
df_all_joined <- df1_2_joined %>%
  left_join(df3_recoded, by = "Q25_normalized")


# Replace NA values with 0 in columns starting with "Q64"
df_all_joined <- df_all_joined %>%
  mutate(across(starts_with("Q64"), ~replace_na(., 0)))

# Filter only complete cases

# Check for non-missing values across Q25.x, Q25.y, and Q25 columns
complete_cases <- df_all_joined %>%
  filter(!is.na(Q25.x) & !is.na(Q25.y) & !is.na(Q25))

# View the structure or a summary of the dataframe with complete cases
str(complete_cases)
summary(complete_cases)

# Optionally, you might want to keep only the rows where Q25.x, Q25.y, and Q25 are non-missing
df_all_joined_comp <- df_all_joined %>%
  filter(!is.na(Q25.x) & !is.na(Q25.y) & !is.na(Q25))

library(tidyr)
# Function to diagnose missing data
diagnose_missing_data <- function(df) {
  vars <- colnames(df)  # Automatically take all column names from the dataframe
  missing_summary <- df %>%
    select(all_of(vars)) %>%
    summarise(across(everything(), ~ sum(is.na(.)))) %>%
    pivot_longer(cols = everything(), names_to = "Variable", values_to = "Missing_Values")
  
  total_rows <- nrow(df)
  
  missing_summary <- missing_summary %>%
    mutate(Percentage_Missing = (Missing_Values / total_rows) * 100)
  
  return(missing_summary)
}

# Run missing data diagnosis on all variables of the complete joined data
missing_data_summary <- diagnose_missing_data(df_all_joined_comp)

colnames(df_all_joined_comp)

vars_relevant <- c("Autonomy Support T1", "Need Satisfaction T2", "Positive Affect T3"       ,        "Negative Affect T3"  , 
                   "Life Satisfaction T3"  ,           "Total Subjective Well-Being T3"  , "Integrative Regulation T3" ,      "Suppressive Regulation T3" ,      
                   "Dysregulation T3"      ,           "Total Emotion Regulation T3"   ,   "Posttraumatic Growth T3"  ,        "Intrinsic Aspirations T3" ,       
                   "Extrinsic Aspirations T3"   ,      "Total Aspirations T3")

library(e1071) 

calculate_stats <- function(data, variables) {
  results <- data.frame(Variable = character(),
                        Skewness = numeric(),
                        Kurtosis = numeric(),
                        Shapiro_Wilk_F = numeric(),
                        Shapiro_Wilk_p_value = numeric(),
                        KS_Statistic = numeric(),
                        KS_p_value = numeric(),
                        stringsAsFactors = FALSE)
  
  for (var in variables) {
    if (var %in% names(data)) {
      # Calculate skewness and kurtosis
      skew <- skewness(data[[var]], na.rm = TRUE)
      kurt <- kurtosis(data[[var]], na.rm = TRUE)
      
      # Perform Shapiro-Wilk test
      shapiro_test <- shapiro.test(data[[var]])
      
      # Perform Kolmogorov-Smirnov test
      ks_test <- ks.test(data[[var]], "pnorm", mean = mean(data[[var]], na.rm = TRUE), sd = sd(data[[var]], na.rm = TRUE))
      
      # Add results to the dataframe
      results <- rbind(results, c(var, skew, kurt, shapiro_test$statistic, shapiro_test$p.value, ks_test$statistic, ks_test$p.value))
    } else {
      warning(paste("Variable", var, "not found in the data. Skipping."))
    }
  }
  
  colnames(results) <- c("Variable", "Skewness", "Kurtosis", "Shapiro_Wilk_F", "Shapiro_Wilk_p_value", "KS_Statistic", "KS_p_value")
  return(results)
}

# Example usage with your data
df_normality_results <- calculate_stats(df_all_joined_comp, vars_relevant)

## DESCRIPTIVE TABLES

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

df_descriptive_stats <- calculate_descriptive_stats(df_all_joined_comp, vars_relevant)


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


correlation_matrix <- calculate_correlation_matrix(df_all_joined_comp, vars_relevant, method = "pearson")

## BOXPLOTS

library(ggplot2)
library(tidyr)
library(stringr)

create_boxplots <- function(df, vars) {
  # Ensure that vars are in the dataframe
  df <- df[, vars, drop = FALSE]
  
  # Reshape the data to a long format
  long_df <- df %>%
    gather(key = "Variable", value = "Value")
  
  # Create side-by-side boxplots for each variable
  p <- ggplot(long_df, aes(x = Variable, y = Value)) +
    geom_boxplot() +
    stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = "Side-by-Side Boxplots", x = "Variable", y = "Value") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1)) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 10))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = "side_by_side_boxplots.png", plot = p, width = 10, height = 6)
}

create_boxplots(df_all_joined_comp, vars_relevant)

colnames(df_all_joined_comp)


# SEM

# Rename variables by replacing spaces with underscores
df_all_joined_comp <- df_all_joined_comp %>%
  rename_with(~ gsub(" ", "_", .), everything())

library(lavaan)
library(dplyr)
library(tidyr)

# Define the SEM model for mediation using the new variable names with underscores
model <- '
  # Mediation paths
  Positive_Affect_T3 ~ b1*Need_Satisfaction_T2 + c1*Autonomy_Support_T1
  Negative_Affect_T3 ~ b2*Need_Satisfaction_T2 + c2*Autonomy_Support_T1
  Life_Satisfaction_T3 ~ b3*Need_Satisfaction_T2 + c3*Autonomy_Support_T1
  Integrative_Regulation_T3 ~ b4*Need_Satisfaction_T2 + c4*Autonomy_Support_T1
  Suppressive_Regulation_T3 ~ b5*Need_Satisfaction_T2 + c5*Autonomy_Support_T1
  Dysregulation_T3 ~ b6*Need_Satisfaction_T2 + c6*Autonomy_Support_T1
  Posttraumatic_Growth_T3 ~ b7*Need_Satisfaction_T2 + c7*Autonomy_Support_T1
  Intrinsic_Aspirations_T3 ~ b8*Need_Satisfaction_T2 + c8*Autonomy_Support_T1
  Extrinsic_Aspirations_T3 ~ b9*Need_Satisfaction_T2 + c9*Autonomy_Support_T1

  # Mediator path
  Need_Satisfaction_T2 ~ a*Autonomy_Support_T1

  # Indirect effects
  indirect1 := a*b1
  indirect2 := a*b2
  indirect3 := a*b3
  indirect4 := a*b4
  indirect5 := a*b5
  indirect6 := a*b6
  indirect7 := a*b7
  indirect8 := a*b8
  indirect9 := a*b9
'

# Fit the model using listwise deletion for handling missing data
fit <- sem(model, data = df_all_joined_comp, missing = "listwise")

# Summarize the results including standardized estimates and fit measures
results <- summary(fit, standardized = TRUE, fit.measures = TRUE)

# Extracting coefficients and formatting results
coefficients <- parameterEstimates(fit, standardized = TRUE, se = TRUE, ci = TRUE, header = TRUE, rsquare = TRUE)



# Demographics

# Function to extract value labels from each variable
extract_value_labels <- function(var) {
  labels <- attr(var, "labels")
  if (!is.null(labels)) {
    paste(names(labels), labels, sep = " - ", collapse = ", ")
  } else {
    NA
  }
}

# Prepare a data frame to capture variable names and value labels
variable_info <- tibble(
  Variable = names(df_all_joined_comp),
  ValueLabels = sapply(df_all_joined_comp, extract_value_labels)
)

# View the variable information with value labels
print(variable_info)


library(dplyr)

# Define the specific list of variables
vars <- c("Q64_3", "Q64_9", "Q64_4", "Q64_5", "Q64_6", "Q64_2", "Q64_1", "Q64_8",
           "Q4", "Q6", "Q48.0", "Q47", "Q81", "Q60")

# Function to create frequency tables for given variables
create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each category variable
  for (category in categories) {
    # Ensure the category variable is a factor
    data[[category]] <- factor(data[[category]])
    
    # Calculate counts
    counts <- table(data[[category]])
    
    # Create a dataframe for this category
    freq_table <- data.frame(
      "Category" = rep(category, length(counts)),
      "Level" = names(counts),
      "Count" = as.integer(counts),
      stringsAsFactors = FALSE
    )
    
    # Calculate and add percentages
    freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
    
    # Add the result to the list
    all_freq_tables[[category]] <- freq_table
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

# Apply the function to the dataset and the specified variables
df_freq <- create_frequency_tables(df_all_joined_comp, vars)



# Define label mappings as provided
labels <- list(
  Q64_3 = c("1" = "Black"),
  Q64_9 = c("1" = "White"),
  Q64_4 = c("1" = "Hispanic or Latinx"),
  Q64_5 = c("1" = "Asian (including South Asian)"),
  Q64_6 = c("1" = "Middle Eastern or North African"),
  Q64_2 = c("1" = "Indigenous from outside of Canada"),
  Q64_1 = c("1" = "First Nations, Métis, or Inuit"),
  Q64_8 = c("1" = "Other, please specify:"),
  Q4 = c("1" = "Employed full-time", "2" = "Employed part-time", "3" = "Unemployed", 
         "4" = "Full-time student", "5" = "Retired", "6" = "Full-time homemaker or carer", "7" = "Part-time student"),
  Q6 = c("1" = "Male", "2" = "Female", "3" = "Transgender", "4" = "If self-definition is preferred, fill in the blank",
         "5" = "Genderqueer", "6" = "Non-Binary"),
  Q48_0 = c("12" = "Less than $20,000", "13" = "$20,000 to $40,000", "14" = "$40,000 to $60,000",
            "15" = "$60,000 to $80,000", "17" = "$80,000 and over"),
  Q47 = c("1" = "No certificate, diploma or degree", "2" = "Secondary (high) school diploma or equivalency certificate",
          "3" = "Apprenticeship or trades certificate or diploma", "4" = "College, CEGEP or other non-university certificate or diploma",
          "5" = "Bachelor's degree", "6" = "Master's degree", "7" = "Earned doctorate"),
  Q81 = c("14" = "Canadian citizen", "15" = "Permanent resident", "16" = "Temporary resident (visitor, student, worker or other)",
          "17" = "Asylum-seeker or refugee", "18" = "Other, please specify:"),
  Q60 = c("1" = "Born in Quebec", "2" = "Born in Canada, but outside of Quebec", 
          "3" = "Born outside of Canada")
)

# Apply labels to the Level column based on Category and existing Level
df_freq$LevelLabel <- mapply(function(category, level) {
  if (!is.null(labels[[category]])) {
    labels[[category]][level]
  } else {
    level  # Return the original level if no label is found
  }
}, df_freq$Category, df_freq$Level)

# Print the updated dataframe
print(df_freq)


#Merge categories

# Define numeric merges as a list where each entry corresponds to a variable
merge_maps <- list(
  Q4 = c("3" = "19", "5" = "19", "6" = "19"),      # Merge levels 3, 5, and 6 into 19 in Q4
  Q48.0 = c("12" = "19", "13" = "19"),             # Merge levels 12 and 13 into 19 in Q48.0
  Q47 = c("6" = "20", "7" = "20", "1" = "19", "2" = "19", "3" = "19"),  # Merge levels 6, 7, 1, 2, and 3 into 19 in Q47
  Q81 = c("15" = "19", "16" = "19", "17" = "19", "18" = "19")   # Merge levels 15, 16, 17, and 18 into 19 in Q81
)


# Update the dataframe by applying the merge maps more directly
df_all_joined_comp_merged <- df_all_joined_comp %>%
  mutate(
    Q4 = as.character(Q4),
    Q48.0 = as.character(Q48.0),
    Q47 = as.character(Q47),
    Q81 = as.character(Q81)
  ) %>%
  mutate(
    Q4 = replace(Q4, Q4 %in% names(merge_maps$Q4), merge_maps$Q4[Q4[Q4 %in% names(merge_maps$Q4)]]),
    Q48.0 = replace(Q48.0, Q48.0 %in% names(merge_maps$Q48.0), merge_maps$Q48.0[Q48.0[Q48.0 %in% names(merge_maps$Q48.0)]]),
    Q47 = replace(Q47, Q47 %in% names(merge_maps$Q47), merge_maps$Q47[Q47[Q47 %in% names(merge_maps$Q47)]]),
    Q81 = replace(Q81, Q81 %in% names(merge_maps$Q81), merge_maps$Q81[Q81[Q81 %in% names(merge_maps$Q81)]])
  )


# Apply the function to the dataset and the specified variables
df_freq_v2 <- create_frequency_tables(df_all_joined_comp_merged, vars)

# Dummyfy columns

library(fastDummies)

vars <- c("Q4", "Q6","Q48.0","Q60", "Q47", "Q81")

df_all_joined_comp_merged_dum <- dummy_cols(df_all_joined_comp_merged, select_columns = vars)

colnames(df_all_joined_comp_merged_dum)

# Controlled MOdel

# SEM

# Rename variables by replacing spaces with underscores
df_all_joined_comp <- df_all_joined_comp %>%
  rename_with(~ gsub(" ", "_", .), everything())

library(lavaan)
library(dplyr)
library(tidyr)

# Define the SEM model for mediation using the new variable names with underscores
model2 <- '
  # Mediation paths
  Positive_Affect_T3 ~ b1*Need_Satisfaction_T2 + c1*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Negative_Affect_T3 ~ b2*Need_Satisfaction_T2 + c2*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Life_Satisfaction_T3 ~ b3*Need_Satisfaction_T2 + c3*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Integrative_Regulation_T3 ~ b4*Need_Satisfaction_T2 + c4*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Suppressive_Regulation_T3 ~ b5*Need_Satisfaction_T2 + c5*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Dysregulation_T3 ~ b6*Need_Satisfaction_T2 + c6*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Posttraumatic_Growth_T3 ~ b7*Need_Satisfaction_T2 + c7*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Intrinsic_Aspirations_T3 ~ b8*Need_Satisfaction_T2 + c8*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3
  Extrinsic_Aspirations_T3 ~ b9*Need_Satisfaction_T2 + c9*Autonomy_Support_T1 + ctrl1*Q64_9 + ctrl2*Q64_5 + ctrl3*Q4_2 + ctrl4*Q4_4 + ctrl5*Q4_19 + ctrl6*Q6_2 + ctrl7*Q48.0_15 + ctrl8*Q48.0_17 + ctrl9*Q48.0_19 + ctrl10*Q47_4 + ctrl11*Q47_5 + ctrl12*Q81_19 + ctrl13*Q60_2 + ctrl14*Q60_3

  # Mediator path
  Need_Satisfaction_T2 ~ a*Autonomy_Support_T1

  # Indirect effects
  indirect1 := a*b1
  indirect2 := a*b2
  indirect3 := a*b3
  indirect4 := a*b4
  indirect5 := a*b5
  indirect6 := a*b6
  indirect7 := a*b7
  indirect8 := a*b8
  indirect9 := a*b9
'

colnames(df_all_joined_comp_merged_dum) <- gsub(" ", "_", colnames(df_all_joined_comp_merged_dum))

# Fit the model using listwise deletion for handling missing data
fit2 <- sem(model2, data = df_all_joined_comp_merged_dum, missing = "listwise")

# Summarize the results including standardized estimates and fit measures
results2 <- summary(fit2, standardized = TRUE, fit.measures = TRUE)

# Extracting coefficients and formatting results
coefficients2 <- parameterEstimates(fit2, standardized = TRUE, se = TRUE, ci = TRUE, header = TRUE, rsquare = TRUE)



#New Freq table

# Define new labels for merged categories
new_labels <- list(
  Q4 = c("19" = "Unemployed/Homemaker/Retired"),
  Q48_0 = c("19" = "Less than $40,000"),
  Q47 = c("19" = "No certificate/Diploma/Degree or Secondary or Apprenticeship"),
  Q81 = c("19" = "All residents (Including Temporary, Permanent, etc.)")
)



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
  "Frequency Table 1" = df_freq,
  "Frequency Table 2" = df_freq_v2,
  "Controlled Model" = coefficients2
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_v1.xlsx")


# Example usage
data_list <- list(
  
  
  "Reliability2" = df2_descriptives2,
  "Reliability3" = df3_descriptives2
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "NewReliabilities.xlsx")


