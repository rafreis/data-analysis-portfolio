library(e1071) 

# Load any required libraries if not already done
# library(e1071)  # For skewness and kurtosis functions

calculate_stats <- function(data, variables) {
  # Initialize the results dataframe
  results <- data.frame(
    Variable = character(),
    Skewness = numeric(),
    Kurtosis = numeric(),
    Shapiro_Wilk_F = numeric(),
    Shapiro_Wilk_p_value = numeric(),
    N = numeric(),  # New column to store sample size
    stringsAsFactors = FALSE
  )
  
  for (var in variables) {
    if (var %in% names(data)) {
      # Count of non-missing observations
      n_val <- sum(!is.na(data[[var]]))
      
      # Calculate skewness and kurtosis
      skew <- skewness(data[[var]], na.rm = TRUE)
      kurt <- kurtosis(data[[var]], na.rm = TRUE)
      
      # Perform Shapiro-Wilk test, if sufficient data points
      if (n_val >= 3) {  # Shapiro-Wilk requires at least 3 data points
        shapiro_test <- shapiro.test(na.omit(data[[var]]))
        shapiro_stat <- shapiro_test$statistic
        shapiro_p_value <- shapiro_test$p.value
      } else {
        shapiro_stat = NA
        shapiro_p_value = NA
      }
      
      # Add results to the dataframe
      results <- rbind(results, data.frame(
        Variable = var,
        Skewness = skew,
        Kurtosis = kurt,
        Shapiro_Wilk_F = shapiro_stat,
        Shapiro_Wilk_p_value = shapiro_p_value,
        N = n_val  # Include the sample size
      ))
    } else {
      warning(paste("Variable", var, "not found in the data. Skipping."))
    }
  }
  
  return(results)
}

df_normality_results_dipa <- calculate_stats(df_dipa, scales_dipa)
df_normality_results_diga <- calculate_stats(df_diga, scales_diga)


# Define normal and non-normal scales for Dipa and Diga

scales_dipa_normal <- c("six_min_walk_test_distance_in_meter", "development_1"     ,                 
                        "development_3"   ,                  
                        "development_4"       ,                "development_5"  )

scales_dipa_nonnormal <- c("PCS_12"  ,                          
                           "MCS_12"   ,                           "HCS_MEAN"   ,                        
                           "HLS_SUM"           ,                                    
                           "development_2"  ,
                           "Blood_pressure___at_rest__systolic"  ,                                                                                                                                                                                                                                                          
                           "Blood_pressure___at_rest__diastolic_"                    )



scales_diga_normal <- c(     
                           "development_4" )

scales_diga_nonnormal <- c("six_min_walk_test_distance_in_meter" ,"HBP_MEAN"     ,                      
                         "PCS_12"         ,                     "MCS_12"              ,               
                         "HCS_MEAN"          ,                  "DIGA_HELM_Evaluation_Overall" ,      
                         "HLS_SUM"          ,                   "development_1"          ,            
                         "development_2"      ,                 "development_3"        ,              
                         "development_5" ,
                         "Blood_pressure___at_rest__systolic"  ,                                                                                                                                                                                                                                                          
                         "Blood_pressure___at_rest__diastolic_")

# Compare baseline values

run_unpaired_t_tests <- function(data, scale_vars, group_col, baseline_col, baseline_value) {
  if (!all(c(group_col, baseline_col) %in% names(data))) {
    stop("Specified columns do not exist in the dataframe.")
  }
  
  data_baseline <- data[data[[baseline_col]] == baseline_value, ]
  
  group_levels <- levels(factor(data_baseline[[group_col]]))
  if (length(group_levels) != 2) {
    stop("Group column must have exactly two levels.")
  }
  
  results <- data.frame(Variable = character(),
                        Mean_1 = numeric(),
                        Mean_2 = numeric(),
                        T_Value = numeric(),
                        P_Value = numeric(),
                        N_1 = integer(),  # New column for sample size of the first group
                        N_2 = integer(),  # New column for sample size of the second group
                        stringsAsFactors = FALSE)
  
  for (var in scale_vars) {
    t_test_result <- t.test(data_baseline[[var]] ~ data_baseline[[group_col]])
    
    results <- rbind(results, data.frame(
      Variable = var,
      Mean_1 = mean(data_baseline[data_baseline[[group_col]] == group_levels[1], ][[var]], na.rm = TRUE),
      Mean_2 = mean(data_baseline[data_baseline[[group_col]] == group_levels[2], ][[var]], na.rm = TRUE),
      T_Value = t_test_result$statistic,
      P_Value = t_test_result$p.value,
      N_1 = sum(!is.na(data_baseline[data_baseline[[group_col]] == group_levels[1], ][[var]])),  # Count non-missing values for group 1
      N_2 = sum(!is.na(data_baseline[data_baseline[[group_col]] == group_levels[2], ][[var]]))   # Count non-missing values for group 2
    ))
  }
  
  names(results)[2:3] <- paste("Mean", group_levels, sep = "_")
  return(results)
}


df_ttest_baseline_dipa <- run_unpaired_t_tests(df_dipa, scales_dipa_normal, "Classification", "Pre_Post", "Pre")
df_ttest_baseline_diga <- run_unpaired_t_tests(df_diga, scales_diga_normal, "Classification", "Pre_Post", "Pre")

run_mann_whitney_tests <- function(data, scale_vars, group_col, baseline_col, baseline_value) {
  data_baseline <- data[data[[baseline_col]] == baseline_value, ]
  
  group_levels <- levels(factor(data_baseline[[group_col]]))
  if (length(group_levels) != 2) {
    stop("Group column must have exactly two levels.")
  }
  
  results <- data.frame(
    Variable = character(),
    Mean_1 = numeric(),
    Mean_2 = numeric(),
    U_Value = numeric(),
    P_Value = numeric(),
    N_1 = integer(),
    N_2 = integer(),
    stringsAsFactors = FALSE
  )
  
  for (var in scale_vars) {
    # Ensure the data is numeric by coercing
    group_1_data <- as.numeric(data_baseline[data_baseline[[group_col]] == group_levels[1], var, drop = TRUE])
    group_2_data <- as.numeric(data_baseline[data_baseline[[group_col]] == group_levels[2], var, drop = TRUE])
    
    mw_test_result <- wilcox.test(x = group_1_data, y = group_2_data, exact = FALSE, correct = TRUE)
    
    n_1 <- sum(!is.na(group_1_data))  # Count non-missing values for group 1
    n_2 <- sum(!is.na(group_2_data))  # Count non-missing values for group 2
    
    results <- rbind(results, data.frame(
      Variable = var,
      Mean_1 = mean(group_1_data, na.rm = TRUE),
      Mean_2 = mean(group_2_data, na.rm = TRUE),
      U_Value = mw_test_result$statistic,
      P_Value = mw_test_result$p.value,
      N_1 = n_1,
      N_2 = n_2
    ))
  }
  
  names(results)[2:3] <- paste("Mean", group_levels, sep = "_")
  return(results)
}


df_mannwhitney_baseline_dipa <- run_mann_whitney_tests(df_dipa, scales_dipa_nonnormal, "Classification", "Pre_Post", "Pre")
df_mannwhitney_baseline_diga <- run_mann_whitney_tests(df_diga, scales_diga_nonnormal, "Classification", "Pre_Post", "Pre")

# Run Paired tests for both groups

# Subset data for Control group
df_control_diga <- df_diga %>% filter(Classification == "Control")
df_control_dipa <- df_dipa %>% filter(Classification == "Control")

# Subset data for Treatment group
df_treatment_diga <- df_diga %>% filter(Classification == "Treatment")
df_treatment_dipa <- df_dipa %>% filter(Classification == "Treatment")

library(dplyr)
library(tidyr)

run_paired_t_tests <- function(data, subject_col, timepoint_col, first_timepoint, second_timepoint, measurement_cols) {
  results <- data.frame(Variable = character(),
                        Mean_First = numeric(),
                        Mean_Second = numeric(),
                        T_Value = numeric(),
                        P_Value = numeric(),
                        Effect_Size = numeric(),
                        Num_Pairs = integer(),  # Already has N for pairs
                        stringsAsFactors = FALSE)
  
  for (var in measurement_cols) {
    first_data <- data[data[[timepoint_col]] == first_timepoint & !is.na(data[[var]]), ]
    second_data <- data[data[[timepoint_col]] == second_timepoint & !is.na(data[[var]]), ]
    
    paired_data <- merge(first_data, second_data, by = subject_col, suffixes = c("_first", "_second"))
    num_pairs <- nrow(paired_data)
    
    if (num_pairs > 0) {
      t_test_result <- t.test(paired_data[[paste0(var, "_first")]], paired_data[[paste0(var, "_second")]], paired = TRUE)
      # Calculate mean difference
      mean_diff <- mean(paired_data[[paste0(var, "_second")]] - paired_data[[paste0(var, "_first")]], na.rm = TRUE)
      
      # Calculate standard deviation of differences
      sd_diff <- sd(paired_data[[paste0(var, "_second")]] - paired_data[[paste0(var, "_first")]], na.rm = TRUE)
      
      # Calculate Cohen's d
      effect_size <- mean_diff / sd_diff
      
      results <- rbind(results, data.frame(
        Variable = var,
        Mean_First = mean(paired_data[[paste0(var, "_first")]], na.rm = TRUE),
        Mean_Second = mean(paired_data[[paste0(var, "_second")]], na.rm = TRUE),
        T_Value = t_test_result$statistic,
        P_Value = t_test_result$p.value,
        Effect_Size = effect_size,
        Num_Pairs = num_pairs  # Already has the number of pairs
      ))
    }
  }
  
  return(results)
}


df_paired_t_tests_control_diga <- run_paired_t_tests(df_control_diga, "Patient", "Pre_Post", "Pre", "Post", scales_diga_normal)
df_paired_t_tests_treatment_diga <- run_paired_t_tests(df_treatment_diga, "Patient", "Pre_Post", "Pre", "Post", scales_diga_normal)
df_paired_t_tests_control_dipa <- run_paired_t_tests(df_control_dipa, "Patient", "Pre_Post", "Pre", "Post", scales_dipa_normal)
df_paired_t_tests_treatment_dipa <- run_paired_t_tests(df_treatment_dipa, "Patient", "Pre_Post", "Pre", "Post", scales_dipa_normal)


# Wilcoxon Signed-Rank Test with rank biserial correlation effect size

run_paired_wilcox_tests <- function(data, subject_col, timepoint_col, first_timepoint, second_timepoint, measurement_cols) {
  results_df <- data.frame(Measurement = character(),
                           Mean_First = numeric(),
                           Mean_Second = numeric(),
                           W_Statistic = numeric(),
                           P_Value = numeric(),
                           Effect_Size = numeric(),
                           Num_Pairs = integer(),  # Already has N for pairs
                           stringsAsFactors = FALSE)
  
  for (measurement_col in measurement_cols) {
    first_data <- data[data[[timepoint_col]] == first_timepoint & !is.na(data[[measurement_col]]), ]
    second_data <- data[data[[timepoint_col]] == second_timepoint & !is.na(data[[measurement_col]]), ]
    
    paired_data <- merge(first_data, second_data, by = subject_col, suffixes = c("_first", "_second"))
    num_pairs <- nrow(paired_data)
    
    if (num_pairs > 0) {
      wilcox_result <- wilcox.test(paired_data[[paste0(measurement_col, "_first")]], 
                                   paired_data[[paste0(measurement_col, "_second")]], 
                                   paired = TRUE)
      
      effect_size <- wilcox_result$statistic / (num_pairs * (num_pairs + 1) / 2)
      
      results_df <- rbind(results_df, data.frame(
        Measurement = measurement_col,
        Mean_First = mean(paired_data[[paste0(measurement_col, "_first")]], na.rm = TRUE),
        Mean_Second = mean(paired_data[[paste0(measurement_col, "_second")]], na.rm = TRUE),
        W_Statistic = wilcox_result$statistic,
        P_Value = wilcox_result$p.value,
        Effect_Size = effect_size,
        Num_Pairs = num_pairs  # Already has the number of pairs
      ))
    }
  }
  
  return(results_df)
}


df_wilcoxon_control_diga <- run_paired_wilcox_tests(df_control_diga,"Patient", "Pre_Post", "Pre", "Post", scales_diga_nonnormal)
df_wilcoxon_treatment_diga <- run_paired_wilcox_tests(df_treatment_diga,"Patient", "Pre_Post", "Pre", "Post", scales_diga_nonnormal)
df_wilcoxon_control_dipa <- run_paired_wilcox_tests(df_control_dipa,"Patient", "Pre_Post", "Pre", "Post", scales_dipa_nonnormal)
df_wilcoxon_treatment_dipa <- run_paired_wilcox_tests(df_treatment_dipa,"Patient", "Pre_Post", "Pre", "Post", scales_dipa_nonnormal)


# Run Unpaired Tests on differences

calculate_differences <- function(data, variable_names, id_column, classification_column, timepoint_column, pre_identifier, post_identifier) {
  # Ensure necessary libraries are loaded
  if (!require("dplyr")) install.packages("dplyr")
  library(dplyr)
  
  # Reshape the data into a wide format
  data_wide <- data %>%
    tidyr::pivot_wider(names_from = !!rlang::sym(timepoint_column), 
                       values_from = variable_names, 
                       id_cols = c(!!rlang::sym(id_column), !!rlang::sym(classification_column)))
  
  # Calculate differences for each variable
  for (var in variable_names) {
    pre_col_name <- paste0(var, "_", pre_identifier)
    post_col_name <- paste0(var, "_", post_identifier)
    difference_col_name <- paste0(var, "_difference")
    
    data_wide <- data_wide %>%
      mutate(!!rlang::sym(difference_col_name) := !!rlang::sym(post_col_name) - !!rlang::sym(pre_col_name))
  }
  
  return(data_wide)
}


df_diga_differences <- calculate_differences(df_diga, scales_diga, "Patient","Classification", "Pre_Post", "Pre", "Post")
df_dipa_differences <- calculate_differences(df_dipa, scales_dipa, "Patient", "Classification", "Pre_Post", "Pre", "Post")

colnames(df_dipa_differences)

scales_diga_diff <- c("six_min_walk_test_distance_in_meter_difference",
 "HBP_MEAN_difference"      ,                      "PCS_12_difference"   ,                          
 "MCS_12_difference"         ,                     "HCS_MEAN_difference"  ,                         
 "DIGA_HELM_Evaluation_Overall_difference"  ,      "HLS_SUM_difference"  ,                          
 "development_1_difference"                 ,      "development_2_difference"       ,               
 "development_3_difference"                 ,      "development_4_difference"        ,              
 "development_5_difference", 
 "Blood_pressure___at_rest__systolic_difference"  , "Blood_pressure___at_rest__diastolic__difference")
 
scales_dipa_diff <- c("six_min_walk_test_distance_in_meter_difference",
                                            "PCS_12_difference"   ,                          
                      "MCS_12_difference"         ,                     "HCS_MEAN_difference"  ,                         
                            "HLS_SUM_difference"  ,                          
                      "development_1_difference"                 ,      "development_2_difference"       ,               
                      "development_3_difference"                 ,      "development_4_difference"        ,              
                      "development_5_difference", 
                      "Blood_pressure___at_rest__systolic_difference"  , "Blood_pressure___at_rest__diastolic__difference")

# Calculate normality of difference variables

df_normality_results_diff_dipa <- calculate_stats(df_dipa_differences, scales_dipa_diff)
df_normality_results_diff_diga <- calculate_stats(df_diga_differences, scales_diga_diff)


scales_diga_diff_normal <- c("PCS_12_difference"   ,                          
                                     
                      "development_3_difference"                 ,      "development_4_difference"                      
                      )

scales_diga_diff_nonnormal <- c("six_min_walk_test_distance_in_meter_difference",
                             "HBP_MEAN_difference"      ,                                               
                             "MCS_12_difference"         ,                     "HCS_MEAN_difference"  ,                         
                             "DIGA_HELM_Evaluation_Overall_difference"  ,      "HLS_SUM_difference"  ,                          
                             "development_1_difference"                 ,      "development_2_difference"       ,               
                                           
                             "development_5_difference",
                             
                             "Blood_pressure___at_rest__systolic_difference"  , "Blood_pressure___at_rest__diastolic__difference")

scales_dipa_diff_normal <- c(               
                      "development_3_difference"                 ,      "development_4_difference"        ,              
                      "development_5_difference")

scales_dipa_diff_nonnormal <- c("six_min_walk_test_distance_in_meter_difference",
                      "PCS_12_difference"   ,                          
                      "MCS_12_difference"         ,                     "HCS_MEAN_difference"  ,                         
                      "HLS_SUM_difference"  ,                          
                      "development_1_difference"                 ,      "development_2_difference" 
                      , 
                      "Blood_pressure___at_rest__systolic_difference"  , "Blood_pressure___at_rest__diastolic__difference"
                      )

# T-tests on the differences

run_unpaired_t_tests <- function(data, scale_vars, group_col) {
  if (!all(c(group_col, scale_vars) %in% names(data))) {
    stop("Specified columns do not exist in the dataframe.")
  }
  
  group_levels <- unique(data[[group_col]])
  if (length(group_levels) != 2) {
    stop("Group column must have exactly two levels.")
  }
  
  results <- data.frame(
    Variable = rep(scale_vars, each = 1),
    Mean_1 = numeric(length(scale_vars)),
    Mean_2 = numeric(length(scale_vars)),
    T_Value = numeric(length(scale_vars)),
    P_Value = numeric(length(scale_vars)),
    N_1 = integer(length(scale_vars)),  # New column for the first group
    N_2 = integer(length(scale_vars)),  # New column for the second group
    stringsAsFactors = FALSE
  )
  
  for (i in seq_along(scale_vars)) {
    var <- scale_vars[i]
    t_test_result <- t.test(data[[var]] ~ data[[group_col]])
    
    results$Mean_1[i] <- mean(data[data[[group_col]] == group_levels[1], var, drop = TRUE], na.rm = TRUE)
    results$Mean_2[i] <- mean(data[data[[group_col]] == group_levels[2], var, drop = TRUE], na.rm = TRUE)
    results$T_Value[i] <- t_test_result$statistic
    results$P_Value[i] <- t_test_result$p.value
    results$N_1[i] <- sum(!is.na(data[data[[group_col]] == group_levels[1], var, drop = TRUE]))  # Sample size for group 1
    results$N_2[i] <- sum(!is.na(data[data[[group_col]] == group_levels[2], var, drop = TRUE]))  # Sample size for group 2
  }
  
  names(results)[2:3] <- paste0("Mean_", group_levels)
  
  return(results)
}

# Example usage with your datasets and variable specifications
df_ttest_diff_diga <- run_unpaired_t_tests(df_diga_differences, scales_diga_diff_normal, "Classification")
df_ttest_diff_dipa <- run_unpaired_t_tests(df_dipa_differences, scales_dipa_diff_normal, "Classification")



# Mann-Whitney on the differences

run_mann_whitney_tests <- function(data, scale_vars, group_col) {
  if (!all(c(group_col, scale_vars) %in% names(data))) {
    stop("Specified columns do not exist in the dataframe.")
  }
  
  group_levels <- unique(data[[group_col]])
  if (length(group_levels) != 2) {
    stop("Group column must have exactly two levels.")
  }
  
  results <- data.frame(Variable = rep(scale_vars, each = 1),
                        Mean_1 = numeric(length(scale_vars)),
                        Mean_2 = numeric(length(scale_vars)),
                        U_Value = numeric(length(scale_vars)),
                        P_Value = numeric(length(scale_vars)),
                        N_1 = integer(length(scale_vars)),  # Add N for the first group
                        N_2 = integer(length(scale_vars)),  # Add N for the second group
                        stringsAsFactors = FALSE)
  
  for (i in seq_along(scale_vars)) {
    var <- scale_vars[i]
    mw_test_result <- wilcox.test(data[[var]] ~ data[[group_col]])
    
    results$Mean_1[i] <- mean(data[data[[group_col]] == group_levels[1], var, drop = TRUE], na.rm = TRUE)
    results$Mean_2[i] <- mean(data[data[[group_col]] == group_levels[2], var, drop = TRUE], na.rm = TRUE)
    results$U_Value[i] <- mw_test_result$statistic
    results$P_Value[i] <- mw_test_result$p.value
    results$N_1[i] <- sum(!is.na(data[data[[group_col]] == group_levels[1], var, drop = TRUE]))  # Count non-missing values
    results$N_2[i] <- sum(!is.na(data[data[[group_col]] == group_levels[2], var, drop = TRUE]))  # Count non-missing values
  }
  
  names(results)[2:3] <- paste0("Mean_", group_levels)
  
  return(results)
}


df_mannwhitney_diff_diga <- run_mann_whitney_tests(df_diga_differences, scales_diga_diff_nonnormal, "Classification")
df_mannwhitney_diff_dipa <- run_mann_whitney_tests(df_dipa_differences, scales_dipa_diff_nonnormal, "Classification")


library(ggplot2)
library(gridExtra)

# Additional Visualizations

generate_and_save_histograms <- function(data, variables, filename) {
  # Generate the list of plots
  plots <- lapply(variables, function(var) {
    ggplot(data, aes_string(x = var)) +
      geom_histogram(bins = 30, fill = "blue", color = "black") +
      theme_minimal() +
      ggtitle(paste(var))
  })
  
  # Determine the number of columns and rows dynamically
  n_plots <- length(plots)
  n_cols <- min(n_plots, 4)  # Setting the maximum number of columns to 4
  n_rows <- ceiling(n_plots / n_cols)  # Calculate number of rows needed
  
  # Combine plots into a single grid layout
  combined_plot <- do.call(grid.arrange, c(plots, ncol = n_cols, nrow = n_rows))
  print(combined_plot)
  
  # Save the combined plot to file
  ggsave(filename, plot = combined_plot, width = 20, height = 15)
}

  

generate_and_save_qq_plots <- function(data, variables, filename) {
  # Generate the list of QQ plots
  plots <- lapply(variables, function(var) {
    ggplot(data, aes_string(sample = var)) +
      stat_qq() +
      stat_qq_line() +
      theme_minimal() +
      ggtitle(paste(var))
  })
  
  # Determine the number of columns and rows dynamically
  n_plots <- length(plots)
  n_cols <- min(n_plots, 4)  # Set a practical upper limit for columns
  n_rows <- ceiling(n_plots / n_cols)  # Calculate the number of rows needed
  
  # Combine plots into a single grid layout
  combined_plot <- do.call(grid.arrange, c(plots, ncol = n_cols, nrow = n_rows))
  print(combined_plot)
  
  # Save the combined plot to file
  ggsave(filename, plot = combined_plot, width = 20, height = 15)
}


# For histograms
generate_and_save_histograms(df_diga, scales_diga, "histograms_diga.png")
generate_and_save_histograms(df_dipa, scales_dipa, "histograms_dipa.png")

# For QQ plots
generate_and_save_qq_plots(df_dipa,scales_dipa, "qq_plots_dipa.png")
generate_and_save_qq_plots(df_diga,scales_diga, "qq_plots_diga.png")

# For histograms
generate_and_save_histograms(df_diga_differences, scales_diga_diff, "histograms_diga_diff.png")
generate_and_save_histograms(df_dipa_differences, scales_dipa_diff, "histograms_dipa_diff.png")

# For QQ plots
generate_and_save_qq_plots(df_dipa_differences,scales_dipa_diff, "qq_plots_dipa_diff.png")
generate_and_save_qq_plots(df_diga_differences,scales_diga_diff, "qq_plots_diga_diff.png")


# Additional Boxplots

create_boxplot_comparison <- function(data, variables, factor_column1, factor_column2, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  
  # Transform the data into a long format for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(!!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1))
  
  # Create the boxplot
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Value, fill = !!sym(factor_column2))) +
    geom_boxplot() +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Boxplot Comparison between Pre and Post", 
      x = "Primary Factor",  
      y = "Value",
      fill = "Secondary Factor"
    ) +
    scale_fill_discrete(name = "Secondary Factor")
  
  return(p)
}

factor1_levels <- c("Pre", "Post")
factor_column2 <- "Classification"
factor_column1 <- "Pre_Post"

plot <- create_boxplot_comparison(df_diga, scales_diga, factor_column1, factor_column2, factor1_levels)
print(plot)
ggsave("boxplot_comparison_diga.png", plot = plot, width = 12, height = 8)

plot <- create_boxplot_comparison(df_dipa, scales_dipa, factor_column1, factor_column2, factor1_levels)
print(plot)
ggsave("boxplot_comparison_dipa.png", plot = plot, width = 12, height = 8)


# Dot and Whiskers for Differences

create_mean_ci_plot_no_secondary <- function(data, variables, factor_column1, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(!!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1)) %>%
    group_by(!!sym(factor_column1), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      N = n(),
      SE = SD / sqrt(N),
      CI_lower = Mean - 1.96 * SE,
      CI_upper = Mean + 1.96 * SE,
      .groups = 'drop'
    )
  
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Mean)) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = CI_lower, ymax = CI_upper), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mean and 95% Confidence Interval Plot", 
      x = "Primary Factor",  
      y = "Value"
    )
  
  return(p)
}


factor1_levels <- c("Control", "Treatment")
factor_column1 <- "Classification"

plot <- create_mean_ci_plot_no_secondary(df_diga_differences, scales_diga_diff, factor_column1,factor1_levels)
print(plot)
ggsave("mean_sd_plot_diga_diff.png", plot = plot, width = 12, height = 8)

plot <- create_mean_ci_plot_no_secondary(df_dipa_differences, scales_dipa_diff, factor_column1,factor1_levels)
print(plot)
ggsave("mean_sd_plot_dipa_diff.png", plot = plot, width = 12, height = 8)

# Boxplots

create_boxplot_comparison_no_secondary <- function(data, variables, factor_column1, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(!!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1))
  
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Value)) +
    geom_boxplot() +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Boxplot Comparison", 
      x = "Primary Factor",  
      y = "Value"
    )
  
  return(p)
}


plot <- create_boxplot_comparison_no_secondary(df_diga_differences, scales_diga_diff, factor_column1, factor1_levels)
print(plot)
ggsave("boxplot_comparison_diga_diff.png", plot = plot, width = 12, height = 8)

plot <- create_boxplot_comparison_no_secondary(df_dipa_differences, scales_dipa_diff, factor_column1, factor1_levels)
print(plot)
ggsave("boxplot_comparison_dipa_diff.png", plot = plot, width = 12, height = 8)

# Count N's

library(dplyr)
library(tidyr)

count_disaggregation_no_NAs <- function(data, variables, factor_column1, factor_column2) {
  # Transform the data into a long format for counting
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    filter(!is.na(Value)) %>%
    group_by(!!sym(factor_column1), !!sym(factor_column2), Variable) %>%
    summarise(Count = n(), .groups = 'drop')
  
  return(long_data)
}

disaggregated_counts_diga <- count_disaggregation_no_NAs(df_diga, scales_diga, "Pre_Post", "Classification")
disaggregated_counts_dipa <- count_disaggregation_no_NAs(df_dipa, scales_dipa, "Pre_Post", "Classification")

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
  "Counts - Dipa" = disaggregated_counts_dipa,
  "Counts - Diga" = disaggregated_counts_diga,
  "Normality - Dipa" = df_normality_results_dipa,
  "Normality - Diga" = df_normality_results_diga,
  "Normality - Dipa Diffs" = df_normality_results_diff_dipa,
  "Normality - Diga Diffs" = df_normality_results_diff_diga,
  "T-test Baseline Diga" = df_ttest_baseline_diga,
  "T-test Baseline Dipa" = df_ttest_baseline_dipa,
  "Mann-Whitney Baseline Diga" = df_mannwhitney_baseline_diga,
  "Mann-Whitney Baseline Dipa" = df_mannwhitney_baseline_dipa,
  "Paired T test - Control Diga" = df_paired_t_tests_control_diga, 
  "Paired T test - Treatment Diga" = df_paired_t_tests_treatment_diga ,
  "Paired T test - Control Dipa" = df_paired_t_tests_control_dipa ,
  "Paired T test - Treatment Dipa" = df_paired_t_tests_treatment_dipa,
  "Wilcoxon - Control Diga" =  df_wilcoxon_control_diga,
  "Wilcoxon - Control Dipa" =  df_wilcoxon_control_dipa,
  "Wilcoxon - Treatment Diga" =  df_wilcoxon_treatment_diga,
  "Wilcoxon - Treatment Dipa" =  df_wilcoxon_treatment_dipa,
  "T-test Differences - Diga" = df_ttest_diff_diga,
  "T-test Differences - Dipa" = df_ttest_diff_dipa,
  "Mann-Whitney Differences Diga" = df_mannwhitney_diff_diga,
  "Mann-Whitney Differences Dipa" = df_mannwhitney_diff_dipa
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_NewTests_withBloodTests2.xlsx")


