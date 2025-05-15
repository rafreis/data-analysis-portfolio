# Reaching a Record HIgh

library(dplyr)

# Assuming df_cohort is your dataframe
df_cohort_record <- df %>%
  group_by(case_mbr_key) %>%
  mutate(
    # Calculate the cumulative maximum of the contribution rate up to the previous period
    cumulative_max = cummax(lag(retirement_contribution_rate, default = -Inf)),
    # Identify if the current contribution rate is higher than the previous maximum
    increased_contribution = ifelse(retirement_contribution_rate > cumulative_max, 1, 0),
    # Include the identification in the row before and after the change
    increased_contribution_lag1 = lag(increased_contribution, default = 0),
    increased_contribution_lead1 = lead(increased_contribution, default = 0)
  ) %>%
  # Combine the columns to get the final indicator
  mutate(increased_contribution_final = pmax(increased_contribution, increased_contribution_lag1, increased_contribution_lead1),
         # Include contribution rate change for the row before, current, and after
         contribution_rate_change = retirement_contribution_rate - lag(retirement_contribution_rate),
         contribution_rate_change_lag1 = lag(contribution_rate_change, default = 0),
         contribution_rate_change_lead1 = lead(contribution_rate_change, default = 0),
         contribution_rate_change_final = ifelse(increased_contribution_final == 1, pmax(contribution_rate_change, contribution_rate_change_lag1, contribution_rate_change_lead1, na.rm = TRUE), NA)
  ) %>%
  ungroup()

# Create a binary variable indicating if the contribution rate increased
df_cohort_record <- df_cohort_record %>%
  group_by(case_mbr_key) %>%
  mutate(
    # Calculate the change in contribution rate
    contribution_rate_change = retirement_contribution_rate - lag(retirement_contribution_rate),
    # Identify increases
    increased_contribution_2 = ifelse(contribution_rate_change > 0, 1, 0),
    # Include the identification in the row before and after the change
    increased_contribution_lag1_2 = lag(increased_contribution_2, default = 0),
    increased_contribution_lead1_2 = lead(increased_contribution_2, default = 0)
  ) %>%
  # Combine the columns to get the final indicator
  mutate(increased_contribution_final_2 = pmax(increased_contribution_2, increased_contribution_lag1_2, increased_contribution_lead1_2),
         # Include contribution rate change for the row before, current, and after
         contribution_rate_change_lag1_2 = lag(contribution_rate_change, default = 0),
         contribution_rate_change_lead1_2 = lead(contribution_rate_change, default = 0),
         contribution_rate_change_final_2 = ifelse(increased_contribution_final_2 == 1, pmax(contribution_rate_change, contribution_rate_change_lag1, contribution_rate_change_lead1, na.rm = TRUE), NA)
  ) %>%
  ungroup()

# Filter for members active in January 2018
cohort_2019 <- df %>%
  filter(year(reporting_month) == 2019 & month(reporting_month) == 4) %>%
  select(case_mbr_key) %>%
  distinct()

# Filter the main dataset to include only these members
df_cohort2019 <- df_cohort_record %>%
  filter(case_mbr_key %in% cohort_2019$case_mbr_key)



# Create a summary of active customers by year and month for the cohort
active_customers_summary_cohort <- df_cohort2019 %>%
  group_by(year, month) %>%
  summarise(Active_Customers = n_distinct(case_mbr_key), .groups = 'drop')

# Join this summary back to the main df_cohort based on year and month
df_cohort2019 <- df_cohort2019 %>%
  left_join(active_customers_summary_cohort, by = c("year", "month"))

df_cohort2019 <- df_cohort2019 %>%
  mutate(month = as.numeric(month))

# Adding flags for communication changes in the cohort
df_cohort2019 <- df_cohort2019 %>%
  arrange(case_mbr_key, year, month) %>%
  group_by(case_mbr_key) %>%
  mutate(
    Change_2019 = as.integer((year == 2019 & month >= 8) | (year == 2020 & month < 8)),
    Change_2020 = as.integer((year == 2020 & month >= 8) | (year == 2021 & month < 8)),
    Change_2021 = as.integer((year == 2021 & month >= 8) | (year == 2022 & month < 8)),
    Change_2022 = as.integer((year == 2022 & month >= 8) | (year == 2023 & month < 8)),
    Change_2023 = as.integer(year >= 2023 & month >= 8)
  ) %>%
  ungroup()


# Log transformation for the cohort
df_cohort2019 <- df_cohort2019 %>%
  mutate(log_retirement_contribution_rate = log(retirement_contribution_rate + 1),
         log_annual_pensionable_salary = log(annual_pensionable_salary + 1))


# Replace NAs with 0 for the first observation of each group where there is no lag value
df_cohort2019 <- df_cohort2019 %>%
  mutate(increased_contribution = ifelse(is.na(increased_contribution), 0, increased_contribution))

# Replace NAs with 0 for the first observation of each group where there is no lag value
df_cohort <- df_cohort2019

# Create separate Change dummy variables for April and August
df_cohort <- df_cohort %>%
  mutate(
    Change_April_2021 = ifelse(year == 2020 & month == 4, 1, 0),
    Change_August_2021 = ifelse(year == 2021 & month == 8, 1, 0),
    Change_April_2022 = ifelse(year == 2022 & month == 4, 1, 0),
    Change_August_2022 = ifelse(year == 2022 & month == 8, 1, 0),
    Change_April_2023 = ifelse(year == 2023 & month == 4, 1, 0),
    Change_August_2023 = ifelse(year == 2023 & month == 8, 1, 0)
  )

df_cohort <- df_cohort %>%
  mutate(reporting_month = ymd(reporting_month)) %>%
  filter(reporting_month > as.Date("2021-02-28"))

# Filter to keep only the data for August
df_august_april <- df_cohort %>%
  filter(month %in% c(4, 8))

df_august <- df_august_april


create_segmented_frequency_tables <- function(data, vars, primary_segmenting_categories, secondary_segmenting_categories) {
  all_freq_tables <- list()  # Initialize an empty list to store all frequency tables
  
  # Iterate over each primary segmenting category
  for (primary_segment in primary_segmenting_categories) {
    # Ensure the primary segmenting category is treated as a factor
    data[[primary_segment]] <- factor(data[[primary_segment]])
    
    # Iterate over each secondary segmenting category
    for (secondary_segment in secondary_segmenting_categories) {
      # Ensure the secondary segmenting category is treated as a factor
      data[[secondary_segment]] <- factor(data[[secondary_segment]])
      
      # Iterate over each variable for which frequencies are to be calculated
      for (var in vars) {
        # Ensure the variable is treated as a factor
        data[[var]] <- factor(data[[var]])
        
        # Calculate counts segmented by both segmenting categories
        counts <- table(data[[primary_segment]], data[[secondary_segment]], data[[var]])
        
        # Convert the table into a data frame format for easier handling
        freq_table <- as.data.frame(counts, responseName = "Count")
        names(freq_table) <- c("Primary_Segment", "Secondary_Segment", "Level", "Count")
        
        # Add the variable name
        freq_table$Variable <- var
        
        # Calculate percentages within each combination of primary segment, secondary segment, and variable level
        freq_table <- freq_table %>%
          group_by(Primary_Segment, Secondary_Segment) %>%
          mutate(Total_Count = sum(Count)) %>%
          ungroup() %>%
          mutate(Percentage = (Count / Total_Count) * 100)
        
        # Add the result to the list
        table_id <- paste(primary_segment, secondary_segment, var, sep = "_")
        all_freq_tables[[table_id]] <- freq_table
      }
    }
  }
  
  # Combine all frequency tables into a single data frame
  combined_freq_table <- do.call(rbind, all_freq_tables)
  
  return(combined_freq_table)
}

df_freq_gender <- create_segmented_frequency_tables(df_august_april, "increased_contribution_final", "reporting_month", "gender")
df_freq_contr <- create_segmented_frequency_tables(df_august_april, "increased_contribution_final" ,"increased_contribution_final_2", "reporting_month")


age_percentiles <- quantile(df_august_april$member_age, probs = c(0, 0.30, 0.70, 1), na.rm = TRUE)

# Print the percentiles
print(age_percentiles)

categorize <- function(data, variable, cutoffs, labels) {
  # Ensure that the number of labels is one more than the number of cutoffs
  if (length(labels) != length(cutoffs) + 1) {
    stop("The number of labels must be equal to the number of cutoffs plus one.")
  }
  
  # Create a new variable name by appending "_cat" to the original variable name
  new_var_name <- paste(variable, "cat", sep = "_")
  
  # Create a new variable for categorized data
  data[[new_var_name]] <- cut(
    data[[variable]],
    breaks = c(-Inf, cutoffs, Inf), 
    labels = labels,
    include.lowest = TRUE
  )
  
  return(data)
}

cutoffs <- c(35, 50)
labels <- c("Younger than 35", "35-49.9", "50 or older")
df_august_april <- categorize(df_august_april, "member_age", cutoffs, labels)

salary_percentiles <- quantile(df_august_april$annual_pensionable_salary, probs = c(0, 0.30, 0.70, 1), na.rm = TRUE)

# Print the percentiles
print(salary_percentiles)

cutoffs <- c(250000, 500000)
labels <- c("Lower than 250,000", "250,000 - 499,999", "500,000 or higher")
df_august_april <- categorize(df_august_april, "annual_pensionable_salary", cutoffs, labels)


df_freq_age <- create_segmented_frequency_tables(df_august_april, "increased_contribution_final", "reporting_month", "member_age_cat")
df_freq_salary <- create_segmented_frequency_tables(df_august_april, "increased_contribution_final", "reporting_month", "annual_pensionable_salary_cat")


library(dplyr)

calculate_descriptive_stats_bygroups <- function(data, desc_vars, cat_column1, cat_column2) {
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

df_desc_gender <- calculate_descriptive_stats_bygroups(df_august_april, "contribution_rate_change_final", "reporting_month", "gender")
df_desc_salary <- calculate_descriptive_stats_bygroups(df_august_april, "contribution_rate_change_final", "reporting_month", "annual_pensionable_salary_cat")
df_desc_age <- calculate_descriptive_stats_bygroups(df_august_april, "contribution_rate_change_final", "reporting_month", "member_age_cat")


# Overall

create_two_factor_frequency_table <- function(data, factor1, factor2) {
  # Ensure the category variables are factors
  data[[factor1]] <- factor(data[[factor1]])
  data[[factor2]] <- factor(data[[factor2]])
  
  # Calculate counts
  counts <- table(data[[factor1]], data[[factor2]])
  
  # Create a dataframe for this category pair
  freq_table <- as.data.frame(counts)
  colnames(freq_table) <- c(factor1, factor2, "Count")
  
  # Calculate and add percentages
  freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
  
  return(freq_table)
}


df_freq_overall <- create_two_factor_frequency_table(df_august_april, "increased_contribution_final", "reporting_month")
df_desc_overall <- calculate_descriptive_stats_bygroups(df_august_april, "contribution_rate_change_final", "reporting_month", "increased_contribution_final")


# Model

library(lme4)
library(broom.mixed)
library(MuMIn)
library(caret)
library(pROC)

fit_mixed_effects_logistic <- function(formula, data) {
  # Fit the mixed-effects logistic regression model
  logistic_model_mixed <- glmer(formula, family = binomial, data = data)
  
  # Get tidy output for the mixed-effects logistic regression model
  model_results <- tidy(logistic_model_mixed)
  
  # Calculate AIC and BIC
  aic_value <- AIC(logistic_model_mixed)
  bic_value <- BIC(logistic_model_mixed)
  
  # Calculate conditional and marginal R-squared
  r_squared_values <- r.squaredGLMM(logistic_model_mixed)
  
  # Calculate Log-Likelihood and Deviance
  log_likelihood <- logLik(logistic_model_mixed)
  deviance_value <- deviance(logistic_model_mixed)
  
  # Create a dataframe for fit indices
  fit_indices <- data.frame(
    AIC = aic_value,
    BIC = bic_value,
    Log_Likelihood = as.numeric(log_likelihood),
    Deviance = deviance_value,
    Marginal_R2 = r_squared_values[1, "R2m"],
    Conditional_R2 = r_squared_values[1, "R2c"]
  )
  
  return(list(model_results = model_results, fit_indices = fit_indices))
}

# Example usage
formula <- increased_contribution_final ~ member_age + log_annual_pensionable_salary + Change_2021 +
  Change_2022 + Change_2023 + 
  (1 | case_mbr_key)
result <- fit_mixed_effects_logistic(formula, df_august_april)

df_params_salaryage <- result$model_results
df_fit_salaryage <- result$fit_indices

formula2 <- increased_contribution_final ~ member_age  + Change_2021 +
  Change_2022 + Change_2023 + 
  (1 | case_mbr_key)
result2 <- fit_mixed_effects_logistic(formula2, df_august)

df_params_age <- result2$model_results
df_fit_age <- result2$fit_indices

formula3 <- increased_contribution_final ~  Change_2021 +
  Change_2022 + Change_2023 + 
  (1 | case_mbr_key)
result3 <- fit_mixed_effects_logistic(formula3, df_august)

df_params_onlychange <- result3$model_results
df_fit_onlychange <- result3$fit_indices


# Example usage
data_list_cohort <- list(
  "Model_recordHigh" = df_params_salaryage,
  "Model_recordHigh - Fit" = df_fit_salaryage
)

# Save to Excel with APA formatting for the cohort
save_apa_formatted_excel(data_list_cohort, "APA_Formatted_Tables_Cohort_RecordHigh.xlsx")

# Example usage
data_list_cohort <- list(
  "CrossTab - Record High" = df_freq_contr
)

# Save to Excel with APA formatting for the cohort
save_apa_formatted_excel(data_list_cohort, "APA_Formatted_Tables_CrossTab_RecordHigh2.xlsx")
