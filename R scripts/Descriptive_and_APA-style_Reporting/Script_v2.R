library(openxlsx)
library(dplyr)

# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/shaunlev")

# Specify the path and password to the protected Excel file
file_path <- "DataSet for Contribution History P.xlsx"

# Read the Excel file with password protection
df <- read.xlsx(file_path)

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

# Trim all values

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))

# Filter for member_status = "01"
df <- df %>%
  filter(member_status == "01")

# Converting numeric dates to Date objects
df$reporting_month <- as.Date(df$reporting_month, origin = "1899-12-30")
df$date_of_birth <- as.Date(df$date_of_birth, origin = "1899-12-30")

# Replace zeros in annual_pensionable_salary with NA
df <- df %>%
  mutate(annual_pensionable_salary = ifelse(annual_pensionable_salary == 0, NA, annual_pensionable_salary))

# Checking for missing values
summary(is.na(df))

# Generating new time variables
df$year <- format(df$reporting_month, "%Y")
df$month <- format(df$reporting_month, "%m")

# Identify first and last appearance for each case_mbr_key
df <- df %>%
  arrange(case_mbr_key, reporting_month) %>%
  group_by(case_mbr_key) %>%
  mutate(
    first_appearance = reporting_month == min(reporting_month),
    last_appearance = reporting_month == max(reporting_month)
  )


# Summary statistics for age grouped by gender
library(dplyr)
df_age_by_gender <- df %>%
  group_by(gender) %>%
  summarize(mean_age = mean(member_age, na.rm = TRUE),
            sd_age = sd(member_age, na.rm = TRUE))

# Calculate individual-level last entry
df_individual_last <- df %>%
  group_by(case_mbr_key) %>%
  arrange(desc(reporting_month)) %>%  # Ensure the data is sorted so the last entry is indeed the most recent
  slice(1) %>%  # Selects the first row in the descending order, which is the most recent
  ungroup()  # Ungroup to perform further operations

# Since the last entry selection removes the grouping, re-calculate averages if needed
df_individual_last <- df_individual_last %>%
  mutate(
    Last_Age = member_age,
    Last_Salary = annual_pensionable_salary,
    Last_Contribution_Rate = retirement_contribution_rate
  )


df_individual_stats <- df %>%
  group_by(case_mbr_key) %>%
  summarise(
    First_Date = min(reporting_month),
    Last_Date = max(reporting_month),
    Last_Age = last(member_age),
    Last_Salary = last(annual_pensionable_salary),
    Last_Contribution_Rate = last(retirement_contribution_rate),
    Gender = last(gender),
    .groups = 'drop'
  ) %>%
  mutate(
    Membership_Duration = as.integer(difftime(Last_Date, First_Date, units = "days")) / 365.25  # Convert to years
  )

# Calculate group-level averages based on individual averages
df_group_averages <- df_individual_last %>%
  
  summarise(
    Mean_Age = mean(Last_Age, na.rm = TRUE),
    Mean_Salary = mean(Last_Salary, na.rm = TRUE),
    Mean_Contribution_Rate = mean(Last_Contribution_Rate, na.rm = TRUE),
    .groups = 'drop'
  )

# Calculate group-level averages based on individual averages
df_group_averages_by_gender <- df_individual_last %>%
  group_by(gender) %>%
  summarise(
    Mean_Age = mean(Last_Age, na.rm = TRUE),
    Mean_Salary = mean(Last_Salary, na.rm = TRUE),
    Mean_Contribution_Rate = mean(Last_Contribution_Rate, na.rm = TRUE),
    .groups = 'drop'
  )


# Counting active and dropped out customers by month and year
df_customer_activity_stats <- df %>%
  group_by(year, month) %>%
  summarise(
    Active_Customers = n_distinct(case_mbr_key),
    Avg_Age = mean(member_age, na.rm = TRUE),
    Proportion_Females = mean(gender == "Female", na.rm = TRUE),
    Avg_Salary = mean(annual_pensionable_salary, na.rm = TRUE),
    Avg_Contribution_Rate = mean(retirement_contribution_rate, na.rm = TRUE),
    .groups = 'drop'
  )

# Create a summary of active customers by year and month
active_customers_summary <- df %>%
  group_by(year, month) %>%
  summarise(Active_Customers = n_distinct(case_mbr_key), .groups = 'drop')

# Join this summary back to the main df based on year and month
df <- df %>%
  left_join(active_customers_summary, by = c("year", "month"))

# First, ensure that 'year' and 'month' are in the correct format if not already
df$year <- as.integer(df$year)
df$month <- as.integer(df$month)

# Then, add flags for communication changes
df <- df %>%
  arrange(case_mbr_key, year, month) %>%
  group_by(case_mbr_key) %>%
  mutate(
    Change_2018 = as.integer((year == 2018 & month >= 8) | (year == 2019 & month < 8)),
    Change_2019 = as.integer((year == 2019 & month >= 8) | (year == 2020 & month < 8)),
    Change_2020 = as.integer((year == 2020 & month >= 8) | (year == 2021 & month < 8)),
    Change_2021 = as.integer((year == 2021 & month >= 8) | (year == 2022 & month < 8)),
    Change_2022 = as.integer((year == 2022 & month >= 8) | (year == 2023 & month < 8)),
    Change_2023 = as.integer(year >= 2023 & month >= 8)
  ) %>%
  ungroup()

df <- df %>%
  mutate(
    Pandemic = as.integer((year == 2020 & month >= 8) | (year == 2021) | (year == 2022 & month <= 7))
  )

# Log transformation
df <- df %>%
  mutate(log_retirement_contribution_rate = log(retirement_contribution_rate + 1))  # log(x + 1) to handle zero values

# Evaluate predictors

library(ggplot2)

df_clean <- df[!is.na(df$annual_pensionable_salary), ]

# Now try plotting again with the cleaned or filled data
plot <- ggplot(df_clean, aes(x = annual_pensionable_salary)) +
  geom_histogram(bins = 30, fill = "skyblue", color = "black") +
  labs(title = "Distribution of Annual Pensionable Salary", x = "Annual Pensionable Salary", y = "Count")

# Explicitly print the plot
print(plot)

# Now try plotting again with the cleaned or filled data
plot_age <- ggplot(df_clean, aes(x = member_age)) +
  geom_histogram(bins = 30, fill = "skyblue", color = "black") +
  labs(title = "Distribution of Age", x = "Age", y = "Count")

# Explicitly print the plot
print(plot_age)


# Log transformation
df <- df %>%
  mutate(log_annual_pensionable_salary = log(annual_pensionable_salary + 1))  # log(x + 1) to handle zero values

#Evaluate Outliers

calculate_z_scores <- function(data, vars, id_var, time_vars) {
  # Calculate z-scores for each variable
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~scale(.) %>% as.vector, .names = "z_{.col}")) %>%
    select(!!sym(id_var), all_of(time_vars), starts_with("z_"))
  
  return(z_score_data)
}

scales <- c("log_annual_pensionable_salary", "log_retirement_contribution_rate", "member_age")
id_var <- "case_mbr_key"
time_vars <- c("month", "year")

z_scores_dataset <- calculate_z_scores(df, scales, id_var, time_vars)

#Remove Outliers

# Merge the dataframes on 'case_mbr_key', 'Month', and 'Year'
merged_df <- df %>%
  left_join(z_scores_dataset, by = c("case_mbr_key", "month", "year"))

# Replace values outside the ±3 Z-score range with NA
df_cleaned <- merged_df %>%
  mutate(
    log_annual_pensionable_salary = ifelse(!is.na(z_log_annual_pensionable_salary) & (z_log_annual_pensionable_salary < -3 | z_log_annual_pensionable_salary > 3), NA, log_annual_pensionable_salary),
    log_retirement_contribution_rate = ifelse(!is.na(z_log_retirement_contribution_rate) & (z_log_retirement_contribution_rate < -3 | z_log_retirement_contribution_rate > 3), NA, log_retirement_contribution_rate),
    member_age = ifelse(!is.na(z_member_age) & (z_member_age < -3 | z_member_age > 3), NA, member_age)
  )

# Count replaced cases
replaced_counts <- merged_df %>%
  summarise(
    log_annual_pensionable_salary_removed = sum(!is.na(z_log_annual_pensionable_salary) & (z_log_annual_pensionable_salary < -3 | z_log_annual_pensionable_salary > 3), na.rm = TRUE),
    log_retirement_contribution_rate_removed = sum(!is.na(z_log_retirement_contribution_rate) & (z_log_retirement_contribution_rate < -3 | z_log_retirement_contribution_rate > 3), na.rm = TRUE),
    member_age_removed = sum(!is.na(z_member_age) & (z_member_age < -3 | z_member_age > 3), na.rm = TRUE)
  )

df <- df_cleaned

# Linear Mixed Model
library(lme4)
library(broom.mixed)
library(afex)
library(car)

# Function to fit LMM with multiple random effects and return formatted results
fit_lmm_and_format <- function(formula, data, save_plots = FALSE, plot_param = "") {
  # Fit the linear mixed-effects model
  lmer_model <- lmer(formula, data = data)
  
  # Print the summary of the model for fit statistics
  model_summary <- summary(lmer_model)
  print(model_summary)
  
  # Extract the tidy output for fixed effects and assign it to lmm_results
  lmm_results <- tidy(lmer_model, "fixed")
  
  # Optionally print the tidy output for fixed effects
  print(lmm_results)
  
  # Extract random effects variances
  random_effects_variances <- as.data.frame(VarCorr(lmer_model))
  print(random_effects_variances)
  
  # Initialize vif_values as NULL
  vif_values <- NULL
  
  # Calculate and print VIF for fixed effects if there are multiple fixed effects
  if (length(fixef(lmer_model)) > 1) {
    vif_values <- vif(lmer_model)
    print(vif_values)
  }
  
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
  
  # Add the sample size (N) to the fixed effects results
  sample_size <- nobs(lmer_model)
  lmm_results <- lmm_results %>%
    mutate(N = sample_size)
  
  # Return results
  return(list(fixed = lmm_results, random_variances = random_effects_variances, vif = vif_values))
}

# Only annual changes
lmm_formula_annualchanges <- log_retirement_contribution_rate ~ Change_2018 + Change_2019 + 
  Change_2020 + Change_2021 + Change_2022 + 
  Change_2023 + (1 | case_mbr_key)

# Use the custom function to fit the model and format results
df_model_results_annualchanges <- fit_lmm_and_format(
  lmm_formula_annualchanges, 
  df, 
  save_plots = TRUE, 
  plot_param = "annual_changes"
)
print(df_model_results_annualchanges)

df_model_results_annualchanges_pars <- df_model_results_annualchanges$fixed

# Control for Age
lmm_formula_annualchanges_age <- log_retirement_contribution_rate ~ Change_2018 + Change_2019 + 
  Change_2020 + Change_2021 + Change_2022 + 
  Change_2023 + member_age + (1 | case_mbr_key)

# Use the custom function to fit the model and format results
df_model_results_annualchangesage <- fit_lmm_and_format(
  lmm_formula_annualchanges_age, 
  df, 
  save_plots = TRUE, 
  plot_param = "annual_changes_age"
)
print(df_model_results_annualchangesage)

df_model_results_annualchangesage_pars <- df_model_results_annualchangesage$fixed

# Control for Age and Salary
lmm_formula_annualchanges_agesalary <- log_retirement_contribution_rate ~ Change_2018 + Change_2019 + 
  Change_2020 + Change_2021 + Change_2022 + 
  Change_2023 +  + member_age + log_annual_pensionable_salary + (1 | case_mbr_key)

# Use the custom function to fit the model and format results
df_model_results_annualchangesagesalary <- fit_lmm_and_format(
  lmm_formula_annualchanges_agesalary, 
  df, 
  save_plots = TRUE, 
  plot_param = "annual_changes_age_salary"
)
print(df_model_results_annualchangesagesalary)

df_model_results_annualchangesagesalary_pars <- df_model_results_annualchangesagesalary$fixed


# Visualizations

library(lubridate)
df_customer_activity_stats <- df_customer_activity_stats %>%
  mutate(Date = ymd(paste(year, month, "01", sep = "-")))

# Plot for Average Age
age_plot <- ggplot(df_customer_activity_stats, aes(x = Date, y = Avg_Age)) +
  geom_line(color = "blue") +
  labs(title = "Average Age Over Time", x = "Month/Year", y = "Average Age") +
  theme_minimal()

# Plot for Average Salary
salary_plot <- ggplot(df_customer_activity_stats, aes(x = Date, y = Avg_Salary)) +
  geom_line(color = "orange") +
  labs(title = "Average Salary Over Time", x = "Month/Year", y = "Average Salary") +
  theme_minimal()

# Plot for Average Contribution Rate
contribution_plot <- ggplot(df_customer_activity_stats, aes(x = Date, y = Avg_Contribution_Rate)) +
  geom_line(color = "red") +
  labs(title = "Average Contribution Rate Over Time", x = "Month/Year", y = "Average Contribution Rate") +
  theme_minimal()

# Print the plots
print(age_plot)
print(salary_plot)
print(contribution_plot)


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
  "Age By Gender" = df_age_by_gender,
  "Customer Activity" = df_customer_activity_stats,
  "Group Stats" = df_group_averages,
  "Group Stats by Gender" = df_group_averages_by_gender,
  "Individual Stats" = df_individual_stats,
  "Model Results - Only Changes" = df_model_results_annualchanges_pars,
  "Model Results - Age" = df_model_results_annualchangesage_pars,
  "Model Results - Salary" = df_model_results_annualchangesagesalary_pars
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")


# No New Entrants
library(dplyr)
library(lubridate)

# Filter for members active in January 2018
cohort_2018 <- df %>%
  filter(year(reporting_month) == 2018 & month(reporting_month) == 1) %>%
  select(case_mbr_key) %>%
  distinct()

# Filter the main dataset to include only these members
df_cohort <- df %>%
  filter(case_mbr_key %in% cohort_2018$case_mbr_key)


# Summary statistics for age grouped by gender in the cohort
df_age_by_gender_cohort <- df_cohort %>%
  group_by(gender) %>%
  summarize(mean_age = mean(member_age, na.rm = TRUE),
            sd_age = sd(member_age, na.rm = TRUE))

# Calculate individual-level last entry for the cohort
df_individual_last_cohort <- df_cohort %>%
  group_by(case_mbr_key) %>%
  arrange(desc(reporting_month)) %>%
  slice(1) %>%
  ungroup()

df_individual_last_cohort <- df_individual_last_cohort %>%
  mutate(
    Last_Age = member_age,
    Last_Salary = annual_pensionable_salary,
    Last_Contribution_Rate = retirement_contribution_rate
  )

df_individual_stats_cohort <- df_cohort %>%
  group_by(case_mbr_key) %>%
  summarise(
    First_Date = min(reporting_month),
    Last_Date = max(reporting_month),
    Last_Age = last(member_age),
    Last_Salary = last(annual_pensionable_salary),
    Last_Contribution_Rate = last(retirement_contribution_rate),
    Gender = last(gender),
    .groups = 'drop'
  ) %>%
  mutate(
    Membership_Duration = as.integer(difftime(Last_Date, First_Date, units = "days")) / 365.25  # Convert to years
  )

df_group_averages_cohort <- df_individual_last_cohort %>%
  summarise(
    Mean_Age = mean(Last_Age, na.rm = TRUE),
    Mean_Salary = mean(Last_Salary, na.rm = TRUE),
    Mean_Contribution_Rate = mean(Last_Contribution_Rate, na.rm = TRUE),
    .groups = 'drop'
  )

df_group_averages_by_gender_cohort <- df_individual_last_cohort %>%
  group_by(gender) %>%
  summarise(
    Mean_Age = mean(Last_Age, na.rm = TRUE),
    Mean_Salary = mean(Last_Salary, na.rm = TRUE),
    Mean_Contribution_Rate = mean(Last_Contribution_Rate, na.rm = TRUE),
    .groups = 'drop'
  )

# Customer activity stats for the cohort
df_customer_activity_stats_cohort <- df_cohort %>%
  group_by(year, month) %>%
  summarise(
    Active_Customers = n_distinct(case_mbr_key),
    Avg_Age = mean(member_age, na.rm = TRUE),
    Proportion_Females = mean(gender == "Female", na.rm = TRUE),
    Avg_Salary = mean(annual_pensionable_salary, na.rm = TRUE),
    Avg_Contribution_Rate = mean(retirement_contribution_rate, na.rm = TRUE),
    .groups = 'drop'
  )

# Create a summary of active customers by year and month for the cohort
active_customers_summary_cohort <- df_cohort %>%
  group_by(year, month) %>%
  summarise(Active_Customers = n_distinct(case_mbr_key), .groups = 'drop')

# Join this summary back to the main df_cohort based on year and month
df_cohort <- df_cohort %>%
  left_join(active_customers_summary_cohort, by = c("year", "month"))

# Adding flags for communication changes in the cohort
df_cohort <- df_cohort %>%
  arrange(case_mbr_key, year, month) %>%
  group_by(case_mbr_key) %>%
  mutate(
    CChange_2018 = as.integer((year == 2018 & month >= 8) | (year == 2019 & month < 8)),
    Change_2019 = as.integer((year == 2019 & month >= 8) | (year == 2020 & month < 8)),
    Change_2020 = as.integer((year == 2020 & month >= 8) | (year == 2021 & month < 8)),
    Change_2021 = as.integer((year == 2021 & month >= 8) | (year == 2022 & month < 8)),
    Change_2022 = as.integer((year == 2022 & month >= 8) | (year == 2023 & month < 8)),
    Change_2023 = as.integer(year >= 2023 & month >= 8)
  ) %>%
  ungroup()




# Log transformation for the cohort
df_cohort <- df_cohort %>%
  mutate(log_retirement_contribution_rate = log(retirement_contribution_rate + 1),
         log_annual_pensionable_salary = log(annual_pensionable_salary + 1))

# Linear Mixed Model for the cohort
lmm_formula_annualchanges_cohort <- log_retirement_contribution_rate ~ Change_2018 + Change_2019 + 
  Change_2020 + Change_2021 + Change_2022 + 
  Change_2023 + (1 | case_mbr_key)

df_model_results_annualchanges_cohort <- fit_lmm_and_format(
  lmm_formula_annualchanges_cohort, 
  df_cohort, 
  save_plots = TRUE, 
  plot_param = "annual_changes_cohort"
)

df_model_results_annualchanges_cohort_pars <- df_model_results_annualchanges_cohort$fixed

# Control for Age in the cohort
lmm_formula_annualchanges_age_cohort <- log_retirement_contribution_rate ~ Change_2018 + Change_2019 + 
  Change_2020 + Change_2021 + Change_2022 + 
  Change_2023 + member_age + (1 | case_mbr_key)

df_model_results_annualchangesage_cohort <- fit_lmm_and_format(
  lmm_formula_annualchanges_age_cohort, 
  df_cohort, 
  save_plots = TRUE, 
  plot_param = "annual_changes_age_cohort"
)

df_model_results_annualchangesage_cohort_pars <- df_model_results_annualchangesage_cohort$fixed

# Control for Age and Salary in the cohort
lmm_formula_annualchanges_agesalary_cohort <- log_retirement_contribution_rate ~ Change_2018 + Change_2019 + 
  Change_2020 + Change_2021 + Change_2022 + 
  Change_2023 + member_age + log_annual_pensionable_salary + (1 | case_mbr_key)

df_model_results_annualchangesagesalary_cohort <- fit_lmm_and_format(
  lmm_formula_annualchanges_agesalary_cohort, 
  df_cohort, 
  save_plots = TRUE, 
  plot_param = "annual_changes_age_salary_cohort"
)

df_model_results_annualchangesagesalary_cohort_pars <- df_model_results_annualchangesagesalary_cohort$fixed

# Visualizations for the cohort
df_customer_activity_stats_cohort <- df_customer_activity_stats_cohort %>%
  mutate(Date = ymd(paste(year, month, "01", sep = "-")))

# Plot for Average Age
age_plot_cohort <- ggplot(df_customer_activity_stats_cohort, aes(x = Date, y = Avg_Age)) +
  geom_line(color = "blue") +
  labs(title = "Average Age Over Time (Cohort)", x = "Month/Year", y = "Average Age") +
  theme_minimal()

# Plot for Average Salary
salary_plot_cohort <- ggplot(df_customer_activity_stats_cohort, aes(x = Date, y = Avg_Salary)) +
  geom_line(color = "orange") +
  labs(title = "Average Salary Over Time (Cohort)", x = "Month/Year", y = "Average Salary") +
  theme_minimal()

# Plot for Average Contribution Rate
contribution_plot_cohort <- ggplot(df_customer_activity_stats_cohort, aes(x = Date, y = Avg_Contribution_Rate)) +
  geom_line(color = "red") +
  labs(title = "Average Contribution Rate Over Time (Cohort)", x = "Month/Year", y = "Average Contribution Rate") +
  theme_minimal()

# Print the plots
print(age_plot_cohort)
print(salary_plot_cohort)
print(contribution_plot_cohort)




# Modelling the choice to increase

library(dplyr)

# Create a binary variable indicating if the contribution rate increased
df_cohort <- df_cohort %>%
  group_by(case_mbr_key) %>%
  mutate(
    contribution_rate_change = retirement_contribution_rate - lag(retirement_contribution_rate),
    increased_contribution = ifelse(contribution_rate_change > 0, 1, 0)
  ) %>%
  ungroup()

# Replace NAs with 0 for the first observation of each group where there is no lag value
df_cohort <- df_cohort %>%
  mutate(increased_contribution = ifelse(is.na(increased_contribution), 0, increased_contribution))


# Create the Change_ dummy variables
df_cohort <- df_cohort %>%
  mutate(
    Change_2018 = ifelse((year == 2018 & month == 8) | (year == 2019 & month == 4), 1, 0),
    Change_2019 = ifelse((year == 2019 & month == 8) | (year == 2020 & month == 4), 1, 0),
    Change_2022 = ifelse((year == 2022 & month == 8) | (year == 2023 & month == 4), 1, 0),
    Change_2023 = ifelse((year == 2023 & month == 8) | (year == 2024 & month == 4), 1, 0)
  )

# Filter to keep only the data for August
df_august <- df_cohort %>%
  filter(month == "08")

# Frequency Table
create_segmented_frequency_tables <- function(data, vars, segmenting_categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each segmenting category
  for (segment in segmenting_categories) {
    # Iterate over each variable for which frequencies are to be calculated
    for (var in vars) {
      # Ensure both the segmenting category and variable are treated as factors
      segment_factor <- factor(data[[segment]])
      var_factor <- factor(data[[var]])
      
      # Calculate counts segmented by the segmenting category
      counts <- table(segment_factor, var_factor)
      
      # Melt the table into a long format for easier handling
      freq_table <- as.data.frame(counts)
      names(freq_table) <- c("Segment", "Level", "Count")
      
      # Add the variable name and calculate percentages within each segment
      freq_table$Variable <- var
      freq_table <- freq_table %>%
        group_by(Segment) %>%
        mutate(Percentage = (Count / sum(Count)) * 100) %>%
        ungroup()
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

df_freq_choice_segmented <- create_segmented_frequency_tables(df_august,  "increased_contribution", "year")



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
formula <- increased_contribution ~ member_age + log_annual_pensionable_salary + 
  Change_2018 + Change_2019 + Change_2022 + Change_2023 + 
  (1 | case_mbr_key)
result <- fit_mixed_effects_logistic(formula, df_august)

df_params_salaryage <- result$model_results
df_fit_salaryage <- result$fit_indices


formula2 <- increased_contribution ~ member_age  +
  Change_2018 + Change_2019 + Change_2022 + Change_2023 + 
  (1 | case_mbr_key)
result2 <- fit_mixed_effects_logistic(formula2, df_august)

df_params_age <- result2$model_results
df_fit_age <- result2$fit_indices


formula3 <- increased_contribution ~  
Change_2018 + Change_2019 + Change_2022 + Change_2023 + 
  (1 | case_mbr_key)
result3 <- fit_mixed_effects_logistic(formula3, df_august)

df_params_onlychange <- result3$model_results
df_fit_onlychange <- result3$fit_indices


# Export New Results

# Example usage
data_list_cohort <- list(
  "Age By Gender (Cohort)" = df_age_by_gender_cohort,
  "Customer Activity (Cohort)" = df_customer_activity_stats_cohort,
  "Group Stats (Cohort)" = df_group_averages_cohort,
  "Group Stats by Gender (Cohort)" = df_group_averages_by_gender_cohort,
  "Individual Stats (Cohort)" = df_individual_stats_cohort,
  "Model Result-Changes(Cohort)" = df_params_onlychange,
  "Model Results - Age (Cohort)" = df_params_age,
  "Model Results - Salary (Cohort)" = df_params_salaryage,
  "Model Fit - Changes (Cohort)" = df_fit_onlychange,
  "Model Fit - Age (Cohort)" = df_fit_age,
  "Model Fit - Salary (Cohort)" = df_fit_salaryage
)

# Save to Excel with APA formatting for the cohort
save_apa_formatted_excel(data_list_cohort, "APA_Formatted_Tables_Cohort.xlsx")