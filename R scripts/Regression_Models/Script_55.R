# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/tonivc")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Deal_Pipeline_All_Opportunities__export_Nov-08-2024 (10.11) Kopie.xlsx", sheet = "Data Excerpt Nov8")

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

str(df)
head(df)


library(dplyr)
library(tidyr)
library(stringr)

# Create binary columns for each unique reason
df <- df %>%
  mutate(Pass.Reason = strsplit(as.character(Pass.Reason), ", ")) %>% 
  unnest(Pass.Reason) %>%
  mutate(value = 1) %>%
  pivot_wider(names_from = Pass.Reason, values_from = value, values_fill = list(value = 0)) 

# View result
head(df)


# Filter rows where Pass.Reason is not NA
df_cleaned <- df %>%
  filter(!is.na(Time.to.Pass))

names(df_cleaned) <- gsub(" ", "_", trimws(names(df_cleaned)))
names(df_cleaned) <- gsub("\\s+", "_", trimws(names(df_cleaned), whitespace = "[\\h\\v\\s]+"))
names(df_cleaned) <- gsub("\\(", "_", names(df_cleaned))
names(df_cleaned) <- gsub("\\)", "_", names(df_cleaned))
names(df_cleaned) <- gsub("\\-", "_", names(df_cleaned))
names(df_cleaned) <- gsub("/", "_", names(df_cleaned))
names(df_cleaned) <- gsub("\\\\", "_", names(df_cleaned)) 
names(df_cleaned) <- gsub("\\?", "", names(df_cleaned))
names(df_cleaned) <- gsub("\\'", "", names(df_cleaned))
names(df_cleaned) <- gsub("\\,", "_", names(df_cleaned))
names(df_cleaned) <- gsub("\\$", "", names(df_cleaned))
names(df_cleaned) <- gsub("\\+", "_", names(df_cleaned))


df_cleaned <- df_cleaned %>%
  filter(No_answer != 1) %>%  # Remove rows where No_Answer == 1
  select(-No_answer)          # Drop the No_Answer column

df_cleaned <- df_cleaned %>%
  filter("NA" != 1) %>%  # Remove rows where No_Answer == 1
  select(-"NA")          # Drop the No_Answer column

# View the cleaned dataframe
head(df_cleaned)

colnames(df_cleaned)


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

scales_freq <- c("Funding.Round"   ,              
 "Transaction.Type"        ,       "Vertical"               ,       "Status" ,                        
 "Origination.of.the.deal"  ,   "Last.Status.before.Pass")

df_freq <- create_frequency_tables(df, scales_freq)



# Load required libraries
library(dplyr)
library(ggplot2)
library(car)

# Step 2: Exploratory Data Analysis (EDA)

# Summary statistics
summary(df_cleaned$Time.to.Pass)

# Check for missing values
sum(is.na(df_cleaned$Time.to.Pass))

# Histogram
ggplot(df_cleaned, aes(x = Time.to.Pass)) +
  geom_histogram(binwidth = 3, fill = "blue", alpha = 0.6) +
  labs(title = "Distribution of Time to Pass", x = "Time to Pass", y = "Frequency")

# Boxplot
ggplot(df_cleaned, aes(y = Time.to.Pass)) +
  geom_boxplot(fill = "red", alpha = 0.6) +
  labs(title = "Boxplot of Time to Pass", y = "Time to Pass")

# Log-transformation

df_cleaned <- df_cleaned %>%
  mutate(Log_Time_to_Pass = log1p(Time.to.Pass))  # log1p ensures log(0) is handled correctly


# Histogram
ggplot(df_cleaned, aes(x = Log_Time_to_Pass)) +
  geom_histogram(binwidth = 0.1, fill = "blue", alpha = 0.6) +
  labs(title = "Distribution of Log-Time to Pass", x = "Log-Time to Pass", y = "Frequency")

# Boxplot
ggplot(df_cleaned, aes(y = Log_Time_to_Pass)) +
  geom_boxplot(fill = "red", alpha = 0.6) +
  labs(title = "Boxplot of Log-Time to Pass", y = "Log-Time to Pass")


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

vars <- c("Time.to.Pass", "Log_Time_to_Pass")

# Example usage with your dataframe and list of variables
df_descriptive_stats <- calculate_descriptive_stats(df_cleaned, vars)

colnames(df_cleaned)

# Check if both columns exist
if ("not_growing" %in% colnames(df_cleaned) & "Not_growing" %in% colnames(df_cleaned)) {
  
  # Merge the two columns by taking the maximum value (since they are binary)
  df_cleaned <- df_cleaned %>%
    mutate(Not_growing = pmax(not_growing, Not_growing, na.rm = TRUE)) # Remove the redundant column
  
  print("Merged Not_growing and Not_Growing successfully!")
} else {
  print("One or both columns do not exist in df_cleaned.")
}

# List of correctly named binary Pass Reason variables
binary_reasons <- c("Out_of_scope"      ,            
                     "Too_early"         ,             "No_VCs_on_cap_table"    ,        "Product"        ,               
                     "Not_growing"        ,            "Broken_cap_table"        ,       "Too_expensive"   ,              
                    "Geography"            ,          "Business_Model"            ,     "Timing"            ,            
                    "Total_addressable_market"  ,     "Team"              ,            
                    "Direct_competitor"      ,         "Too_late_stage"                      
)


# Load required libraries
library(dplyr)
library(boot)


# Convert binary Pass Reason columns to factors
df_cleaned <- df_cleaned %>%
  mutate(across(all_of(binary_reasons), as.factor))  # Convert all binary columns to factors

str(df_cleaned)
# Function for bootstrapped t-tests
boot_t_test <- function(data, group_col, measurement_col, n_boot = 1000) {
  # Check if the variable is binary (only 0 and 1)
  group_levels <- unique(na.omit(data[[group_col]]))
  if (length(group_levels) != 2 || !all(group_levels %in% c(0, 1))) {
    print(paste("Skipping", group_col, "- Not a properly formatted binary variable."))
    return(NULL)
  }
  
  # Extract group data
  group1_data <- na.omit(data %>% filter(.data[[group_col]] == 1) %>% pull(measurement_col))
  group2_data <- na.omit(data %>% filter(.data[[group_col]] == 0) %>% pull(measurement_col))
  
  # Debugging: Print unique values
  print(paste("Processing:", group_col))
  print(paste("Group 1 size:", length(group1_data)))
  print(paste("Group 2 size:", length(group2_data)))
  
  # Skip if groups are too small
  if (length(group1_data) < 5 || length(group2_data) < 5) {
    print(paste("Skipping", group_col, "- Not enough observations in one or both groups."))
    return(NULL)
  }
  
  # Skip if data are constant (all values identical)
  if (sd(group1_data) == 0 || sd(group2_data) == 0) {
    print(paste("Skipping", group_col, "- Data are essentially constant."))
    return(NULL)
  }
  
  # Run t-test
  t_test_result <- t.test(group1_data, group2_data, var.equal = TRUE)
  
  # Extract statistics for APA reporting
  mean_group1 <- mean(group1_data)
  mean_group2 <- mean(group2_data)
  sd_group1 <- sd(group1_data)
  sd_group2 <- sd(group2_data)
  t_value <- t_test_result$statistic
  p_value <- t_test_result$p.value
  conf_low <- t_test_result$conf.int[1]
  conf_high <- t_test_result$conf.int[2]
  
  # Bootstrapped t-test function
  boot_fn <- function(data, indices) {
    resampled <- data[indices, ]
    t_resampled <- t.test(resampled$group1, resampled$group2, var.equal = TRUE)
    return(t_resampled$statistic)
  }
  
  # Ensure equal group sizes for bootstrapping
  min_size <- min(length(group1_data), length(group2_data))
  boot_data <- data.frame(
    group1 = sample(group1_data, min_size, replace = TRUE),
    group2 = sample(group2_data, min_size, replace = TRUE)
  )
  
  # Debugging: Print sample data structure
  print(str(boot_data))
  
  # Run bootstrapping
  boot_result <- boot(data = boot_data, statistic = boot_fn, R = n_boot)
  
  # Return statistics in a dataframe
  return(data.frame(
    Variable = group_col,
    Mean_Group1 = mean_group1,
    SD_Group1 = sd_group1,
    Mean_Group2 = mean_group2,
    SD_Group2 = sd_group2,
    T_Value = t_value,
    P_Value = p_value,
    Conf_Low = conf_low,
    Conf_High = conf_high,
    Boot_CI_Low = boot.ci(boot_result, type = "perc")$percent[4],  # Bootstrapped CI Lower Bound
    Boot_CI_High = boot.ci(boot_result, type = "perc")$percent[5]  # Bootstrapped CI Upper Bound
  ))
}

# Ensure only existing columns are tested
binary_reasons <- binary_reasons[binary_reasons %in% colnames(df_cleaned)]

# Initialize results dataframe
boot_results <- data.frame(
  Variable = character(),
  Mean_Group1 = numeric(),
  SD_Group1 = numeric(),
  Mean_Group2 = numeric(),
  SD_Group2 = numeric(),
  T_Value = numeric(),
  P_Value = numeric(),
  Conf_Low = numeric(),
  Conf_High = numeric(),
  Boot_CI_Low = numeric(),
  Boot_CI_High = numeric(),
  stringsAsFactors = FALSE
)

# Run bootstrapped t-tests for each binary reason
for (reason in binary_reasons) {
  result <- boot_t_test(df_cleaned, reason, "Log_Time_to_Pass")
  
  if (!is.null(result)) {
    boot_results <- rbind(boot_results, result)
  }
}

# View full bootstrapped t-test results
print(boot_results)

# OLS Regression
library(dplyr)
library(broom)
library(car)  # For VIF calculation

# Load required libraries
fit_ols_and_format <- function(data, predictors, response_vars, save_plots = FALSE) {
  ols_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Fit the OLS multiple regression model
    lm_model <- lm(formula, data = data, na.action = na.omit)
    
    # Save lm_model to the global environment for debugging
    assign("lm_model", lm_model, envir = .GlobalEnv)
    
    # Check for aliasing (perfect collinearity)
    alias_info <- alias(lm_model)$Complete
    if (!is.null(alias_info)) {
      print("Warning: The following predictors are perfectly collinear and will be removed:")
      print(colnames(alias_info))
      
      # Remove collinear predictors from the model
      predictors <- setdiff(predictors, colnames(alias_info))
      
      # Refit the model after removing collinear variables
      formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
      lm_model <- lm(formula, data = data, na.action = na.omit)
      
      # Save updated model to environment
      assign("lm_model", lm_model, envir = .GlobalEnv)
    }
    
    # Get the summary of the model
    model_summary <- summary(lm_model)
    
    # Extract the R-squared, F-statistic, and p-value
    r_squared <- model_summary$r.squared
    f_statistic <- model_summary$fstatistic[1]
    p_value <- pf(f_statistic, model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE)
    
    # Print the summary of the model for fit statistics
    print(summary(lm_model))
    
    # Compute VIF only if there are at least two predictors
    if (length(predictors) > 1) {
      vif_values <- car::vif(lm_model)
      vif_df <- data.frame(Variable = names(vif_values), VIF = vif_values)
    } else {
      vif_df <- data.frame(Variable = predictors, VIF = NA)  # If only one predictor, VIF is NA
    }
    
    # Extract model results and format them
    ols_results <- broom::tidy(lm_model) %>%
      mutate(ResponseVariable = response_var, R_Squared = r_squared, F_Statistic = f_statistic, P_Value = p_value) %>%
      left_join(vif_df, by = c("term" = "Variable"))  # Merge VIF values with model results
    
    # Optionally print the tidy output
    print(ols_results)
    
    # Store the results in a list
    ols_results_list[[response_var]] <- ols_results
  }
  
  # Return the list of model results
  return(ols_results_list)
}



colnames(df_cleaned)
control_vars <- c("Ticket.size" , "Round.Size", "Distance.to.Vienna")

df_cleaned <- df_cleaned %>%
  mutate(across(all_of(control_vars), as.numeric))


binary_reasons_clean <- c("Out_of_scope"      ,            
                    "Too_early"         ,             "No_VCs_on_cap_table"    ,        "Product"        ,               
                    "Not_growing"        ,            "Broken_cap_table"        ,       "Too_expensive"   ,              
                    "Geography"            ,          "Business_Model"            ,     "Timing"            ,            
                    "Total_addressable_market"  ,     "Team"              ,            
                    "Direct_competitor"      ,           "Too_late_stage"            
)

ivs <- c(binary_reasons_clean, control_vars)
dv <- "Log_Time_to_Pass"


df_cleaned <- df_cleaned %>%
  mutate(across(all_of(binary_reasons_clean), as.numeric))  # Convert all binary columns to factors

ols_results_list <- fit_ols_and_format(
  data = df_cleaned,
  predictors = ivs,  # Replace with actual predictor names
  response_vars = dv,       # Replace with actual response variable names
  save_plots = TRUE
)

alias(lm_model)

# Sum across all binary variables
dummyver <- df_cleaned %>%
  rowwise() %>%
  mutate(Total = sum(c_across(starts_with("num_")), na.rm = TRUE)) %>%
  ungroup()

# Check results
dummyver %>%
  summarise(mean_Total = mean(Total), min_Total = min(Total), max_Total = max(Total))


# Combine the results into a single dataframe
df_modelresults<- bind_rows(ols_results_list)

ivs_nocontrol <- c(binary_reasons_clean)

ols_results_list_nocontrol <- fit_ols_and_format(
  data = df_cleaned,
  predictors = ivs_nocontrol,  # Replace with actual predictor names
  response_vars = dv,       # Replace with actual response variable names
  save_plots = TRUE
)



# Combine the results into a single dataframe
df_modelresults_nocontrol<- bind_rows(ols_results_list_nocontrol)


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
  "Frequency Table" = df_freq, 
  "Descriptive Stats" = df_descriptive_stats, 
  "T-tests" = boot_results,
  "Model - no control" = df_modelresults_nocontrol,
  "Model - control" = df_modelresults
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

# Evaluate Outliers


library(dplyr)

calculate_z_scores <- function(data, vars, id_var, z_threshold = 3) {
  # Prepare data by handling NA values and calculating Z-scores
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~ as.numeric(.))) %>%
    mutate(across(all_of(vars), ~ replace(., is.na(.), mean(., na.rm = TRUE)))) %>%
    mutate(across(all_of(vars), ~ (.-mean(., na.rm = TRUE))/ifelse(sd(.) == 0, 1, sd(., na.rm = TRUE)), .names = "z_{.col}"))
  
  # Include original raw scores and Z-scores
  z_score_data <- z_score_data %>%
    select(!!sym(id_var), all_of(vars), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  return(z_score_data)
}

# Run the function on your dataframe
df_zscores <- calculate_z_scores(df_cleaned, "Time.to.Pass", "Opportunity.Id", 3.5)


# Function to remove values in df_recoded based on z-score table
library(dplyr)
library(tidyr)
library(stringr)

# Function to remove outlier values based on Z-score table
remove_outliers_based_on_z_scores <- function(data, z_score_data, id_var) {
  # Reshape the z-score data to long format
  z_score_data_long <- z_score_data %>%
    pivot_longer(cols = starts_with("z_"), names_to = "Variable", values_to = "z_value") %>%
    pivot_longer(cols = starts_with("flag_z_"), names_to = "Flag_Variable", values_to = "Flag_Value") %>%
    filter(str_replace(Flag_Variable, "flag_", "") == Variable) %>%
    select(!!sym(id_var), Variable, Flag_Value)
  
  # Remove flagged outliers from the original data
  for (i in 1:nrow(z_score_data_long)) {
    if (z_score_data_long$Flag_Value[i] == "Outlier") {
      col_name <- gsub("z_", "", z_score_data_long$Variable[i]) # Get original column name
      data <- data %>%
        mutate(!!sym(col_name) := ifelse(!!sym(id_var) == z_score_data_long[[id_var]][i], NA, !!sym(col_name)))
    }
  }
  
  return(data)
}

# Example usage
df_cleaned_nooutliers <- remove_outliers_based_on_z_scores(df_cleaned, df_zscores, id_var = "Opportunity.Id")


#Models

library(broom)

# Poisson
poisson_model <- glm(Time.to.Pass ~ Out_of_scope + Too_early + No_VCs_on_cap_table + Product + 
                       Not_growing + Broken_cap_table + Too_expensive + Geography + Business_Model + 
                       Timing + Total_addressable_market + Team + Direct_competitor + Too_late_stage + 
                       Ticket.size + Round.Size + Distance.to.Vienna,
                     data = df_cleaned_nooutliers, family = poisson(link = "log"))

# Negative Binomial
library(MASS)
nb_model <- glm.nb(Time.to.Pass ~ Out_of_scope + Too_early + No_VCs_on_cap_table + Product + 
                     Not_growing + Broken_cap_table + Too_expensive + Geography + Business_Model + 
                     Timing + Total_addressable_market + Team + Direct_competitor + Too_late_stage + 
                     Ticket.size + Round.Size + Distance.to.Vienna,
                   data = df_cleaned_nooutliers)


# Poisson regression results
df_poisson <- tidy(poisson_model) %>%
  mutate(model = "Poisson")

# Negative binomial regression results
df_nb <- tidy(nb_model) %>%
  mutate(model = "Negative Binomial")

df_poisson_exp <- tidy(poisson_model, exponentiate = TRUE) %>%
  mutate(model = "Poisson (exp)")

df_nb_exp <- tidy(nb_model, exponentiate = TRUE) %>%
  mutate(model = "Negative Binomial (exp)")

df_models_combined <- bind_rows(df_poisson, df_nb)
df_models_combined_exp <- bind_rows(df_poisson_exp, df_nb_exp)

library(performance)
library(see)
library(patchwork)  # For combining plots
library(DHARMa)

# Check residuals and diagnostics
check_poisson <- check_model(poisson_model)

# Plot diagnostics
plot(check_poisson)

check_nb <- check_model(nb_model)
plot(check_nb)

# Compare AIC values
AIC(poisson_model, nb_model)

str(df_cleaned_nooutliers)


# New Model

df_cleaned_nooutliers <- df_cleaned_nooutliers %>%
  mutate(Inbound_friendly_VC = ifelse(Origination.of.the.deal == "Inbound - friendly VC", 1, 0))


df_cleaned_nooutliers <- df_cleaned_nooutliers %>%
  mutate(
    Rejected_after_Opportunity = ifelse(Last.Status.before.Pass == "Opportunity", 1, 0),
    Rejected_after_Qualified_Opportunity = ifelse(Last.Status.before.Pass == "Qualified Opportunity", 1, 0),
    Rejected_after_Investment_Candidate = ifelse(Last.Status.before.Pass == "Investment Candidate", 1, 0),
    Rejected_after_IC1 = ifelse(Last.Status.before.Pass == "IC 1 approved", 1, 0)
  )


library(dplyr)
library(broom)
library(performance)
library(see)
library(ggplot2)

# 1. Create binary dependent variables
df_cleaned_nooutliers <- df_cleaned_nooutliers %>%
  mutate(
    reject_after_opportunity = as.numeric(Last.Status.before.Pass == "Opportunity"),
    reject_after_qualified = as.numeric(Last.Status.before.Pass == "Qualified Opportunity"),
    reject_after_candidate = as.numeric(Last.Status.before.Pass == "Investment Candidate"),
    reject_after_ic1 = as.numeric(Last.Status.before.Pass == "IC 1 approved"),
    inbound_friendly_vc = as.numeric(`Origination.of.the.deal` == "Inbound - friendly VC")
  )

# Ensure predictors are numeric
df_cleaned_nooutliers <- df_cleaned_nooutliers %>%
  mutate(across(c(Ticket.size, Round.Size, Distance.to.Vienna), as.numeric))

# Define predictors
predictors <- c("Out_of_scope", "Too_early", "No_VCs_on_cap_table", "Product", 
                "Not_growing", "Broken_cap_table", "Too_expensive", "Geography", 
                "Business_Model", "Timing", "Total_addressable_market", "Team", 
                "Direct_competitor", "Too_late_stage", "inbound_friendly_vc", 
                "Ticket.size", "Round.Size", "Distance.to.Vienna")

# MODEL 1: Rejection after Opportunity
logit_opportunity <- glm(reject_after_opportunity ~ ., 
                         data = df_cleaned_nooutliers[, c("reject_after_opportunity", predictors)], 
                         family = binomial)

# MODEL 2: Rejection after Qualified Opportunity
logit_qualified <- glm(reject_after_qualified ~ ., 
                       data = df_cleaned_nooutliers[, c("reject_after_qualified", predictors)], 
                       family = binomial)

# MODEL 3: Rejection after Investment Candidate
logit_candidate <- glm(reject_after_candidate ~ ., 
                       data = df_cleaned_nooutliers[, c("reject_after_candidate", predictors)], 
                       family = binomial)


# Performance diagnostics
check_model(logit_opportunity)
check_model(logit_qualified)
check_model(logit_candidate)


library(broom)
library(dplyr)

# Create tidy dataframes for each model
df_logit_opportunity <- tidy(logit_opportunity) %>%
  mutate(model = "Rejection after Opportunity")

df_logit_qualified <- tidy(logit_qualified) %>%
  mutate(model = "Rejection after Qualified Opportunity")

df_logit_candidate <- tidy(logit_candidate) %>%
  mutate(model = "Rejection after Investment Candidate")



# Combine all into a single tidy dataframe
df_logit_all <- bind_rows(
  df_logit_opportunity,
  df_logit_qualified,
  df_logit_candidate
)

# Fit statistics

# Extract fit statistics
fit_poisson <- model_performance(poisson_model)
fit_nb <- model_performance(nb_model)

# View tidy results
df_fit_poisson <- as.data.frame(fit_poisson)
df_fit_nb <- as.data.frame(fit_nb)

fit_logit_opportunity <- model_performance(logit_opportunity)
fit_logit_qualified <- model_performance(logit_qualified)
fit_logit_candidate <- model_performance(logit_candidate)

df_fit_logit_opportunity <- as.data.frame(fit_logit_opportunity)
df_fit_logit_qualified <- as.data.frame(fit_logit_qualified)
df_fit_logit_candidate <- as.data.frame(fit_logit_candidate)

df_fit_all <- bind_rows(
  df_fit_poisson %>% mutate(model = "Poisson"),
  df_fit_nb %>% mutate(model = "Negative Binomial"),
  df_fit_logit_opportunity %>% mutate(model = "Logit - Rejection after Opportunity"),
  df_fit_logit_qualified %>% mutate(model = "Logit - Rejection after Qualified Opportunity"),
  df_fit_logit_candidate %>% mutate(model = "Logit - Rejection after Candidate")
)


# Updated data_list including all new model outputs
data_list <- list(
  "Poisson Coefficients" = df_poisson,
  "Negative Binomial Coef" = df_nb,
  "Poisson Coef (exp)" = df_poisson_exp,
  "Negative Binomial Coef (exp)" = df_nb_exp,
  "Logit Models (All Stages)" = df_logit_all,
  "Fit Statistics All Models" = df_fit_all
)

# Save all to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_NewModels.xlsx")


colnames(df_cleaned_nooutliers)


library(fastDummies)

df_cleaned_nooutliers <- dummy_cols(df_cleaned_nooutliers, select_columns = "Origination.of.the.deal")

names(df_cleaned_nooutliers) <- gsub(" ", "_", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("\\(", "_", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("\\)", "_", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("\\-", "_", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("/", "_", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("\\\\", "_", names(df_cleaned_nooutliers)) 
names(df_cleaned_nooutliers) <- gsub("\\?", "", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("\\'", "", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("\\,", "_", names(df_cleaned_nooutliers))
names(df_cleaned_nooutliers) <- gsub("&", "_", names(df_cleaned_nooutliers))

# Define predictors
predictors2 <- c("Out_of_scope", "Too_early", "No_VCs_on_cap_table", "Product", 
                "Not_growing", "Broken_cap_table", "Too_expensive", "Geography", 
                "Business_Model", "Timing", "Total_addressable_market", "Team", 
                "Direct_competitor", "Too_late_stage", 
                "Ticket.size", "Round.Size", "Distance.to.Vienna",
                "Origination.of.the.deal_already_in_the_portfolio"    ,   
                "Origination.of.the.deal_Inbound___advisor___consultants", "Origination.of.the.deal_Inbound___founder_direct"       ,
                "Origination.of.the.deal_Inbound___friendly_VC"   ,        "Origination.of.the.deal_Outbound___events___conferences",
                "Origination.of.the.deal_Outbound___general"       ,       "Origination.of.the.deal_Outbound___recall"  ,            
                "Origination.of.the.deal_Outbound___research"       ,      "Origination.of.the.deal_NA")

# MODEL 1: Rejection after Opportunity
logit_opportunity2 <- glm(reject_after_opportunity ~ ., 
                         data = df_cleaned_nooutliers[, c("reject_after_opportunity", predictors2)], 
                         family = binomial)

# MODEL 2: Rejection after Qualified Opportunity
logit_qualified2 <- glm(reject_after_qualified ~ ., 
                       data = df_cleaned_nooutliers[, c("reject_after_qualified", predictors2)], 
                       family = binomial)

# MODEL 3: Rejection after Investment Candidate
logit_candidate2 <- glm(reject_after_candidate ~ ., 
                       data = df_cleaned_nooutliers[, c("reject_after_candidate", predictors2)], 
                       family = binomial)


# Performance diagnostics
check_model(logit_opportunity2)
check_model(logit_qualified2)
check_model(logit_candidate2)






library(broom)
library(dplyr)

# Create tidy dataframes for each model
df_logit_opportunity2 <- tidy(logit_opportunity2) %>%
  mutate(model = "Rejection after Opportunity")

df_logit_qualified2 <- tidy(logit_qualified2) %>%
  mutate(model = "Rejection after Qualified Opportunity")

df_logit_candidate2 <- tidy(logit_candidate2) %>%
  mutate(model = "Rejection after Investment Candidate")

# Combine all into a single tidy dataframe
df_logit_all2 <- bind_rows(
  df_logit_opportunity2,
  df_logit_qualified2,
  df_logit_candidate2
)


fit_logit_opportunity2 <- model_performance(logit_opportunity2)
fit_logit_qualified2 <- model_performance(logit_qualified2)
fit_logit_candidate2 <- model_performance(logit_candidate2)

df_fit_logit_opportunity2 <- as.data.frame(fit_logit_opportunity2)
df_fit_logit_qualified2 <- as.data.frame(fit_logit_qualified2)
df_fit_logit_candidate2 <- as.data.frame(fit_logit_candidate2)

df_fit_all2 <- bind_rows(

  df_fit_logit_opportunity2 %>% mutate(model = "Logit - Rejection after Opportunity"),
  df_fit_logit_qualified2 %>% mutate(model = "Logit - Rejection after Qualified Opportunity"),
  df_fit_logit_candidate2 %>% mutate(model = "Logit - Rejection after Candidate")
)


# Updated data_list including all new model outputs
data_list <- list(

  "Logit Models Updated" = df_logit_all2,
  "Fit Statistics All Models" = df_fit_all2
)

# Save all to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_NewModels2.xlsx")



library(dplyr)

correlation_vars <- dplyr::select(
  df_cleaned_nooutliers,
  Ticket.size, Round.Size, Distance.to.Vienna, dplyr::starts_with("Origination.of.the.deal_")
)


colnames(df_cleaned_nooutliers)

# Compute Pearson and Spearman correlation matrices
cor_pearson <- cor(correlation_vars, use = "pairwise.complete.obs", method = "pearson")
cor_spearman <- cor(correlation_vars, use = "pairwise.complete.obs", method = "spearman")

# Optional: round to 3 decimals for clarity
cor_pearson <- round(cor_pearson, 3)
cor_spearman <- round(cor_spearman, 3)

# Convert to data frames and keep rownames as a column
cor_pearson_df <- as.data.frame(cor_pearson)
cor_pearson_df$Variable <- rownames(cor_pearson_df)
cor_pearson_df <- cor_pearson_df[, c("Variable", setdiff(names(cor_pearson_df), "Variable"))]

cor_spearman_df <- as.data.frame(cor_spearman)
cor_spearman_df$Variable <- rownames(cor_spearman_df)
cor_spearman_df <- cor_spearman_df[, c("Variable", setdiff(names(cor_spearman_df), "Variable"))]

# Export both to Excel
save_apa_formatted_excel(
  list(
    "Correlation Matrix - Pearson" = cor_pearson_df,
    "Correlation Matrix - Spearman" = cor_spearman_df
  ),
  "APA_Formatted_Tables_Correlations_Matrices.xlsx"
)




correlation_vars <- df_cleaned_nooutliers %>%
  dplyr::select(
    Ticket.size, Round.Size, Distance.to.Vienna,
    starts_with("Origination.of.the.deal_"),
    Out_of_scope, Too_early, No_VCs_on_cap_table, Product, 
    Not_growing, Broken_cap_table, Too_expensive, Geography, 
    Business_Model, Timing, Total_addressable_market, Team, 
    Direct_competitor, Too_late_stage
  )


library(Hmisc)
library(openxlsx)

# Compute correlations with significance testing
correlation_test <- rcorr(as.matrix(correlation_vars), type = "spearman")

# Extract correlation coefficients and p-values
correlation_coeff <- round(correlation_test$r, 3)
correlation_pvals <- correlation_test$P

# Flag significant correlations (e.g., p < .05)
significance_flags <- ifelse(correlation_pvals < 0.001, "***",
                             ifelse(correlation_pvals < 0.01, "**",
                                    ifelse(correlation_pvals < 0.05, "*", "")))

# Combine coefficients and significance flags
correlation_with_flags <- matrix(paste0(correlation_coeff, significance_flags), 
                                 nrow = nrow(correlation_coeff),
                                 dimnames = dimnames(correlation_coeff))

# Convert to data frame
correlation_with_flags_df <- as.data.frame(correlation_with_flags)
correlation_with_flags_df$Variable <- rownames(correlation_with_flags_df)
correlation_with_flags_df <- correlation_with_flags_df[, c("Variable", setdiff(names(correlation_with_flags_df), "Variable"))]

# Export to Excel
write.xlsx(correlation_with_flags_df, "Correlation_Matrix_Flagged2.xlsx")
