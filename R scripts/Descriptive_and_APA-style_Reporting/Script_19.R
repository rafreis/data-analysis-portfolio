# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/danieladibas")

# Read data from a CSV file
df <- read.csv("analysis_dataset.csv")

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



# Trim all values and convert to lowercase

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
    # Convert character strings to lowercase
    x <- tolower(x)
  }
  return(x)
}))


colnames(df)

withinsubj_factor <- "period"
betweensubj_factor <- "Intervention_Control"
dvs <- c("RCADS", "YPCORE")

# Convert 'period' column to a factor with a specified order
df$period <- factor(df$period, levels = c("baseline", "midpoint", "end of treatment"), ordered = TRUE)

# Convert 'Intervention_Control' column to a factor (assuming no specific order needed)
df$Intervention_Control <- factor(df$Intervention_Control)


library(dplyr)

# Assuming 'df' is your original dataframe
df_unique <- df %>%
  group_by(ID) %>%
  slice(1) %>%
ungroup()

library(tidyr)

# Function to diagnose missing data
diagnose_missing_data <- function(df, vars) {
  missing_summary <- df %>%
    select(all_of(vars)) %>%
    summarise(across(everything(), ~ sum(is.na(.)))) %>%
    pivot_longer(cols = everything(), names_to = "Variable", values_to = "Missing_Values")
  
  total_rows <- nrow(df)
  
  missing_summary <- missing_summary %>%
    mutate(Percentage_Missing = (Missing_Values / total_rows) * 100)
  
  return(missing_summary)
}

missing_data_summary <- diagnose_missing_data(df, dvs)


# Outlier Evaluation

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
df_zscores <- calculate_z_scores(df, c("RCADS", "YPCORE"), "ID")


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

vars_cat <-  c("Intervention_Control", "Presenting.issue" , "N..of.sessions" , "client_sex_at_birth" )
df_freq <- create_frequency_tables(df_unique, vars_cat)

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

vars_cont <- c("N..of.sessions" , "client_age_when_first_seen_years", "time_between_baseline_follow_up")

df_descriptive_stats <- calculate_descriptive_stats(df_unique, c(vars_cont, dvs))



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

df_desc_onefactor <- calculate_descriptive_stats_bygroups(df, dvs, withinsubj_factor)


## DESCRIPTIVE STATS BY FACTORS - TWO FACTORS AS LEVELS OF COLUMNS

library(dplyr)

calculate_descriptive_stats_by2groups <- function(data, desc_vars, cat_column1, cat_column2) {
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

df_desc_twofactor <- calculate_descriptive_stats_by2groups(df, dvs, betweensubj_factor, withinsubj_factor)

# Interaction Plot

library(ggplot2)
library(dplyr)
library(tidyr)

create_interaction_plot <- function(data, within_subj_factor, between_subj_factor, dvs) {
  # Transform data to long format
  long_data <- data %>%
    pivot_longer(
      cols = all_of(dvs),
      names_to = "Variable",
      values_to = "Value"
    )
  
  # Calculate summary statistics with mean, SD, SEM, and CI
  summary_data <- long_data %>%
    group_by(!!sym(within_subj_factor), !!sym(between_subj_factor), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      N = n(),
      SD = sd(Value, na.rm = TRUE),
      SEM = SD / sqrt(N),
      CI = qt(0.975, df = N-1) * SEM,  # 95% confidence interval
      Lower = Mean - CI,
      Upper = Mean + CI,
      .groups = 'drop'
    )
  
  # Create the plot with mean points and error bars for 95% CI
  p <- ggplot(summary_data, aes(x = !!sym(within_subj_factor), y = Mean, color = !!sym(between_subj_factor))) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = Lower, ymax = Upper), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Interaction Plot: Mean and 95% CI by Group",
      x = paste("Within-Subject Factor: Period"),  # Rename the x-axis title
      y = "Mean Value",
      color = paste("Between-Subject Factor: Group")  # Set the legend title
    )
  
  # Print the plot
  return(p)
}

create_interaction_plot(df, "period", "Intervention_Control", c("RCADS", "YPCORE"))


# Linear Mixed Model

library(lme4)
library(broom.mixed)
library(afex)

# Function to fit LMM for multiple response variables and return formatted results
fit_lmm_and_format <- function(data, within_subject_var, between_subject_var, random_effects, response_vars, save_plots = FALSE) {
  lmm_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    formula <- as.formula(paste(response_var, "~", within_subject_var, "*", between_subject_var, "+", random_effects))
    
    # Fit the linear mixed-effects model
    lmer_model <- lmer(formula, data = data)
    
    # Print the summary of the model for fit statistics
    summary_output <- summary(lmer_model)
    print(summary_output)
    
    # Extract the tidy output for fixed effects
    fixed_effects_results <- tidy(lmer_model, effects = "fixed")
    # Extract the tidy output for random effects
    random_effects_results <- tidy(lmer_model, effects = "ran_pars")
    
    # Optionally print the tidy outputs
    print(fixed_effects_results)
    print(random_effects_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lmer_model))
      qqline(residuals(lmer_model))
      dev.off()
    }
    
    # Concatenate fixed and random effects results into one dataframe
    results_combined <- bind_rows(
      fixed_effects_results %>% mutate(Effect = "Fixed"),
      random_effects_results %>% mutate(Effect = "Random")
    )
    
    # Store the combined results in a list
    lmm_results_list[[response_var]] <- results_combined
  }
  
  return(lmm_results_list)
}

lmm_results <- fit_lmm_and_format(
  data = df,
  within_subject_var = withinsubj_factor,
  between_subject_var = betweensubj_factor,
  random_effects = "(1|ID)",
  response_vars = dvs,
  save_plots = TRUE
)

df_model_RCADS <- lmm_results$RCADS
df_model_YPCORE <- lmm_results$YPCORE

colnames(df_unique)


# Calculate specific percentiles for the age variable
percentiles <- quantile(df_unique$client_age_when_first_seen_years, probs = c(0.33, 0.66))

# Printing the calculated percentiles
print(percentiles)

breaks <- c(-Inf, 14, Inf)  # Define breaks at -Inf, 14, 16, and Inf
labels <- c("Younger than 15",  "15 or Older")  # Labels for the age groups

# Create a new factor variable based on these breaks
df_unique$age_group <- cut(df_unique$client_age_when_first_seen_years, breaks = breaks, labels = labels, right = TRUE)
df$age_group <- cut(df$client_age_when_first_seen_years, breaks = breaks, labels = labels, right = TRUE)

colnames(df)

# Function to fit LMM for multiple response variables and return formatted results
fit_lmm_and_format2 <- function(data, within_subject_var, between_subject_var, between_subject_var2, random_effects, response_vars, save_plots = FALSE) {
  lmm_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically to include two between-subjects factors and their interaction
    formula <- as.formula(paste(response_var, "~", within_subject_var, "*", between_subject_var, "*", between_subject_var2, "+", random_effects))
    
    # Fit the linear mixed-effects model
    lmer_model <- lmer(formula, data = data)
    
    # Print the summary of the model for fit statistics
    summary_output <- summary(lmer_model)
    print(summary_output)
    
    # Extract the tidy output for fixed effects
    fixed_effects_results <- tidy(lmer_model, effects = "fixed")
    # Extract the tidy output for random effects
    random_effects_results <- tidy(lmer_model, effects = "ran_pars")
    
    # Optionally print the tidy outputs
    print(fixed_effects_results)
    print(random_effects_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lmer_model))
      qqline(residuals(lmer_model))
      dev.off()
    }
    
    # Concatenate fixed and random effects results into one dataframe
    results_combined <- bind_rows(
      fixed_effects_results %>% mutate(Effect = "Fixed"),
      random_effects_results %>% mutate(Effect = "Random")
    )
    
    # Store the combined results in a list
    lmm_results_list[[response_var]] <- results_combined
  }
  
  return(lmm_results_list)
}

# Example usage of the function
lmm_results_sex <- fit_lmm_and_format2(
  data = df,
  within_subject_var = withinsubj_factor,
  between_subject_var = betweensubj_factor,
  between_subject_var2 = "client_sex_at_birth",  # add the name of the second between-subject factor
  random_effects = "(1|ID)",
  response_vars = dvs,
  save_plots = TRUE
)

df_model_sex_RCADS <- lmm_results_sex$RCADS
df_model_sex_YPCORE <- lmm_results_sex$YPCORE


lmm_results_age <- fit_lmm_and_format2(
  data = df,
  within_subject_var = withinsubj_factor,
  between_subject_var = betweensubj_factor,
  between_subject_var2 = "age_group",  # add the name of the second between-subject factor
  random_effects = "(1|ID)",
  response_vars = dvs,
  save_plots = TRUE
)

df_model_age_RCADS <- lmm_results_age$RCADS
df_model_age_YPCORE <- lmm_results_age$YPCORE

df <- df %>%
  mutate(Presenting.issue = ifelse(Presenting.issue == "low mood", "anxiety and low mood", Presenting.issue))

df <- df %>%
  mutate(Presenting.issue = ifelse(Presenting.issue == "anxiety and low mood", "anxiety and low mood", Presenting.issue))

lmm_results_issue <- fit_lmm_and_format2(
  data = df,
  within_subject_var = withinsubj_factor,
  between_subject_var = betweensubj_factor,
  between_subject_var2 = "Presenting.issue",  # add the name of the second between-subject factor
  random_effects = "(1|ID)",
  response_vars = dvs,
  save_plots = TRUE
)

df_model_issue_RCADS <- lmm_results_issue$RCADS
df_model_issue_YPCORE <- lmm_results_issue$YPCORE


# Function to fit LMM for multiple response variables and return formatted results
fit_lmm_and_format3 <- function(data, within_subject_var, between_subject_var, numeric_covariate, random_effects, response_vars, save_plots = FALSE) {
  lmm_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically to include interactions between the numeric covariate and categorical factors
    formula <- as.formula(paste(response_var, "~", within_subject_var, "*", between_subject_var, "*", numeric_covariate, "+", random_effects))
    
    # Fit the linear mixed-effects model
    lmer_model <- lmer(formula, data = data)
    
    # Print the summary of the model for fit statistics
    summary_output <- summary(lmer_model)
    print(summary_output)
    
    # Extract the tidy output for fixed effects
    fixed_effects_results <- tidy(lmer_model, effects = "fixed")
    # Extract the tidy output for random effects
    random_effects_results <- tidy(lmer_model, effects = "ran_pars")
    
    # Optionally print the tidy outputs
    print(fixed_effects_results)
    print(random_effects_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lmer_model))
      qqline(residuals(lmer_model))
      dev.off()
    }
    
    # Concatenate fixed and random effects results into one dataframe
    results_combined <- bind_rows(
      fixed_effects_results %>% mutate(Effect = "Fixed"),
      random_effects_results %>% mutate(Effect = "Random")
    )
    
    # Store the combined results in a list
    lmm_results_list[[response_var]] <- results_combined
  }
  
  return(lmm_results_list)
}

# Example usage of the function
lmm_results_sessions <- fit_lmm_and_format3(
  data = df,
  within_subject_var = withinsubj_factor,
  between_subject_var = betweensubj_factor,
  numeric_covariate = "N..of.sessions",  # specify the numeric variable name here
  random_effects = "(1|ID)",
  response_vars = dvs,
  save_plots = TRUE
)

df_model_sessions_RCADS <- lmm_results_sessions$RCADS
df_model_sessions_YPCORE <- lmm_results_sessions$YPCORE

breaks <- c(-Inf, 10, Inf)  # Define breaks at -Inf, 14, 16, and Inf
labels <- c("10 or less",  "More than 10")  # Labels for the age groups
df$sessions_group <- cut(df$N..of.sessions, breaks = breaks, labels = labels, right = TRUE)


print_distinct_values <- function(data, columns) {
  for (column in columns) {
    # Retrieve distinct values
    distinct_values <- distinct(data, .data[[column]])[[column]]
    # Print the results
    cat("Distinct values in", column, ":", toString(distinct_values), "\n")
  }
}

# Call the function
print_distinct_values(df, c("client_sex_at_birth", "age_group", "Presenting.issue", "sessions_group"))


library(dplyr)

# Initialize an empty dataframe to store all results
all_results <- tibble()

# Loop over each distinct level of 'client_sex_at_birth' and run the LMM function
levels_client_sex <- c("female", "male")

for (sex in levels_client_sex) {
  df_subset <- df %>% filter(client_sex_at_birth == sex)
  lmm_results <- fit_lmm_and_format(
    data = df_subset,
    within_subject_var = withinsubj_factor,
    between_subject_var = betweensubj_factor,
    random_effects = "(1|ID)",
    response_vars = c("RCADS", "YPCORE"),
    save_plots = FALSE
  )
  # Label and combine results immediately, then append to the overall dataframe
  temp_results <- label_and_combine_results(lmm_results, c("RCADS", "YPCORE"), sex)
  all_results <- bind_rows(all_results, temp_results)
}

# Factors to iterate through
factors <- list(
  age_group = c("15 or Older", "Younger than 15"),
  Presenting.issue = c("anxiety and low mood", "anxiety"),
  sessions_group = c("More than 10", "10 or less")
)

# Loop through each factor and its levels
for (factor_name in names(factors)) {
  levels <- factors[[factor_name]]
  
  for (level in levels) {
    df_subset <- df %>% filter(!!sym(factor_name) == level)
    lmm_results <- fit_lmm_and_format(
      data = df_subset,
      within_subject_var = withinsubj_factor,
      between_subject_var = betweensubj_factor,
      random_effects = "(1|ID)",
      response_vars = c("RCADS", "YPCORE"),
      save_plots = FALSE
    )
    # Label and combine results immediately, then append to the overall dataframe
    temp_results <- label_and_combine_results(lmm_results, c("RCADS", "YPCORE"), level)
    all_results <- bind_rows(all_results, temp_results)
  }
}

# Now, `all_results` contains the combined results for all groups and factors.
print(all_results)

#Export

## Export Results

# List all objects in the environment
all_objects <- ls()

# Use mget to get the actual objects based on their names
all_objects_list <- mget(all_objects, envir = .GlobalEnv)

# Filter this list to only include objects that are dataframes
dataframe_objects <- Filter(is.data.frame, all_objects_list)

# Get the names of these dataframe objects
dataframe_names <- names(dataframe_objects)

# Print the names
print(dataframe_names)



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

data_list <- list(
  "Frequency Table" = df_freq, 
  "Descriptive Stats" = df_descriptive_stats,
  "Descriptives by Factor" = df_desc_onefactor,
  "Descriptives by Factors" = df_desc_twofactor,
  "Model RCADS" = df_model_RCADS,
  "Model YPCORE" = df_model_YPCORE,
  "Model Age RCADS" = df_model_age_RCADS,
  "Model Age YPCORE" = df_model_age_YPCORE,
  "Model Issue RCADS" = df_model_issue_RCADS,
  "Model Issue YPCORE" = df_model_issue_YPCORE,
  "Model Sessions RCADS" = df_model_sessions_RCADS,
  "Model Sessions YPCORE" = df_model_sessions_YPCORE,
  "Model Sex RCADS" = df_model_sex_RCADS,
  "Model Sex YPCORE" = df_model_sex_YPCORE,
  "Segmented Models" = all_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")
