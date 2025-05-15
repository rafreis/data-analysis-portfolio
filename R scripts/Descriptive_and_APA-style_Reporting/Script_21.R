# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dan_sleep")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Clean_Data.xlsx")

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

# Arrange and group by Q27_Cleaned and Questionnaire, then keep the first entry for each combination
df <- df %>%
  group_by(Q27_Cleaned, Questionnaire) %>%
  slice(1) %>%  # Keep only the first row for each group
  ungroup()


# Defining Scales

scale_quality <- c("Q2_1.score", "Q2_2.score", "Q2_3.score", "Q2_4.score", "Q2_6.score", "Q2_7.score", "Q2_8.score", "Q5_1.score")

scale_bedtime_stress_anxiety <- c("Q2_9.score", "Q2_10.score", "Q3_1.score", "Q3_2.score", "Q3_3.score", "Q3_4.score", "Q3_5.score", "Q3_6.score", "Q3_7.score", "Q3_8.score", "Q3_9.score")

scale_consistency <- c("Q2_5.score", "Q4_1.score", "Q4_2.score", "Q4_3.score", "Q4_4.score", "Q4_5.score", "Q4_6.score", "Q4_7.score")

scale_next_day_alertness <- c("Q23_1.score", "Q23_2.score", "Q23_3.score", "Q23_4.score", "Q23_5.score", "Q23_6.score", "Q23_7.score", "Q23_8.score", "Q23_9.score", "Q24_1.score", "Q24_2.score", "Q24_3.score", "Q24_4.score", "Q24_5.score", "Q25_1.score", "Q25_2.score")

scale_sociod_cat <- c("Questionnaire", "Choice")

scale_sociod_catunique <- c("School", "Gender", "Age", "Year", "Test.Group")

# Filter the dataframe for unique entries of Q27_Cleaned
df_unique_q27 <- df %>%
  group_by(Q27_Cleaned) %>%
  sample_n(1) %>%
  ungroup()

# Calculate the mean and standard deviation of the Age column in the unique dataframe
mean_age_unique <- mean(df_unique_q27$Age, na.rm = TRUE)
sd_age_unique <- sd(df_unique_q27$Age, na.rm = TRUE)



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


df_freq <- create_frequency_tables(df, scale_sociod_cat)

# Create the frequency tables on the unique dataframe
df_freq_unique <- create_frequency_tables(df_unique_q27, scale_sociod_catunique)



## RELIABILITY ANALYSIS

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

scales <- list(
  
  "Quality" = c("Q2_1.score", "Q2_2.score", "Q2_3.score", "Q2_4.score", "Q2_6.score", "Q2_7.score", "Q2_8.score", "Q5_1.score"),
  "Bedtime_stress_anxiety" = c("Q2_9.score", "Q2_10.score", "Q3_1.score", "Q3_2.score", "Q3_3.score", "Q3_4.score", "Q3_5.score", "Q3_6.score", "Q3_7.score", "Q3_8.score", "Q3_9.score"),
  "Consistency" = c("Q2_5.score", "Q4_1.score", "Q4_2.score", "Q4_3.score", "Q4_4.score", "Q4_5.score", "Q4_6.score", "Q4_7.score"),
  "Next_Day_Alertness" = c("Q23_1.score", "Q23_2.score", "Q23_3.score", "Q23_4.score", "Q23_5.score", "Q23_6.score", "Q23_7.score", "Q23_8.score", "Q23_9.score", "Q24_1.score", "Q24_2.score", "Q24_3.score", "Q24_4.score", "Q24_5.score", "Q25_1.score", "Q25_2.score")
)


alpha_results <- reliability_analysis(df, scales)

df_recoded <- alpha_results$data_with_scales
df_descriptives <- alpha_results$statistics


# Remove rows where Test.Group is "Unknown"
df_recoded <- subset(df_recoded, Test.Group != "Unknown")
# Remove rows where Test.Group is "unknown"
df_recoded <- subset(df_recoded, Test.Group != "unknown")

scales_DV <- c("Quality","Bedtime_stress_anxiety", "Consistency",  "Next_Day_Alertness")

scales_Group <- "Test.Group"

scales_Time <- "Questionnaire"


#Normality Assessment

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
df_normality_results <- calculate_stats(df_recoded, scales_DV)

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    Median = numeric(),
    SEM = numeric(),
    SD = numeric(),
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  for (var in desc_vars) {
    print(paste("Processing variable:", var))  # Debug print
    variable_data <- data[[var]]
    if (length(variable_data) == 0) {
      warning(paste("Variable", var, "is empty or does not exist. Skipping."))
      next
    }
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

# Example usage with your dataframe and list of variables
df_descriptive_stats <- calculate_descriptive_stats(df_recoded, scales_DV)

# Visualization

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
    theme(axis.text.x = element_text(hjust = 0.5, vjust = 1)) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 10))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = "side_by_side_boxplots.png", plot = p, width = 10, height = 6)
}

create_boxplots(df_recoded, scales_DV)

str(df_recoded)


library(ggplot2)
library(tidyr)
library(dplyr)

plot_pre_post_boxplots <- function(df, scales_DV, time_var, group_var) {
  # Reshape the data to long format for selected scales
  long_df <- df %>%
    pivot_longer(
      cols = scales_DV,
      names_to = "Variable",
      values_to = "Value"
    ) %>%
    select(!!rlang::sym(time_var), !!rlang::sym(group_var), Variable, Value)  # Dynamically select time and group variables
  
  # Generate a list of plots, one for each variable
  plots <- lapply(unique(long_df$Variable), function(var) {
    p <- ggplot(data = filter(long_df, Variable == var), aes(x = .data[[time_var]], y = Value, fill = .data[[group_var]])) +
      geom_boxplot(position = position_dodge(width = 0.8)) +
      stat_summary(
        fun = mean, geom = "point", shape = 16, size = 3, color = "red",
        position = position_dodge(width = 0.8)  # Ensure it aligns with the boxplots
      ) +
      labs(title = paste("Boxplot of", var, "across different time points"),
           x = "Time Point",
           y = "Score") +
      theme_minimal() +
      scale_fill_brewer(palette = "Dark2")  # Choosing a color palette that provides good contrast
    
    # Construct a filename using the variable name
    filename <- paste0("Boxplot_", gsub("[^A-Za-z0-9]", "", var), ".png")  # Remove non-alphanumeric characters for the filename
    ggsave(filename, plot = p, width = 10, height = 6, units = "in")  # Save the plot
    p
  })
  
  return(plots)
}

# Example usage
# Specify your group and time variables as well as the scales to plot
time_var = "Questionnaire"  # Adjust as needed
group_var = "Test.Group"  # Adjust as needed

# Generate the plots
plots <- plot_pre_post_boxplots(df_recoded, scales_DV, time_var, group_var)

lapply(plots, print)  

library(lme4)
library(broom.mixed)
library(afex)

# Set reference levels for the factors in the data
df_recoded$Questionnaire <- relevel(as.factor(df_recoded$Questionnaire), ref = "questionnaire_1")
df_recoded$Test.Group <- relevel(as.factor(df_recoded$Test.Group), ref = "group 4 (wait list)")

# Function to fit LMM for multiple response variables and return formatted results
fit_lmm_and_format <- function(data, within_subject_var, between_subject_var, random_effects, response_vars, save_plots = FALSE) {
  lmm_results_list <- list()  # Initialize an empty list to store results for each response variable
  
  for (response_var in response_vars) {
    # Construct the formula dynamically for each response variable
    formula <- as.formula(paste(response_var, "~", within_subject_var, "*", between_subject_var, "+", random_effects))
    
    # Fit the linear mixed-effects model
    lmer_model <- lmer(formula, data = data)
    
    # Print the full summary of the model for fit statistics
    summary_output <- summary(lmer_model)
    print(summary_output)  # Print full model summary to the console
    
    # Extract the tidy output for fixed effects
    fixed_effects_results <- tidy(lmer_model, effects = "fixed")
    # Extract the tidy output for random effects
    random_effects_results <- tidy(lmer_model, effects = "ran_pars")
    
    # Optionally print the tidy outputs for debugging or review
    print(fixed_effects_results)
    print(random_effects_results)
    
    # Generate and save residual plots if requested
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)  # Clean up variable name for plot file
      
      # Residuals vs Fitted plot
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      # QQ Plot for residuals
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lmer_model))
      qqline(residuals(lmer_model))
      dev.off()
    }
    
    # Combine fixed and random effects results into one dataframe
    results_combined <- bind_rows(
      fixed_effects_results %>% mutate(Effect = "Fixed"),
      random_effects_results %>% mutate(Effect = "Random")
    )
    
    # Store the combined results for the current response variable in a list
    lmm_results_list[[response_var]] <- results_combined
  }
  
  return(lmm_results_list)  # Return the list of combined results for each response variable
}



# Example usage:
response_vars <- scales_DV  # Define the response variables you want to loop through
lmm_results <- fit_lmm_and_format(
  data = df_recoded,
  within_subject_var = "Questionnaire",
  between_subject_var = "Test.Group",
  random_effects = "(1|Q27_Cleaned) + (1|School)",
  response_vars = response_vars,
  save_plots = TRUE
)


model_quality <- lmm_results$Quality
model_bedtime <- lmm_results$Bedtime_stress_anxiety
model_consistency <- lmm_results$Consistency
model_nextday <- lmm_results$Next_Day_Alertness

# Load necessary libraries
library(nlme)
library(broom.mixed)

# Function to fit a Multivariate LMM including within-subject and between-subject factors
fit_multivariate_lmm_and_format <- function(data, response_vars, within_subject_var, between_subject_var, random_effects_spec, random_effects_vars, save_plots = FALSE) {
  
  # Ensure the random effects variables are in the dataframe and are factors
  if (!all(random_effects_vars %in% names(data))) {
    stop("One or more random effects variables are not found in the dataframe: ", paste(random_effects_vars[!random_effects_vars %in% names(data)], collapse=", "))
  }
  data[, random_effects_vars] <- lapply(data[, random_effects_vars, drop = FALSE], factor)
  
  # Construct the multivariate response formula
  response_formula <- paste(response_vars, collapse = ", ")
  full_formula <- as.formula(paste("cbind(", response_formula, ") ~", within_subject_var, "*", between_subject_var))
  
  # Print the model formula for debugging
  cat("Model formula:\n")
  print(full_formula)
  
  # Print the structure of the data, focusing on random effects variables
  cat("Structure of data (focus on random effects variables):\n")
  print(summary(data[random_effects_vars]))
  
  # Fit the multivariate linear mixed-effects model with random effects
  mlmer_model <- try(
    lme(
      full_formula, 
      data = data, 
      random = as.formula(random_effects_spec),  # Flexible random effects specification
      na.action = na.exclude, 
      method = "REML"
    ), 
    silent = FALSE
  )
  
  # Handle model fitting failure
  if (inherits(mlmer_model, "try-error")) {
    cat("Model fitting failed with an error:\n")
    print(mlmer_model)
    return(NULL)
  }
  
  # Print the summary of the model for fit statistics
  summary_output <- summary(mlmer_model)
  print(summary_output)
  
  # Extract and return tidy output for fixed effects coefficients
  tidy_output <- tidy(mlmer_model, effects = "fixed")
  print(tidy_output)
  
  # Optionally generate and save residual plots
  if (save_plots) {
    plot_name_prefix <- "Multivariate_Model"
    
    jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
    plot(residuals(mlmer_model) ~ fitted(mlmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
    dev.off()
    
    jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
    qqnorm(residuals(mlmer_model))
    qqline(residuals(mlmer_model))
    dev.off()
  }
  
  # Return both summary output and tidy dataframe
  return(list(
    summary_stats = summary_output,
    tidy_coefficients = tidy_output
  ))
}

# Example usage:
response_vars <- c("Quality", "Bedtime_stress_anxiety", "Consistency", "Next_Day_Alertness")
random_effects_vars <- c("Q27_Cleaned", "School")  # Specify actual random effects column names
random_effects_spec <- "~1 | Q27_Cleaned/School"  # Flexible random effects structure

# Call the function with the relevant parameters
multivariate_results <- fit_multivariate_lmm_and_format(
  data = df_recoded,
  response_vars = response_vars,
  within_subject_var = "Questionnaire",
  between_subject_var = "Test.Group",
  random_effects_spec = random_effects_spec,
  random_effects_vars = random_effects_vars,
  save_plots = TRUE
)

# Access the summary statistics and tidy output
summary_output <- multivariate_results$summary_stats
model_multivariate_coefficients <- multivariate_results$tidy_coefficients


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
  "Frequency Table Unique" = df_freq_unique, 
  "Descriptive Stats" = df_descriptive_stats, 
  "Reliability" = df_descriptives,
  "Model - Quality" = model_quality,
  "Model - Consist" = model_consistency,
  "Model - NextDay" = model_nextday,
  "Model - BedTime" = model_bedtime,
  "Multivariate Model" = model_multivariate_coefficients
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")

## DESCRIPTIVE STATS BY FACTORS - TWO FACTORS AS LEVELS OF COLUMNS

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


descriptive_stats_bygroup <- calculate_descriptive_stats_bygroups(df_recoded, scales_DV, "Test.Group", "Questionnaire")

# Example usage
data_list <- list(
  
  "Descriptive Stats by Groups" = descriptive_stats_bygroup
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Descriptives_groups.xlsx")



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
      
      # Add the variable name and calculate percentages
      freq_table$Variable <- var
      freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


df_freq_segmented <- create_segmented_frequency_tables(df, "Test.Group", "Questionnaire")

colnames(df)
