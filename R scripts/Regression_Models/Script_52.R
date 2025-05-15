# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/s_deutz")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Data_Clean.xlsx")


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

# Clean variable names and values

library(dplyr)

# Rename specific columns
df <- df %>%
  rename(
    EverNegotiated = v_503,
    ExperienceinNegotiation = v_315,
    EverWorkedEnergyIndustry = Additional.Demographics_v_674,
    RealSalaryConsideration = Additional.Demographics_v_675
  )

# Recode the 'occupationtype' variable
df$occupationtype <- factor(df$occupationtype,
                            levels = c(1, 2, 3, 4, 5, 6, 7),
                            labels = c("Practical activity; dealing with things and tools",
                                       "Researching, observing activity; scientific handling of data",
                                       "Creative, artistic activity",
                                       "Interactive activity; advising, educating, healing people",
                                       "Entrepreneurial, commercial activity",
                                       "Organizing, administrative activity",
                                       "Other")
)

# Confirm changes
colnames(df)

# Remove rows where 'Consent' is NA
df2 <- df[!is.na(df$Consent), ]
# Remove rows where 'Salary.Negotiation.Scenario_v_633' is NA
df3 <- df[!is.na(df$`Salary.Negotiation.Scenario_v_633`), ]


# Convert selected variables to character type
df3$Language <- as.character(df3$Language)
df3$Gender <- as.character(df3$Gender)
df3$job <- as.character(df3$job)
df3$EverNegotiated <- as.character(df3$EverNegotiated)
df3$ExperienceinNegotiation <- as.character(df3$ExperienceinNegotiation)
df3$ComprehensionCheck <- as.character(df3$ComprehensionCheck)
df3$EverWorkedEnergyIndustry <- as.character(df3$EverWorkedEnergyIndustry)
df3$RealSalaryConsideration <- as.character(df3$RealSalaryConsideration)
df3$Final.Question_SelfExclusion <- as.character(df3$Final.Question_SelfExclusion)

colnames(df3)

# Filter the DataFrame to keep only rows where 'ComprehensionCheck' equals "1" and is not NA
df4 <- df3[!is.na(df3$ComprehensionCheck) & df3$ComprehensionCheck == "1", ]

# Filter the DataFrame to keep only rows where 'Final.Question_SelfExclusion' equals "1" and is not NA
df5 <- df4[df4$Final.Question_SelfExclusion == "1" & !is.na(df4$Final.Question_SelfExclusion), ]

# Optionally, check the structure of the DataFrame to confirm the changes
str(df)


# Recode 'Gender'
df5$Gender <- factor(df5$Gender,
                    levels = c("1", "2", "3"),
                    labels = c("Male", "Female", "Diverse"))

# Recode 'job'
df5$job <- factor(df5$job,
                 levels = c("1", "2", "3", "4"),
                 labels = c("School", "University", "Employed", "Other"))

# Recode 'EverNegotiated'
df5$EverNegotiated <- factor(df5$EverNegotiated,
                            levels = c("1", "2"),
                            labels = c("Yes", "No"))

# Recode 'EverWorkedEnergyIndustry'
df5$EverWorkedEnergyIndustry <- factor(df5$EverWorkedEnergyIndustry,
                                      levels = c("1", "2"),
                                      labels = c("Yes", "No"))

# Recode 'Language'
df5$Language <- factor(df5$Language,
                      levels = c("1", "2"),
                      labels = c("Yes", "No"))


scales_SocioD_char <- c("Language", "Gender", "job", "EverNegotiated", "ExperienceinNegotiation", "occupationtype", "EverWorkedEnergyIndustry", "RealSalaryConsideration")
scales_SocioD_num <- c("Age", "working.hours", "Salary.Negotiation.Scenario_v_633", "Actual.Request_v_482")
scales_Big5 <- c("Big5_v_623", "Big5_v_624", "Big5_v_625", "Big5_v_626", "Big5_v_627", "Big5_v_628", "Big5_v_629", "Big5_v_630", "Big5_v_631", "Big5_v_632")
scales_Forcing_Yielding <- c("Forcing_Yielding_v_648", "Forcing_Yielding_v_649", "Forcing_Yielding_v_650", "Forcing_Yielding_v_651", "Forcing_Yielding_v_652", "Forcing_Yielding_v_653", "Forcing_Yielding_v_654", "Forcing_Yielding_v_655")
scales_Masculinity_Implications <- c("Masculinity.Implications_v_662", "Masculinity.Implications_v_663", "Masculinity.Implications_v_664", "Masculinity.Implications_v_665")

# Coalesce the gender-specific GRDS_GRD columns into unified columns
df5$GRDS_GRD_1 <- coalesce(df5$GRDS_GRD_Men_v_583, df5$Women_GRDS_GRD_v_613)
df5$GRDS_GRD_2 <- coalesce(df5$GRDS_GRD_Men_v_584, df5$Women_GRDS_GRD_v_614)
df5$GRDS_GRD_3 <- coalesce(df5$GRDS_GRD_Men_v_585, df5$Women_GRDS_GRD_v_615)
df5$GRDS_GRD_4 <- coalesce(df5$GRDS_GRD_Men_v_586, df5$Women_GRDS_GRD_v_616)
df5$GRDS_GRD_5 <- coalesce(df5$GRDS_GRD_Men_v_587, df5$Women_GRDS_GRD_v_617)
df5$GRDS_GRD_6 <- coalesce(df5$GRDS_GRD_Men_v_588, df5$Women_GRDS_GRD_v_618)
df5$GRDS_GRD_7 <- coalesce(df5$GRDS_GRD_Men_v_589, df5$Women_GRDS_GRD_v_619)
df5$GRDS_GRD_8 <- coalesce(df5$GRDS_GRD_Men_v_590, df5$Women_GRDS_GRD_v_620)
df5$GRDS_GRD_9 <- coalesce(df5$GRDS_GRD_Men_v_591, df5$Women_GRDS_GRD_v_621)
df5$GRDS_GRD_10 <- coalesce(df5$GRDS_GRD_Men_v_592, df5$Women_GRDS_GRD_v_622)

# Create a list of the integrated GRDS_GRD variables
scales_grds_grd <- c("GRDS_GRD_1", "GRDS_GRD_2", "GRDS_GRD_3", "GRDS_GRD_4", 
                   "GRDS_GRD_5", "GRDS_GRD_6", "GRDS_GRD_7", "GRDS_GRD_8", 
                   "GRDS_GRD_9", "GRDS_GRD_10")


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

df_freq <- create_frequency_tables(df5, scales_SocioD_char)

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

# Example usage with your dataframe and list of variables
df_descriptive_sociod <- calculate_descriptive_stats(df5, scales_SocioD_num)


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
    ITC = numeric(),  # Item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- psych::alpha(subset_data)  # Ensure 'psych' package is loaded
    alpha_val <- alpha_results$total$raw_alpha
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"]  # ITC for the item
      item_mean <- mean(item_data, na.rm = TRUE)
      item_sem <- sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data)))
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = item_mean,
        SEM = item_sem,
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
        Alpha = NA  # Alpha is not calculated per item
      ))
    }
    
    # Calculate the total score for the scale as the sum of items
    scale_sum <- rowSums(subset_data, na.rm = TRUE)
    data[[paste0(scale, "_Total")]] <- scale_sum  # Store scale total in the dataframe
    scale_mean_overall <- mean(scale_sum, na.rm = TRUE)
    scale_sem <- sd(scale_sum, na.rm = TRUE) / sqrt(sum(!is.na(scale_sum)))
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = paste0(scale, "_Total"),
      Mean = scale_mean_overall,
      SEM = scale_sem,
      StDev = sd(scale_sum, na.rm = TRUE),
      ITC = NA,  # ITC is not applicable for the total scale
      Alpha = alpha_val  # Scale-level alpha
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}

colnames(df5)

scales <- list(
  "GRDS" = c("GRDS_GRD_1", "GRDS_GRD_2", "GRDS_GRD_3", "GRDS_GRD_4", "GRDS_GRD_5"),
  "GRD" = c("GRDS_GRD_6", "GRDS_GRD_7", "GRDS_GRD_8", "GRDS_GRD_9", "GRDS_GRD_10"),
  "Forcing" = c("Forcing_Yielding_v_648", "Forcing_Yielding_v_649", 
                "Forcing_Yielding_v_650", "Forcing_Yielding_v_651"),
  "Yielding" = c("Forcing_Yielding_v_652", "Forcing_Yielding_v_653", 
                 "Forcing_Yielding_v_654", "Forcing_Yielding_v_655"),
  "Masculinity_Implications" = c("Masculinity.Implications_v_662", "Masculinity.Implications_v_663", 
                                 "Masculinity.Implications_v_664", "Masculinity.Implications_v_665")
)


alpha_results <- reliability_analysis(df5, scales)

df5_recoded <- alpha_results$data_with_scales
df5_reliability_scales <- alpha_results$statistics


colnames(df5_recoded)

## DESCRIPTIVE STATS BY FACTORS - FACTORS ARE DIFFERENT COLUMNS

calculate_means_and_sds_by_factors <- function(data, variables, factors) {
  # Create an empty list to store intermediate results
  results_list <- list()
  
  # Iterate over each variable
  for (var in variables) {
    # Create a temporary data frame to store results for this variable
    temp_results <- data.frame(Variable = var)
    
    # Iterate over each factor
    for (factor in factors) {
      # Aggregate data by factor
      agg_data <- aggregate(data[[var]], by = list(data[[factor]]), 
                            FUN = function(x) c(Mean = mean(x, na.rm = TRUE), SD = sd(x, na.rm = TRUE)))
      
      # Create columns for each level of the factor
      for (level in unique(data[[factor]])) {
        level_agg_data <- agg_data[agg_data[, 1] == level, ]
        
        if (nrow(level_agg_data) > 0) {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- level_agg_data$x[1, "Mean"]
          temp_results[[paste0(factor, "_", level, "_SD")]] <- level_agg_data$x[1, "SD"]
        } else {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- NA
          temp_results[[paste0(factor, "_", level, "_SD")]] <- NA
        }
      }
    }
    
    # Add the results for this variable to the list
    results_list[[var]] <- temp_results
  }
  
  # Combine all the results into a single dataframe
  descriptive_stats_bygroup <- do.call(rbind, results_list)
  return(descriptive_stats_bygroup)
}

# Assuming df5_recoded is your DataFrame
df5_recoded <- df5_recoded %>%
  rename(
    GRDS = GRDS_Total,
    GRD = GRD_Total,
    Forcing = Forcing_Total,
    Yielding = Yielding_Total,
    Masculinity_Implications = Masculinity_Implications_Total
  )

scales_ivs <- c(
  "GRDS",
  "GRD",
  "Forcing", 
          
  "Yielding", 
          
  "Masculinity_Implications"
)

factors <- scales_SocioD_char

df5_descriptives_scales_bygroup <- calculate_means_and_sds_by_factors(df5_recoded, scales_ivs, factors)

df5_descriptives_scales <- calculate_descriptive_stats(df5_recoded, scales_ivs)


## BOXPLOTS DIVIDED BY FACTOR

library(ggplot2)

# Specify the factor as a column name
factor_var <- "Gender"  # replace with your actual factor column name
vars <- scales_ivs

# Loop for generating and saving boxplots
for (var in vars) {
  # Create boxplot for each variable using ggplot2
  p <- ggplot(df5_recoded, aes_string(x = factor_var, y = var)) +
    geom_boxplot() +
    stat_summary(fun.y = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = paste("Boxplot of", var, "by", factor_var),
         x = factor_var, y = var)
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = paste("boxplot_", var, ".png", sep = ""), plot = p, width = 10, height = 6)
}


# outlier examination

calculate_z_scores <- function(data, vars, id_var, z_threshold = 3) {
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~ (.-mean(.))/sd(.), .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  return(z_score_data)
}

id_var <- "lfdn"
df_zscores <- calculate_z_scores(df5_recoded, scales_ivs, id_var)

# Normality Assessment
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
df_normality_results <- calculate_stats(df5_recoded, scales_ivs)

colnames(df5_recoded)


# Hypotheses Testing

library(broom)
library(boot)

# Function to perform bootstrapping on linear model coefficients
bootstrap_lm <- function(data, formula, R = 1000) {
  boot_fn <- function(data, indices) {
    # Refit the model on the bootstrapped data
    fit <- lm(formula, data = data[indices, ])
    # Return the vector of all coefficients
    return(coef(fit))
  }
  # Perform the bootstrap
  boot(data = data, statistic = boot_fn, R = R)
}


# Main function to fit models and summarize results
fit_ols_and_format <- function(data, predictors, response_vars, gender, R = 1000, save_plots = FALSE) {
  results_list <- list()
  data_subset <- data[data$Gender == gender, ]
  
  for (response_var in response_vars) {
    formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Bootstrapping the model
    boot_results <- bootstrap_lm(data_subset, formula, R)
    
    # Fit the model normally for additional statistics
    model_fit <- lm(formula, data = data_subset)
    model_summary <- summary(model_fit)
    tidy_results <- broom::tidy(model_fit)
    
    # Initialize CI columns
    tidy_results$CI_Lower <- rep(NA, nrow(tidy_results))
    tidy_results$CI_Upper <- rep(NA, nrow(tidy_results))
    
    # Loop over terms in the model and extract confidence intervals for each term
    for (i in seq_len(nrow(tidy_results))) {
      term_name <- tidy_results$term[i]
      
      # Ensure we are extracting confidence intervals for each term separately
      if (term_name %in% names(boot_results$t0)) {
        term_index <- which(names(boot_results$t0) == term_name)
        
        # Try extracting confidence intervals for each coefficient
        tryCatch({
          boot_cis <- boot::boot.ci(boot_results, index = term_index, type = "basic")
          tidy_results$CI_Lower[i] <- boot_cis$basic[4]  # Lower CI
          tidy_results$CI_Upper[i] <- boot_cis$basic[5]  # Upper CI
        }, error = function(e) {
          message("Could not compute CIs for term: ", term_name)
        })
      }
    }
    
    tidy_results$R_Squared <- rep(model_summary$r.squared, nrow(tidy_results))
    
    if (save_plots) {
      plot_path <- paste0(gsub(" ", "_", response_var), "_Diagnostic_Plots.jpeg")
      jpeg(plot_path)
      par(mfrow = c(2, 2))
      plot(model_fit)
      dev.off()
    }
    
    results_list[[response_var]] <- tidy_results
  }
  
  return(results_list)
}


# Define predictors and response variables
predictors <- c("GRDS", "GRD")
response_vars <- c("Salary.Negotiation.Scenario_v_633", "Actual.Request_v_482", "Forcing", "Yielding", "Masculinity_Implications")

# Fit models for men and print results
df_results_men <- fit_ols_and_format(df5_recoded, predictors, response_vars, gender = "Male", R = 1000, save_plots = TRUE)
print(df_results_men)


# Fit models for women
df_results_women <- fit_ols_and_format(df5_recoded, predictors, response_vars, gender = "Female", R = 1000, save_plots = TRUE)

concatenate_results <- function(results_list) {
  # Use dplyr::bind_rows to concatenate all the data frames in the list
  combined_df <- dplyr::bind_rows(results_list, .id = "ResponseVariable")
  
  return(combined_df)
}

# Concatenate results for men
df_combined_men_df <- concatenate_results(df_results_men)

# Concatenate results for women
df_combined_women_df <- concatenate_results(df_results_women)

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
  "Final DF" = df5_recoded,
  "Sample Description" = df_freq,
  "Sample Description 2" = df_descriptive_sociod,
  "Reliability Results" = df5_reliability_scales,
  "Descriptives" = df5_descriptives_scales,
  "Descriptives by Group" = df5_descriptives_scales_bygroup,
  "Normality Tests" = df_normality_results,
  "Model Results - Men" = df_combined_men_df,
  "Model Results - Women" = df_combined_women_df
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")




library(dplyr)
library(openxlsx)

# Filter dataset for Gender = Male
df_male <- df5_recoded %>%
  filter(Gender == "Male")

# Identify numeric and factor/character variables
numeric_vars <- names(df_male)[sapply(df_male, is.numeric)]
factor_vars <- names(df_male)[sapply(df_male, function(x) is.factor(x) || is.character(x))]

# Calculate descriptive statistics for numeric variables
numeric_summary <- calculate_descriptive_stats(df_male, numeric_vars)

# Calculate frequency tables for factor variables
factor_summary <- create_frequency_tables(df_male, factor_vars)

# Create a new Excel workbook
wb <- createWorkbook()

# Add a sheet for numeric summary
addWorksheet(wb, "Numeric Summary")
writeData(wb, "Numeric Summary", numeric_summary)

# Add a sheet for factor summary
addWorksheet(wb, "Factor Summary")
writeDataTable(wb, "Factor Summary", factor_summary, tableStyle = "TableStyleMedium2")

# Save the workbook
saveWorkbook(wb, "summary_output.xlsx", overwrite = TRUE)

# Print completion message
cat("The data has been successfully written to 'summary_output.xlsx'")









library(dplyr)
library(openxlsx)

# Filter dataset for Gender = Male
df_male <- df5_recoded %>%
  filter(Gender == "Female")

# Identify numeric and factor/character variables
numeric_vars <- names(df_male)[sapply(df_male, is.numeric)]
factor_vars <- names(df_male)[sapply(df_male, function(x) is.factor(x) || is.character(x))]

# Calculate descriptive statistics for numeric variables
numeric_summary <- calculate_descriptive_stats(df_male, numeric_vars)

# Calculate frequency tables for factor variables
factor_summary <- create_frequency_tables(df_male, factor_vars)

# Create a new Excel workbook
wb <- createWorkbook()

# Add a sheet for numeric summary
addWorksheet(wb, "Numeric Summary")
writeData(wb, "Numeric Summary", numeric_summary)

# Add a sheet for factor summary
addWorksheet(wb, "Factor Summary")
writeDataTable(wb, "Factor Summary", factor_summary, tableStyle = "TableStyleMedium2")

# Save the workbook
saveWorkbook(wb, "summary_output_women.xlsx", overwrite = TRUE)

# Print completion message
cat("The data has been successfully written to 'summary_output_women.xlsx'")

