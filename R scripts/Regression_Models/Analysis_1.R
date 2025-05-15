# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/amandaplhc")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Cleaned_Data.xlsx")

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
str(df)


# Load necessary libraries
library(tidyverse)
library(lubridate)
library(ggplot2)
library(nlme) # Use nlme::lme for p-values in mixed models

# Convert Date from numeric to actual date format
df$Date <- as.Date(df$Date, origin = "1899-12-30")

df <- df %>%
  group_by(Patients) %>%
  filter(n() > 1) %>%
  ungroup()


# Check structure
glimpse(df)

# --- 2. Correlation Between Program Participation & Symptom Improvement --- #

# Define a proxy for participation (using Submission Order or Submission Count)
df_participation <- df %>%
  group_by(Patients) %>%
  summarise(
    Total_Submissions = max(Submission.Order, na.rm = TRUE),
    First_PO = first(PO.Scorad),
    Last_PO = last(PO.Scorad),
    Improvement = First_PO - Last_PO
  ) %>%
  filter(Total_Submissions != 1)

# Correlation between Total Submissions and Improvement
cor_test <- cor.test(df_participation$Total_Submissions, df_participation$Improvement, use = "complete.obs")
print(cor_test)

# Scatterplot of Program Participation vs. Improvement
ggplot(df_participation, aes(x = Total_Submissions, y = Improvement)) +
  geom_point() +
  geom_smooth(method = "lm", se = FALSE, color = "blue") +
  labs(title = "Correlation Between Participation & Symptom Improvement",
       x = "Total Submissions (Program Participation)", y = "PO Scorad Improvement") +
  theme_minimal()

# --- 3. Linear Mixed Model for More Robust Analysis --- #
library(broom.mixed) # For tidy output

# Mixed model accounting for repeated measures per patient
# List of dependent variables to model
dependent_vars <- c("PO.Scorad", "Dryness", "Redness", "Swelling", "Scratch.marks", "Skin.thickening", "Itching.intensity", "Pyschological.State", "Sleep.Quality")

# Fit models for all dependent variables
results_days <- map_dfr(dependent_vars, function(var) {
  model <- lme(as.formula(paste(var, "~ Days.Since.First.Submission")), 
               random = ~1 | Patients, data = df, method = "REML")
  model_summary <- summary(model)
  
  tibble(
    Dependent_Variable = var,
    term = rownames(model_summary$tTable),
    estimate = model_summary$tTable[, "Value"],
    std.error = model_summary$tTable[, "Std.Error"],
    statistic = model_summary$tTable[, "t-value"],
    p.value = model_summary$tTable[, "p-value"],
    logLik = model_summary$logLik,
    AIC = AIC(model),
    BIC = BIC(model),
    sigma = model_summary$sigma
  )
})

# Output tidy dataframe with results
print(results_days)


# Fit models for all dependent variables considering submission
results_submissionorder <- map_dfr(dependent_vars, function(var) {
  model <- lme(as.formula(paste(var, "~ Submission.Order")), 
               random = ~1 | Patients, data = df, method = "REML")
  model_summary <- summary(model)
  
  tibble(
    Dependent_Variable = var,
    term = rownames(model_summary$tTable),
    estimate = model_summary$tTable[, "Value"],
    std.error = model_summary$tTable[, "Std.Error"],
    statistic = model_summary$tTable[, "t-value"],
    p.value = model_summary$tTable[, "p-value"],
    logLik = model_summary$logLik,
    AIC = AIC(model),
    BIC = BIC(model),
    sigma = model_summary$sigma
  )
})

# Output tidy dataframe with results
print(results_submissionorder)


# Compute improvement for each dependent variable
df_improvement <- df %>%
  group_by(Patients, Submission.Category) %>%
  summarise(across(c(PO.Scorad, Dryness, Redness, Swelling, Scratch.marks, Skin.thickening, Itching.intensity, Pyschological.State, Sleep.Quality), 
                   list(First = ~first(.), Last = ~last(.)), .names = "{col}_{fn}")) %>%
  mutate(across(ends_with("_Last"), ~ . - get(sub("_Last", "_First", cur_column())), .names = "{sub('_Last', '_Improvement', .col)}"))

# Convert to long format for visualization
df_improvement_long <- df_improvement %>%
  pivot_longer(cols = ends_with("_Improvement"), names_to = "Variable", values_to = "Improvement")

# Boxplot of improvement by Submission Category
ggplot(df_improvement_long, aes(x = Submission.Category, y = Improvement, fill = Submission.Category)) +
  geom_boxplot(alpha = 0.7) +
  facet_wrap(~ Variable, scales = "free") +
  labs(title = "Total Improvement Across Submission Categories",
       x = "Submission Category", y = "Improvement") +
  theme_minimal()


# --- 5. ANOVA to Test Differences Across Submission Categories --- #

# Function to perform one-way ANOVA
perform_one_way_anova <- function(data, response_vars, factors) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  
  # Iterate over response variables
  for (var in response_vars) {
    # Iterate over factors
    for (factor in factors) {
      # One-way ANOVA
      anova_model <- aov(reformulate(factor, response = var), data = data)
      model_summary <- summary(anova_model)
      
      # Extract relevant statistics
      model_results <- data.frame(
        Variable = var,
        Effect = factor,
        Sum_Sq = model_summary[[1]]$"Sum Sq"[1],  # Extract Sum Sq for the factor
        Mean_Sq = model_summary[[1]]$"Mean Sq"[1],  # Extract Mean Sq for the factor
        Df = model_summary[[1]]$Df[1],  # Extract Df for the factor
        FValue = model_summary[[1]]$"F value"[1],  # Extract F value for the factor
        pValue = model_summary[[1]]$"Pr(>F)"[1]  # Extract p-value for the factor
      )
      
      anova_results <- rbind(anova_results, model_results)
    }
  }
  
  return(anova_results)
}

factor <- "Submission.Category"  # Specify the factor (independent variable)
dependent_vars <- c("PO.Scorad_Improvement", "Dryness_Improvement", "Redness_Improvement", "Swelling_Improvement", "Scratch.marks_Improvement", "Skin.thickening_Improvement", "Itching.intensity_Improvement", "Pyschological.State_Improvement", "Sleep.Quality_Improvement")


# Perform one-way ANOVA
one_way_anova_results <- perform_one_way_anova(df_improvement, dependent_vars, factor)

library(multcomp)
library(emmeans)

# Function to perform appropriate post-hoc test for multiple response variables and factors
perform_posthoc_tests <- function(data, response_vars, factors, equal_variance) {
  results_list <- list() # Initialize a list to store results
  
  # Iterate over each response variable
  for (response_var in response_vars) {
    # Iterate over each factor
    for (factor in factors) {
      # Fit the model
      model <- aov(reformulate(factor, response = response_var), data = data)
      
      # If equal variance is assumed, perform Tukey's post-hoc test
      if (equal_variance) {
        posthoc_result <- TukeyHSD(model)
      } else {
        # Else perform Dunnett-T3 test
        # Calculate estimated marginal means (emmeans)
        emm <- emmeans(model, specs = factor)
        
        # Perform Dunnett-T3 test
        posthoc_result <- cld(emm, adjust = "dunnett")
      }
      
      # Store the results in the list with a unique name
      result_name <- paste0("Posthoc_", response_var, "_", factor)
      results_list[[result_name]] <- posthoc_result
    }
  }
  
  return(results_list)
}


posthoc_results <- perform_posthoc_tests(df_improvement, dependent_vars, factor, equal_variance = TRUE)

# Function to integrate post-hoc test results into a data frame
integrate_posthoc_results <- function(posthoc_results) {
  # Initialize an empty data frame to store combined results
  combined_results <- data.frame()
  
  # Iterate over each item in the post-hoc results list
  for (result_name in names(posthoc_results)) {
    # Extract the test results
    test_results <- posthoc_results[[result_name]]
    
    # Iterate over each element in the test results
    for (factor_name in names(test_results)) {
      # Retrieve the result matrix and convert it to a data frame
      result_matrix <- test_results[[factor_name]]
      result_df <- as.data.frame(result_matrix)
      
      # Add columns for the response variable and factor
      result_df$Response_Variable <- sub("Posthoc_", "", result_name)
      result_df$Factor <- factor_name
      result_df$Comparison <- rownames(result_df)
      
      # Bind the result to the combined data frame
      combined_results <- rbind(result_df, combined_results)
    }
  }
  
  # Reset row names
  rownames(combined_results) <- NULL
  
  return(combined_results)
}

posthoc_results <- integrate_posthoc_results(posthoc_results)


calculate_means_sds_medians_by_factors <- function(data, variables, factors) {
  # Create an empty list to store intermediate results
  results_list <- list()
  
  # Iterate over each variable
  for (var in variables) {
    # Create a temporary data frame to store results for this variable
    temp_results <- data.frame(Variable = var)
    
    # Iterate over each factor
    for (factor in factors) {
      # Aggregate data by factor: Mean, SD, Median
      agg_data <- aggregate(data[[var]], 
                            by = list(data[[factor]]), 
                            FUN = function(x) c(Mean = mean(x, na.rm = TRUE),
                                                SD = sd(x, na.rm = TRUE),
                                                Median = median(x, na.rm = TRUE)))
      
      # Create columns for each level of the factor
      for (level in unique(data[[factor]])) {
        level_agg_data <- agg_data[agg_data[, 1] == level, ]
        
        if (nrow(level_agg_data) > 0) {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- level_agg_data$x[1, "Mean"]
          temp_results[[paste0(factor, "_", level, "_SD")]] <- level_agg_data$x[1, "SD"]
          temp_results[[paste0(factor, "_", level, "_Median")]] <- level_agg_data$x[1, "Median"]
        } else {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- NA
          temp_results[[paste0(factor, "_", level, "_SD")]] <- NA
          temp_results[[paste0(factor, "_", level, "_Median")]] <- NA
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

# Example usage
descriptive_stats_bygroup <- calculate_means_sds_medians_by_factors(df_improvement, dependent_vars, "Submission.Category")


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
  "Data" = df, 
  "Data Improvement" = df_improvement, 
  "Descriptives" = descriptive_stats_bygroup,
  "ANOVA Results" = one_way_anova_results,
  "Post-Hoc ANOVA" = posthoc_results,
  "Model results - Days" = results_days,
  "Model results - Submissions" = results_submissionorder
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")