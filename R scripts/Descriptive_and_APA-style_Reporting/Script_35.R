setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/karjun24")

library(openxlsx)
df <- read.xlsx("Phd Testing_Transf.xlsx")


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
  return(x)}))


# Load the necessary libraries
library(lme4)
library(broom.mixed)
library(afex)
library(car)
library(dplyr)
library(tibble)


# Ensure that 'Exercises' and 'Loading' are factors
df$Exercises <- as.factor(df$Exercises)
df$Loading <- as.factor(df$Loading)
df$Gender <- as.factor(df$Gender)

# Define the function with specific model details
fit_lmm_and_format <- function(formula, data, response_var_name, save_plots = FALSE, plot_param = "") {
  # Fit the linear mixed-effects model
  lmer_model <- lmer(formula, data = data)
  
  # Print the summary of the model for fit statistics
  model_summary <- summary(lmer_model)
  print(model_summary)
  
  # Extract the tidy output for fixed effects and assign it to lmm_results
  lmm_results <- tidy(lmer_model, "fixed") %>%
    mutate(ResponseVariable = response_var_name, Type = "Fixed Effect")
  
  # Extract random effects variances
  random_effects_variances <- as.data.frame(VarCorr(lmer_model)) %>%
    tibble::rownames_to_column("Term") %>%
    mutate(ResponseVariable = response_var_name, Type = "Random Variance")
  
  # Generate and save residual plots
  if (save_plots) {
    jpeg(paste0("Residuals_vs_Fitted_", plot_param, ".jpeg"))
    plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
    dev.off()
    
    jpeg(paste0("QQ_Plot_", plot_param, ".jpeg"))
    qqnorm(residuals(lmer_model))
    qqline(residuals(lmer_model))
    dev.off()
  }
  
  # Return results
  return(list(fixed = lmm_results, random_variances = random_effects_variances))
}

# Example usage of the function for your specific dataset
results_X30m <- fit_lmm_and_format("X30m ~ Wave + Exercises + Loading + Age + Gender + (1 | Candidate.)", df, "X30m", TRUE, "X30m")
results_X60m <- fit_lmm_and_format("X60m ~ Wave + Exercises + Loading + Age + Gender + (1 | Candidate.)", df, "X60m", TRUE, "X60m")
results_X80m <- fit_lmm_and_format("X80m ~ Wave + Exercises + Loading + Age + Gender + (1 | Candidate.)", df, "X80m", TRUE, "X80m")

# Combine the results into a single dataframe
combined_results <- bind_rows(
  results_X30m$fixed, results_X30m$random_variances,
  results_X60m$fixed, results_X60m$random_variances,
  results_X80m$fixed, results_X80m$random_variances
)

# Print the combined results
print(combined_results)

# Modify column names to remove the 'X' from the start of the variable names
df <- df %>%
  rename_with(~ gsub("^X", "", .), .cols = starts_with("X"))

# Convert numerical values to factor with levels "No" and "Yes"
df$Exercises <- factor(df$Exercises, levels = c(0, 1), labels = c("No", "Yes"))


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


df_descriptive_crosstabs_wave <- calculate_descriptive_stats_bygroups(df, c("30m", "60m", "80m"), "Wave")
df_descriptive_crosstabs_loading <- calculate_descriptive_stats_bygroups(df, c("30m", "60m", "80m"), "Loading")
df_descriptive_crosstabs_exercise <- calculate_descriptive_stats_bygroups(df, c("30m", "60m", "80m"), "Exercises")


library(dplyr)


calculate_descriptive_stats_bygroups2 <- function(data, desc_vars, cat_column1, cat_column2) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each combination of categories
  for (var in desc_vars) {
    # Group data by both categorical columns using dynamic column names
    grouped_data <- data %>%
      group_by(across(all_of(cat_column1)), across(all_of(cat_column2))) %>%
      summarise(
        Mean = mean(.data[[var]], na.rm = TRUE),
        SEM = sd(.data[[var]], na.rm = TRUE) / sqrt(n()),
        SD = sd(.data[[var]], na.rm = TRUE),
        .groups = 'drop' # Drop grouping structure after summarising
      ) %>%
      mutate(Variable = var) %>%
      rename_with(~ c(cat_column1, cat_column2, "Mean", "SEM", "SD", "Variable"))
    
    # Append the results for each variable and category combination to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}


df_descriptive_twowaycrosstabs_Candidate <- calculate_descriptive_stats_bygroups2(df, c("30m", "60m", "80m"),"Candidate.", "Wave")
df_descriptive_twowaycrosstabs_Loading <- calculate_descriptive_stats_bygroups2(df, c("30m", "60m", "80m"),"Loading", "Wave")
df_descriptive_twowaycrosstabs_Exercises <- calculate_descriptive_stats_bygroups2(df, c("30m", "60m", "80m"),"Exercises", "Wave")



#Visualizations
library(ggplot2)

# Ensure the variables are factors in the data frames before plotting
df_descriptive_crosstabs_wave$Wave <- factor(df_descriptive_crosstabs_wave$Wave)
df_descriptive_crosstabs_loading$Loading <- factor(df_descriptive_crosstabs_loading$Loading)
df_descriptive_crosstabs_exercise$Exercises <- factor(df_descriptive_crosstabs_exercise$Exercises)
df_descriptive_twowaycrosstabs_Loading$Wave <- factor(df_descriptive_twowaycrosstabs_Loading$Wave)
df_descriptive_twowaycrosstabs_Exercises$Wave <- factor(df_descriptive_twowaycrosstabs_Exercises$Wave)

library(ggplot2)

# Ensure df_descriptive_crosstabs_exercise is updated with new factor levels for Exercises
df_descriptive_crosstabs_exercise$Exercises <- factor(df_descriptive_crosstabs_exercise$Exercises, levels = c("No", "Yes"))

# Bar Plot for Means by Exercise for all DVs with Data Labels
p_exercise <- ggplot(df_descriptive_crosstabs_exercise, aes(x = Exercises, y = Mean, fill = Variable)) +
  geom_bar(stat = "identity", position = position_dodge()) +
  geom_text(aes(label = sprintf("%.2f", Mean)), position = position_dodge(width = 0.9), vjust = -0.25) +
  labs(title = "Mean Scores by Exercise for 30m, 60m, 80m", x = "Exercise", y = "Mean Score") +
  theme_minimal()
print(p_exercise)
ggsave("Bar_Plots_Exercise_DVs.png", plot = p_exercise, width = 8, height = 6, dpi = 300)

# If needed, update other bar plots similarly by ensuring factor conversion for categorical variables and adding geom_text() for labels
# Example for Wave and Loading
df_descriptive_crosstabs_wave$Wave <- factor(df_descriptive_crosstabs_wave$Wave)
df_descriptive_crosstabs_loading$Loading <- factor(df_descriptive_crosstabs_loading$Loading)

p_wave <- ggplot(df_descriptive_crosstabs_wave, aes(x = Wave, y = Mean, fill = Variable)) +
  geom_bar(stat = "identity", position = position_dodge()) +
  geom_text(aes(label = sprintf("%.2f", Mean)), position = position_dodge(width = 0.9), vjust = -0.25) +
  labs(title = "Mean Scores by Wave for 30m, 60m, 80m", x = "Wave", y = "Mean Score") +
  theme_minimal()
print(p_wave)
ggsave("Bar_Plots_Wave_DVs.png", plot = p_wave, width = 8, height = 6, dpi = 300)

p_loading <- ggplot(df_descriptive_crosstabs_loading, aes(x = Loading, y = Mean, fill = Variable)) +
  geom_bar(stat = "identity", position = position_dodge()) +
  geom_text(aes(label = sprintf("%.2f", Mean)), position = position_dodge(width = 0.9), vjust = -0.25) +
  labs(title = "Mean Scores by Loading for 30m, 60m, 80m", x = "Loading", y = "Mean Score") +
  theme_minimal()
print(p_loading)
ggsave("Bar_Plots_Loading_DVs.png", plot = p_loading, width = 8, height = 6, dpi = 300)

# Interaction Plot for Wave and Loading on all DVs
p_interaction_loading <- ggplot(df_descriptive_twowaycrosstabs_Loading, aes(x = Wave, y = Mean, color = Loading, group = Loading)) +
  geom_line() +
  geom_point() +
  facet_wrap(~ Variable) +  # To display different DVs
  labs(title = "Interaction of Wave and Loading on 30m, 60m, 80m", x = "Wave", y = "Mean Score") +
  theme_minimal()
print(p_interaction_loading)
ggsave("Interaction_Plot_Loading_DVs.png", plot = p_interaction_loading, width = 10, height = 6, dpi = 300)

# Interaction Plot for Wave and Exercises on all DVs
p_interaction_exercises <- ggplot(df_descriptive_twowaycrosstabs_Exercises, aes(x = Wave, y = Mean, color = Exercises, group = Exercises)) +
  geom_line() +
  geom_point() +
  facet_wrap(~ Variable) +  # To display different DVs
  labs(title = "Interaction of Wave and Exercises on 30m, 60m, 80m", x = "Wave", y = "Mean Score") +
  theme_minimal()
print(p_interaction_exercises)
ggsave("Interaction_Plot_Exercises_DVs.png", plot = p_interaction_exercises, width = 10, height = 6, dpi = 300)



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
  "Wave Descriptive Stats" = df_descriptive_crosstabs_wave,
  "Loading Descriptive Stats" = df_descriptive_crosstabs_loading,
  "Exercise Descriptive Stats" = df_descriptive_crosstabs_exercise,
  "Candidate-Wave Crosstabs" = df_descriptive_twowaycrosstabs_Candidate,
  "Loading-Wave Crosstabs" = df_descriptive_twowaycrosstabs_Loading,
  "Exercise-Wave Crosstabs" = df_descriptive_twowaycrosstabs_Exercises,
  "Combined Model Results" = combined_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")


# Calculate the mean of the Age column
mean_age <- mean(df$Age, na.rm = TRUE)
print(paste("Mean Age:", mean_age))

# Get the frequency distribution of the Gender column
gender_distribution <- table(df$Gender)
print("Gender Distribution:")
print(gender_distribution)
