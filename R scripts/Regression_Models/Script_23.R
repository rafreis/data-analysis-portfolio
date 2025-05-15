# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elena_koing")

# Read data from a CSV file
df <- read.csv("Bausch_Meal_Timing_Data_Fiverr (2).csv")

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


colnames(df)



library(dplyr)
library(lubridate)

# Assuming df is already loaded and in the format you provided

# Convert meal_time to POSIXct for easier manipulation
df$meal_time <- ymd_hms(paste(df$meal_date, df$meal_time))
df$sunrise_time <- ymd_hms(paste(df$meal_date, df$sunrise_time))
df$sunset_time <- ymd_hms(paste(df$meal_date, df$sunset_time))

# Define a function to group meal times that are within 30 minutes of each other
group_meals <- function(times) {
  sorted <- sort(times)
  diffs <- c(Inf, difftime(sorted[-1], sorted[-length(sorted)], units = "mins"))
  group <- cumsum(diffs > 30)
  return(group)
}

# Apply the function to group meals together
df <- df %>%
  group_by(patient_id, meal_date, meal_name) %>%
  mutate(meal_group = group_meals(meal_time)) %>%
  group_by(patient_id, meal_date, meal_name, meal_group, sunrise_time, sunset_time) %>%
  summarise(
    aggregated_meal_time = mean(meal_time),
    .groups = 'drop'
  )

df$aggregated_meal_time <- ymd_hms(df$aggregated_meal_time)  # Ensure time format is POSIXct


# Proceed with the earlier analyses using 'aggregated_meal_time' instead of 'meal_time'
# Calculate SD for meal times
meal_regularity_permeal <- df %>%
  group_by(patient_id, meal_name) %>%
  summarise(
    # Calculate the decimal hour of meal time for more precise variability measure
    meal_hour = hour(aggregated_meal_time) + minute(aggregated_meal_time) / 60 + second(aggregated_meal_time) / 3600,
    .groups = 'drop'
  ) %>%
  group_by(patient_id, meal_name) %>%
  summarise(
    # Standard deviation of meal times in decimal hours
    meal_regularity_sd = sd(meal_hour, na.rm = TRUE),
    .groups = 'drop'
  )

# Calculate eating window per day
eating_windows_permeal <- df %>%
  group_by(patient_id, meal_date) %>%
  summarise(
    first_meal = min(aggregated_meal_time),
    last_meal = max(aggregated_meal_time),
    .groups = 'drop'
  ) %>%
  mutate(
    # Calculate the difference in hours, not seconds
    eating_window = as.numeric(difftime(last_meal, first_meal, units = "hours"))
  ) %>%
  group_by(patient_id) %>%
  summarise(
    # Standard deviation of the eating window in hours
    eating_window_sd = sd(eating_window, na.rm = TRUE)
  )

# Calculate phase angle score
phase_angle_permeal <- df %>%
  group_by(patient_id, meal_date) %>%
  summarise(
    first_meal = min(aggregated_meal_time),
    last_meal = max(aggregated_meal_time),
    sunrise = first(sunrise_time),
    sunset = first(sunset_time),
    .groups = 'drop'
  ) %>%
  mutate(
    # Calculate time differences in hours
    morning_phase = abs(as.numeric(difftime(first_meal, sunrise, units = "hours"))),
    evening_phase = abs(as.numeric(difftime(sunset, last_meal, units = "hours")))
  ) %>%
  group_by(patient_id) %>%
  summarise(
    morning_phase_sd = sd(morning_phase, na.rm = TRUE),
    evening_phase_sd = sd(evening_phase, na.rm = TRUE),
    mean_phase_angle = (morning_phase_sd + evening_phase_sd) / 2,
    .groups = 'drop'
  )

# Combine all scores into one dataframe
erd_scores_permeal <- meal_regularity_permeal %>%
  left_join(eating_windows_permeal, by = 'patient_id') %>%
  left_join(phase_angle_permeal, by = 'patient_id') %>%
  mutate(total_erd = meal_regularity_sd + eating_window_sd + mean_phase_angle)



# Overall

library(dplyr)
library(lubridate)

# Calculate Meal Regularity Score Across All Meals
meal_regularity <- df %>%
  group_by(patient_id) %>%
  mutate(
    meal_hour = hour(aggregated_meal_time) + minute(aggregated_meal_time) / 60 + second(aggregated_meal_time) / 3600
  ) %>%
  summarise(
    # Calculate the standard deviation of meal times in decimal hours across all meals
    meal_regularity_sd = sd(meal_hour, na.rm = TRUE),
    .groups = 'drop'
  )

# Calculate Eating Window per Day, Then SD of Eating Window Across All Days
eating_windows <- df %>%
  group_by(patient_id, meal_date) %>%
  summarise(
    first_meal = min(aggregated_meal_time),
    last_meal = max(aggregated_meal_time),
    .groups = 'drop'
  ) %>%
  mutate(
    eating_window = as.numeric(difftime(last_meal, first_meal, units = "hours"))
  ) %>%
  group_by(patient_id) %>%
  summarise(
    eating_window_sd = sd(eating_window, na.rm = TRUE),
    .groups = 'drop'
  )

# Calculate Phase Angle Scores
phase_angle <- df %>%
  group_by(patient_id, meal_date) %>%
  summarise(
    first_meal = min(aggregated_meal_time),
    last_meal = max(aggregated_meal_time),
    sunrise = first(sunrise_time),
    sunset = first(sunset_time),
    .groups = 'drop'
  ) %>%
  mutate(
    morning_phase = abs(as.numeric(difftime(first_meal, sunrise, units = "hours"))),
    evening_phase = abs(as.numeric(difftime(sunset, last_meal, units = "hours")))
  ) %>%
  group_by(patient_id) %>%
  summarise(
    morning_phase_sd = sd(morning_phase, na.rm = TRUE),
    evening_phase_sd = sd(evening_phase, na.rm = TRUE),
    mean_phase_angle = (morning_phase_sd + evening_phase_sd) / 2,
    .groups = 'drop'
  )

# Combine all scores into one dataframe
erd_scores <- meal_regularity %>%
  left_join(eating_windows, by = 'patient_id') %>%
  left_join(phase_angle, by = 'patient_id') %>%
  mutate(
    total_erd = meal_regularity_sd + eating_window_sd + mean_phase_angle
  )



library(ggplot2)
library(dplyr)

# Create a day index for each patient starting at 1 for the earliest day
df <- df %>%
  group_by(patient_id) %>%
  mutate(day_number = as.numeric(as.Date(meal_date) - min(as.Date(meal_date)) + 1)) %>%
  ungroup()

# Calculate the number of plots needed
total_patients <- length(unique(df$patient_id))
patients_per_plot <- 10
num_plots <- ceiling(total_patients / patients_per_plot)

# Create and save plots
for (i in 1:num_plots) {
  # Calculate the range of patient IDs for this plot
  start_index <- (i - 1) * patients_per_plot + 1
  end_index <- min(i * patients_per_plot, total_patients)
  patient_subset <- unique(df$patient_id)[start_index:end_index]
  
  # Subset the data for this range of patient IDs
  df_subset <- df %>% filter(patient_id %in% patient_subset)
  
  # Add placeholders if needed to fill up the 2x5 grid
  if (length(patient_subset) < patients_per_plot) {
    additional_slots <- patients_per_plot - length(patient_subset)
    placeholder_df <- data.frame(
      patient_id = paste0("Placeholder_", seq(1, additional_slots)),
      meal_date = NA,
      meal_name = NA,
      meal_group = NA,
      sunrise_time = NA,
      sunset_time = NA,
      aggregated_meal_time = NA,
      day_number = NA
    )
    df_subset <- bind_rows(df_subset, placeholder_df)
  }
  
  # Ensure 'patient_id' is treated as a factor to control panel ordering
  df_subset$patient_id <- factor(df_subset$patient_id, levels = unique(df_subset$patient_id))
  
  # Create the plot
  p <- ggplot(df_subset, aes(x = hour(aggregated_meal_time) + minute(aggregated_meal_time)/60, y = day_number)) +
    geom_point(color = "green", size = 2, alpha = 0.6) +
    facet_wrap(~ patient_id, ncol = 2, nrow = 5, scales = "free_y") +  # Ensure 2x5 grid
    labs(x = "Time of Day (24h)", y = "Day Number", title = paste("Eating Patterns: Batch", i)) +
    scale_x_continuous(breaks = seq(0, 24, by = 3), labels = sprintf("%02d:00", seq(0, 24, by = 3))) +  # Format x-axis in 24-hour time
    scale_y_continuous(breaks = seq(min(df_subset$day_number, na.rm = TRUE), max(df_subset$day_number, na.rm = TRUE), by = 1),  # Set discrete y-axis breaks
                       labels = as.character(seq(min(df_subset$day_number, na.rm = TRUE), max(df_subset$day_number, na.rm = TRUE), by = 1))) +
    theme_minimal() +
    theme(
      axis.text.x = element_text(angle = 45, hjust = 1),
      axis.text.y = element_text(size = 8),  # Adjust y-axis text size for readability
      strip.background = element_blank(),
      panel.spacing = unit(0.5, "lines"),
      panel.background = element_rect(fill = "grey90", colour = "grey50"),
      plot.background = element_rect(fill = "grey98"),
      strip.text.x = element_text(size = 10, face = "bold")
    )
  
  # Define the filename based on the batch
  filename <- sprintf("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elena_koing/Plots/Eating_Patterns_Batch_%d.png", i)
  
  # Save the plot
  ggsave(filename, plot = p, width = 15, height = 10, dpi = 300)  # Adjusted size to accommodate the matrix layout
  
}





# Only 15 plots

library(ggplot2)
library(dplyr)
library(lubridate)

valid_patient_ids <- df %>%
  group_by(patient_id) %>%
  filter(n_distinct(day_number) >= 5 & max(day_number) <= 50) %>%
  ungroup() %>%
  distinct(patient_id) %>%
  pull(patient_id)

# Filter 'erd_scores' to only include these valid patient IDs
filtered_erd_scores <- erd_scores %>% filter(patient_id %in% valid_patient_ids)

# Select top 5, middle 5, and bottom 5 based on 'total_erd' from the filtered data
top_5_patients <- filtered_erd_scores %>% arrange(desc(total_erd)) %>% slice_head(n = 5) %>% pull(patient_id)
bottom_5_patients <- filtered_erd_scores %>% arrange(total_erd) %>% slice_head(n = 5) %>% pull(patient_id)
mid_range_index <- round(nrow(filtered_erd_scores) / 2)
mid_5_patients <- filtered_erd_scores %>% arrange(total_erd) %>% slice((mid_range_index - 2):(mid_range_index + 2)) %>% pull(patient_id)

# Create subsets for each ERD group
df_top_subset <- df %>% filter(patient_id %in% top_5_patients)
df_mid_subset <- df %>% filter(patient_id %in% mid_5_patients)
df_low_subset <- df %>% filter(patient_id %in% bottom_5_patients)


# Define a function to create and save individual plots for a specific group
create_individual_plots <- function(df_subset, erd_label) {
  for (patient_id in unique(df_subset$patient_id)) {
    # Subset the data for the specific patient
    patient_data <- df_subset %>% filter(patient_id == patient_id)
    
    # Ensure 'aggregated_meal_time' is in POSIXct format
    patient_data$aggregated_meal_time <- ymd_hms(patient_data$aggregated_meal_time)
    
    # Create the plot
    p <- ggplot(patient_data, aes(x = hour(aggregated_meal_time) + minute(aggregated_meal_time)/60, y = day_number)) +
      geom_point(color = "darkblue", size = 3, alpha = 0.8) +  # Use darker color for points
      labs(x = "Time of Day (24h)", y = "Day Number", title = paste("Eating Patterns for Patient", patient_id, "-", erd_label, "ERD")) +
      scale_x_continuous(breaks = seq(0, 24, by = 3), labels = sprintf("%02d:00", seq(0, 24, by = 3))) +  # Format x-axis in 24-hour time
      scale_y_continuous(breaks = seq(min(patient_data$day_number, na.rm = TRUE), max(patient_data$day_number, na.rm = TRUE), by = 1),  # Set discrete y-axis breaks
                         labels = as.character(seq(min(patient_data$day_number, na.rm = TRUE), max(patient_data$day_number, na.rm = TRUE), by = 1))) +
      theme_minimal(base_size = 12) +  # Base size for better readability
      theme(
        panel.background = element_rect(fill = "white", colour = NA),  # White background
        panel.grid = element_blank(),  # Remove gridlines
        axis.text.x = element_text(angle = 0, hjust = 1),
        axis.text.y = element_text(size = 10),  # Adjust y-axis text size for readability
        plot.background = element_rect(fill = "white", colour = NA),  # White plot background
        strip.background = element_blank(),
        panel.spacing = unit(0.5, "lines"),
        strip.text.x = element_text(size = 12, face = "bold")
      )
    
    # Define the filename for the plot
    filename <- sprintf("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elena_koing/Plots/Eating_Patterns_Patient_%s_%s_ERD.png", patient_id, erd_label)
    
    # Save the plot
    ggsave(filename, plot = p, width = 10, height = 8, dpi = 300)
  }
}

# Create individual plots for each ERD group
create_individual_plots(df_top_subset, "High")
create_individual_plots(df_mid_subset, "Mid")
create_individual_plots(df_low_subset, "Low")



library(ggplot2)
library(tidyr)
library(stringr)

create_boxplots <- function(df, vars) {
  # Loop through each variable in vars and create a boxplot for each
  for (var in vars) {
    # Create a separate plot for each variable
    p <- ggplot(df, aes_string(x = "factor(1)", y = var)) +
      geom_boxplot() +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
      labs(title = paste("Boxplot of", var), x = var, y = "Value") +
      theme(axis.text.x = element_blank())  # Remove x-axis labels since they aren't needed
    
    # Print the plot
    print(p)
    
    # Save each plot as a separate file
    ggsave(filename = paste0("boxplot_", var, ".png"), plot = p, width = 6, height = 4)
  }
}

# Example usage:
# Assuming df_recoded is your dataframe and numerical_vars contains the columns you want to plot
create_boxplots(erd_scores, "total_erd")


library(ggplot2)
library(tidyr)
library(stringr)

create_histograms <- function(df, vars) {
  # Loop through each variable in vars and create a histogram for each
  for (var in vars) {
    # Create a separate histogram for each variable
    p <- ggplot(df, aes_string(x = var)) +
      geom_histogram(binwidth = 1, fill = "lightblue", color = "black") +
      geom_vline(aes(xintercept = mean(get(var), na.rm = TRUE)), color = "red", linetype = "dashed", size = 1) +
      labs(title = paste("Histogram of", var), x = var, y = "Frequency") +
      theme_minimal()
    
    # Print the plot
    print(p)
    
    # Save each plot as a separate file
    ggsave(filename = paste0("histogram_", var, ".png"), plot = p, width = 6, height = 4)
  }
}

# Example usage:
# Assuming erd_scores is your dataframe and vars contains the columns you want to plot
create_histograms(erd_scores, c("total_erd"))


library(openxlsx)

# Create a new workbook
wb <- createWorkbook()

# Add worksheets for each data frame
addWorksheet(wb, "erd_scores")
addWorksheet(wb, "erd_scores_permeal")

# Write data to the worksheets
writeData(wb, sheet = "erd_scores", x = erd_scores)
writeData(wb, sheet = "erd_scores_permeal", x = erd_scores_permeal)

# Save the workbook
saveWorkbook(wb, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elena_koing/erd_scores_export.xlsx", overwrite = TRUE)


colnames(erd_scores)

# OLS Regression

library(broom)
library(dplyr)

fit_ols_and_format <- function(data, predictors, response_vars, save_plots = FALSE) {
  ols_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Fit the OLS multiple regression model
    lm_model <- lm(formula, data = data)
    
    # Get the summary of the model
    model_summary <- summary(lm_model)
    
    # Extract the R-squared, F-statistic, and p-value
    r_squared <- model_summary$r.squared
    f_statistic <- model_summary$fstatistic[1]
    p_value <- pf(f_statistic, model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE)
    
    # Print the summary of the model for fit statistics
    print(summary(lm_model))
    
    # Extract the tidy output and assign it to ols_results
    ols_results <- broom::tidy(lm_model) %>%
      mutate(ResponseVariable = response_var, R_Squared = r_squared, F_Statistic = f_statistic, P_Value = p_value)
    
    # Optionally print the tidy output
    print(ols_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lm_model) ~ fitted(lm_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lm_model), main = "Q-Q Plot")
      qqline(residuals(lm_model))
      dev.off()
    }
    
    # Store the results in a list
    ols_results_list[[response_var]] <- ols_results
  }
  
  return(ols_results_list)
}

ols_results_list_both_groups <- fit_ols_and_format(
  data = erd_scores,
  predictors = c("meal_regularity_sd", "eating_window_sd","mean_phase_angle"),  # Replace with actual predictor names
  response_vars = "total_erd",       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults <- bind_rows(ols_results_list_both_groups)


pca_loadings <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Scale = character(),
    Variable = character(),
    ComponentLoading = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[, scales[[scale]]]
    
    # Remove rows with any missing values in the subset data
    subset_data <- subset_data[complete.cases(subset_data), ]
    
    # Perform PCA, standardizing the data
    pca_results <- prcomp(subset_data, scale. = TRUE)
    
    # Extract loadings for the first principal component
    loadings <- pca_results$rotation[, 1]
    
    # Append results for each item in the scale
    for (item in scales[[scale]]) {
      results <- rbind(results, data.frame(
        Scale = scale,
        Variable = item,
        ComponentLoading = loadings[item]
      ))
    }
    
    # Calculate total variance explained by the first principal component
    variance_explained <- summary(pca_results)$importance[2, 1]
    
    # Print summary of the results - Total Variance Explained
    cat("Scale:", scale, "\n")
    cat("Total Variance Explained by First Component:", variance_explained, "\n")
  }
  
  return(results)
}

# Example call with the erd_scores dataset and your defined scales
df_pca_results <- pca_loadings(erd_scores, list(
  scale1 = c("meal_regularity_sd", "eating_window_sd", "mean_phase_angle")
))


bootstrap_pca <- function(data, n_boot = 1000) {
  # Perform PCA on the original data to get reference loadings
  original_pca <- prcomp(data, scale. = TRUE)
  reference_loadings <- original_pca$rotation[, 1]
  
  # Initialize a matrix to store loadings from bootstrap samples
  loadings_matrix <- matrix(NA, nrow = ncol(data), ncol = n_boot)
  rownames(loadings_matrix) <- colnames(data)
  
  # Perform bootstrapping
  for (i in 1:n_boot) {
    # Resample data with replacement
    boot_data <- data[sample(1:nrow(data), replace = TRUE), ]
    
    # Perform PCA on the bootstrapped data
    boot_pca <- prcomp(boot_data, scale. = TRUE)
    boot_loadings <- boot_pca$rotation[, 1]
    
    # Align signs of loadings with the reference
    if (sum(boot_loadings * reference_loadings) < 0) {
      boot_loadings <- -boot_loadings
    }
    
    # Store the aligned loadings
    loadings_matrix[, i] <- boot_loadings
  }
  
  # Calculate mean and confidence intervals for the loadings
  ci <- apply(loadings_matrix, 1, function(x) {
    c(mean = mean(x), lower = quantile(x, 0.025), upper = quantile(x, 0.975))
  })
  
  return(as.data.frame(t(ci)))
}

# Example usage
subset_data <- erd_scores[, c("meal_regularity_sd", "eating_window_sd", "mean_phase_angle")]
subset_data <- subset_data[complete.cases(subset_data), ]
boot_results <- bootstrap_pca(subset_data)
print(boot_results)


# Create a new workbook
wb <- createWorkbook()

# Add worksheets for each data frame
addWorksheet(wb, "PCA Results")
addWorksheet(wb, "Bootstrap")

# Write data to the worksheets
writeData(wb, sheet = "PCA Results", x = df_pca_results)
writeData(wb, sheet = "Bootstrap", x = boot_results)

# Save the workbook
saveWorkbook(wb, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elena_koing/PCA Results.xlsx", overwrite = TRUE)

