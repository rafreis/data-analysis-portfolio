# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/charleshenryher")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Data_Analysis.xlsx")

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


library(dplyr)
library(lubridate)

df <- df %>%
  mutate(
    Start_Date = as.Date(as.numeric(Start_Date), origin = "1899-12-30"),  
    Test_Date = as.Date(as.numeric(Test_Date), origin = "1899-12-30"),
    X = as.numeric(X),
    Y = as.numeric(Y),
    Z = as.numeric(Z)
  )

# Exclude patients on Drug A without known total doses
df_filter1 <- df %>%
  filter(!(Drug == "A" & is.na(Number_Doses)))


# Identify patients who took both drugs
patients_both_drugs <- df_filter1 %>%
  group_by(Patient_Id) %>%
  filter(n_distinct(Drug) > 1) %>%
  pull(Patient_Id) %>%
  unique()

# Exclude patients who took both drugs
df_filter2 <- df_filter1 %>%
  filter(!Patient_Id %in% patients_both_drugs)


# Exclude patients with only one test for Drug B
df_filter3 <- df_filter2 %>%
  group_by(Patient_Id, Drug) %>%
  filter(!(Drug == "B" & n() == 1)) %>%
  ungroup()

# Calculate the time interval in years and the change in parameters
df_filter3 <- df_filter3 %>%
  arrange(Patient_Id, Drug, Test_Date) %>%
  group_by(Patient_Id, Drug) %>%
  mutate(
    # Calculate interval using Start_Date for the first record and Test_Date for subsequent records
    Interval_Years = if_else(row_number() == 1,
                             as.numeric(difftime(Test_Date, Start_Date, units = "days") / 365.25),
                             as.numeric(difftime(Test_Date, lag(Test_Date), units = "days") / 365.25)),
    Change_X = X - lag(X),
    Change_Y = Y - lag(Y),
    Change_Z = Z - lag(Z)
  ) %>%
  ungroup()

# Calculate annual degradation rates
df_filter3 <- df_filter3 %>%
  filter(!is.na(Interval_Years)) %>%
  mutate(
    Rate_X = Change_X / Interval_Years,
    Rate_Y = Change_Y / Interval_Years,
    Rate_Z = Change_Z / Interval_Years
  )



# Initial exploration of degradation rates
summary_select <- df_filter3 %>%
  group_by(Drug) %>%
  summarise(
    Mean_Rate_X = mean(Rate_X, na.rm = TRUE),
    Mean_Rate_Y = mean(Rate_Y, na.rm = TRUE),
    Mean_Rate_Z = mean(Rate_Z, na.rm = TRUE)
  )

print(summary_select)


# Linear Mixed Model

library(lme4)
library(broom.mixed)
library(emmeans)
library(car)

fit_lmm_and_format <- function(formula, data, plot_param = "", save_plots = FALSE) {
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
  if (length(fixef(lmer_model)) > 2) {
    vif_values <- vif(lmer_model)
    print(vif_values)
  }
  
  # Conduct pairwise comparisons using emmeans
  emm_results <- emmeans(lmer_model, specs = pairwise ~ Drug)
  contrasts_results <- summary(emm_results$contrasts)
  print(contrasts_results)
  
  # Generate and save residual plots if requested
  if (save_plots) {
    residual_plot_filename <- paste0("Residuals_vs_Fitted_", plot_param, ".jpeg")
    qq_plot_filename <- paste0("QQ_Plot_", plot_param, ".jpeg")
    
    jpeg(residual_plot_filename)
    plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
    dev.off()
    
    jpeg(qq_plot_filename)
    qqnorm(residuals(lmer_model))
    qqline(residuals(lmer_model))
    dev.off()
  }
  
  # Return results
  return(list(
    fixed = lmm_results,
    random_variances = random_effects_variances,
    vif = vif_values,
    emmeans = emm_results,
    pairwise_contrasts = contrasts_results
  ))
}


# Apply function to parameter X
results_x <- fit_lmm_and_format("Rate_X ~ Drug + (1|Patient_Id)", df_filter3, save_plots = TRUE, plot_param = "X")

# Apply function to parameter Y
results_y <- fit_lmm_and_format("Rate_Y ~ Drug + (1|Patient_Id)", df_filter3, save_plots = TRUE, plot_param = "Y")

# Apply function to parameter Z
results_z <- fit_lmm_and_format("Rate_Z ~ Drug + (1|Patient_Id)", df_filter3, save_plots = TRUE, plot_param = "Z")


fixed_effects_df <- bind_rows(
  results_x$fixed,
  results_y$fixed,
  results_z$fixed
)

random_variances_df <- bind_rows(
  results_x$random_variances,
  results_y$random_variances,
  results_z$random_variances
)


# Combine emmeans results into one dataframe
emmeans_df <- bind_rows(
  as.data.frame(summary(results_x$emmeans$emmeans, infer = c(TRUE, TRUE))) %>%
    mutate(Parameter = "X"),
  as.data.frame(summary(results_y$emmeans$emmeans, infer = c(TRUE, TRUE))) %>%
    mutate(Parameter = "Y"),
  as.data.frame(summary(results_z$emmeans$emmeans, infer = c(TRUE, TRUE))) %>%
    mutate(Parameter = "Z")
)

contrasts_df <- bind_rows(
  as.data.frame(summary(results_x$emmeans$contrasts, infer = c(TRUE, TRUE))) %>%
    mutate(Parameter = "X"),
  as.data.frame(summary(results_y$emmeans$contrasts, infer = c(TRUE, TRUE))) %>%
    mutate(Parameter = "Y"),
  as.data.frame(summary(results_z$emmeans$contrasts, infer = c(TRUE, TRUE))) %>%
    mutate(Parameter = "Z")
)

library(tidyr) 

# Add lagged Change_Z for 1, 2, and 3 periods
df_lagged <- df_filter3 %>%
  arrange(Patient_Id, Test_Date) %>%
  group_by(Patient_Id) %>%
  mutate(
    Lagged_Change_Z1 = lag(Change_Z, 1),
    Lagged_Change_Z2 = lag(Change_Z, 2),
    Lagged_Change_Z3 = lag(Change_Z, 3)
  ) %>%
  ungroup()

# Separate datasets for each lag period
# For Lag 1
df_lag1 <- df_lagged %>%
  select(-Lagged_Change_Z2, -Lagged_Change_Z3) %>%
  drop_na(Lagged_Change_Z1) %>%
  group_by(Patient_Id) %>%
  filter(n() > 1) %>%
  ungroup()

# For Lag 2
df_lag2 <- df_lagged %>%
  select(-Lagged_Change_Z1, -Lagged_Change_Z3) %>%
  drop_na(Lagged_Change_Z2) %>%
  group_by(Patient_Id) %>%
  filter(n() > 1) %>%
  ungroup()

# For Lag 3
df_lag3 <- df_lagged %>%
  select(-Lagged_Change_Z1, -Lagged_Change_Z2) %>%
  drop_na(Lagged_Change_Z3) %>%
  group_by(Patient_Id) %>%
  filter(n() > 1) %>%
  ungroup()

library(lme4)
library(broom)  # For tidy model summaries

# Function to fit models and return tidy results
fit_and_summarize <- function(data, formula) {
  model <- lmer(formula, data = data)
  summary(model)
  tidy(model)
}

# Model with 1 lag
results_lag1 <- fit_and_summarize(df_filter3, "Change_X ~ Lagged_Change_Z1 + (1|Patient_Id) + (1|Drug)")
print(results_lag1)

# Model with 2 lags
results_lag2 <- fit_and_summarize(df_filter3, "Change_X ~ Lagged_Change_Z2 + (1|Patient_Id) + (1|Drug)")
print(results_lag2)

# Model with 3 lags
results_lag3 <- fit_and_summarize(df_filter3, "Change_X ~ Lagged_Change_Z3 + (1|Patient_Id) + (1|Drug)")
print(results_lag3)

library(ggplot2)

# Updated function to plot scatter plots with fitted lines without titles
plot_lag_effect <- function(data, lag_column) {
  ggplot(data, aes_string(x = lag_column, y = "Change_X", color = "Drug")) +
    geom_point(alpha = 0.6, show.legend = TRUE) +  # Show points with some transparency
    geom_smooth(method = "lm", se = TRUE) +        # Add linear model fit lines with confidence intervals
    labs(x = paste("Lagged Change in Z (", lag_column, ")", sep = ""),
         y = "Change in X") +
    theme_minimal() +  # Use a minimal theme
    theme(legend.position = "bottom")  # Position the legend at the bottom
}

# Generate plots for each lag
plot1 <- plot_lag_effect(df_filter3, "Lagged_Change_Z1")
plot2 <- plot_lag_effect(df_filter3, "Lagged_Change_Z2")
plot3 <- plot_lag_effect(df_filter3, "Lagged_Change_Z3")

# If using an environment that supports grid layout, such as RStudio
library(gridExtra)
combined_plot <- grid.arrange(plot1, plot2, plot3, nrow = 1)

# Save the combined plot
ggsave("lag_effect_plots.png", combined_plot, width = 14, height = 5, units = "in")

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
  "Fixed Effects" = fixed_effects_df,
  "Random Variances" = random_variances_df,
  "EM Means" = emmeans_df,
  "Lag 1" = results_lag1,
  "Lag 2" = results_lag2,
  "Lag 3" = results_lag3
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")




library(lme4)
library(broom)
library(dplyr)

# Function to fit models and extract tidy summaries
fit_and_summarize <- function(data, formula) {
  model <- lmer(formula, data = data)
  summary(model)  # Optionally print the summary for immediate insight
  tidy(model)  # Return tidy results for further analysis
}

# Initialize a list to store the results
results_list <- list()

# Get unique drugs from any of the lagged datasets (assuming they are similar)
unique_drugs <- unique(df_lag1$Drug)

# Process each dataset for each drug and each lag
for (drug in unique_drugs) {
  # Iterate over each lag dataset
  for (lag in 1:3) {
    # Select the appropriate lagged dataframe
    current_df <- get(paste0("df_lag", lag))
    
    # Filter data for the current drug
    drug_data <- current_df %>% filter(Drug == drug)
    
    # Define the model formula
    formula <- as.formula(paste("Change_X ~", paste0("Lagged_Change_Z", lag), "+ (1|Patient_Id)"))
    
    # Fit the model and store the results with a unique identifier
    results_list[[paste("drug", drug, "lag", lag)]] <- fit_and_summarize(drug_data, formula)
  }
}

# Extracting specific results for review or further processing
# Assuming specific drugs and lag combinations, update as needed
df_drug_a_lag1 <- results_list[["drug_A_lag1"]]
df_drug_a_lag2 <- results_list[["drug_A_lag2"]]
df_drug_a_lag3 <- results_list[["drug_A_lag3"]]
df_drug_b_lag1 <- results_list[["drug_B_lag1"]]
df_drug_b_lag2 <- results_list[["drug_B_lag2"]]
df_drug_b_lag3 <- results_list[["drug_B_lag3"]]


# Summary statistics for each patient to see the range and standard deviation within patients
variability_summary <- df_lag1 %>%
  group_by(Patient_Id) %>%
  summarise(across(starts_with("Lagged"), list(mean = mean, sd = sd, min = min, max = max), .names = "stat_{.fn}_{.col}"))
print(variability_summary)



library(ggplot2)
library(dplyr)


df_filter3 <- df_filter3 %>%
  arrange(Patient_Id, Test_Date) %>%
  group_by(Patient_Id) %>%
  mutate(Dose_Number = row_number() - 1) %>%
  ungroup()

library(dplyr)
library(ggplot2)

# Define a function to remove outliers based on the IQR
remove_outliers <- function(x) {
  qnt <- quantile(x, probs=c(.25, .75), na.rm = TRUE)
  H <- 1.5 * IQR(x, na.rm = TRUE)
  x[x < (qnt[1] - H) | x > (qnt[2] + H)] <- NA
  return(x)
}

# Apply the function to the Rate columns
df_filter4 <- df_filter3 %>%
  mutate(
    Rate_X = remove_outliers(Rate_X),
    Rate_Y = remove_outliers(Rate_Y),
    Rate_Z = remove_outliers(Rate_Z)
  )

# Drop rows with NA in Rate columns (created by removing outliers)
df_filter4 <- drop_na(df_filter3, c(Rate_X, Rate_Y, Rate_Z))

# Calculate means of rates across all patients
mean_rates <- df_filter4 %>%
  group_by(Drug, Dose_Number) %>%
  summarise(
    Mean_Rate_X = mean(Rate_X, na.rm = TRUE),
    Mean_Rate_Y = mean(Rate_Y, na.rm = TRUE),
    Mean_Rate_Z = mean(Rate_Z, na.rm = TRUE),
    .groups = 'drop'  # This will drop the grouping after summarisation
  )

mean_rates <- mean_rates %>%
  group_by(Drug) %>%
  mutate(
    Cumulative_X = cumsum(Mean_Rate_X),
    Cumulative_Y = cumsum(Mean_Rate_Y),
    Cumulative_Z = cumsum(Mean_Rate_Z)
  ) %>%
  ungroup()

# Add a row for each drug with zero baselines
baseline_rows <- distinct(mean_rates, Drug) %>%
  mutate(
    Dose_Number = 0,
    Mean_Rate_X = 0,
    Mean_Rate_Y = 0,
    Mean_Rate_Z = 0,
    Cumulative_X = 0,  # Initialize cumulative sums as zero
    Cumulative_Y = 0,
    Cumulative_Z = 0
  )

# Bind the baseline rows to the original data and sort
mean_rates <- bind_rows(baseline_rows, mean_rates) %>%
  arrange(Drug, Dose_Number)

mean_rates <- mean_rates %>%
  mutate(Dose_Number = as.integer(Dose_Number))

library(ggplot2)
str(mean_rates)

# Create a vector of dose numbers to use for vertical lines
dose_numbers <- unique(mean_rates$Dose_Number)

# Plot Cumulative Change for Rate X with reference lines at each dose number
plot_cumulative_x <- ggplot(mean_rates, aes(x = Dose_Number, y = Cumulative_X, color = Drug, group = Drug)) +
  geom_line() +
  geom_point() +
  labs(title = "Cumulative Change of X by Drug over Doses", x = "Dose Number", y = "Cumulative Change in X") +
  theme_minimal() +
  theme(legend.title = element_blank(), axis.text.x = element_text(hjust = 0.5)) +
  scale_x_continuous(breaks = dose_numbers) +  # Ensure all dose numbers appear as ticks
  geom_vline(xintercept = dose_numbers, linetype = "dotted", color = "gray")  # Add reference lines

# Plot Cumulative Change for Rate Y
plot_cumulative_y <- ggplot(mean_rates, aes(x = Dose_Number, y = Cumulative_Y, color = Drug, group = Drug)) +
  geom_line() +
  geom_point() +
  labs(title = "Cumulative Change of Y by Drug over Doses", x = "Dose Number", y = "Cumulative Change in Y") +
  theme_minimal() +
  theme(legend.title = element_blank(), axis.text.x = element_text(hjust = 0.5)) +
  scale_x_continuous(breaks = dose_numbers) +
  geom_vline(xintercept = dose_numbers, linetype = "dotted", color = "gray")

# Plot Cumulative Change for Rate Z
plot_cumulative_z <- ggplot(mean_rates, aes(x = Dose_Number, y = Cumulative_Z, color = Drug, group = Drug)) +
  geom_line() +
  geom_point() +
  labs(title = "Cumulative Change of Z by Drug over Doses", x = "Dose Number", y = "Cumulative Change in Z") +
  theme_minimal() +
  theme(legend.title = element_blank(), axis.text.x = element_text(hjust = 0.5)) +
  scale_x_continuous(breaks = dose_numbers) +
  geom_vline(xintercept = dose_numbers, linetype = "dotted", color = "gray")

# Display the plots
plot_cumulative_x
plot_cumulative_y
plot_cumulative_z

# Optionally, save the plots
ggsave("Cumulative_Rate_X_over_Doses.png", plot_cumulative_x, width = 10, height = 6, units = "in")
ggsave("Cumulative_Rate_Y_over_Doses.png", plot_cumulative_y, width = 10, height = 6, units = "in")
ggsave("Cumulative_Rate_Z_over_Doses.png", plot_cumulative_z, width = 10, height = 6, units = "in")