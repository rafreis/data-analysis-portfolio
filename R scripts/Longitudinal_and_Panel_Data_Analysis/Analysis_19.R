setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/judgemaq")

library(readxl)
df <- read_xlsx('CleanedData.xlsx')

names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))

# Handle Missing Data

# Load necessary libraries
library(dplyr)
library(ggplot2)

# Filter out the data for Australia and for the relevant years
aus_data <- df %>% filter(Country == "AUS", Year >= 1960, Year <= 2019)

# Fit a linear model to the existing data
model <- lm(Capital_Stock_N_Bn_USD ~ Year, data = aus_data)

# Predict values for 2020, 2021, and 2022
predicted_values <- predict(model, newdata = data.frame(Year = c(2020, 2021, 2022)))

# Update the dataframe with the predicted values
for (year in c(2020, 2021, 2022)) {
  df <- df %>% 
    mutate(Capital_Stock_N_Bn_USD = ifelse(Country == "AUS" & Year == year, predicted_values[year - 2019], Capital_Stock_N_Bn_USD))
}

# Checking the interpolated values for AUS for 2020-2022
df %>% filter(Country == "AUS", Year >= 2018, Year <= 2022) %>% 
  select(Year, Capital_Stock_N_Bn_USD)

# Custom function to manually create a lagged column
manual_lag <- function(x) {
  c(NA, head(x, -1))
}

# Sort the data by Country and Year
df <- df %>%
  arrange(Country, Year)

# Apply the custom lag function within each group
df <- df %>%
  group_by(Country) %>%
  mutate(
    Lagged_GDP = manual_lag(GDP),
    GDP_Growth = ifelse(is.na(Lagged_GDP), NA, (GDP - Lagged_GDP) / Lagged_GDP * 100)
  ) %>%
  ungroup() # Ungroup for further operations

# Cobb-Douglas function to estimate expected GDP

# Load necessary libraries
library(tidyr)

# Log-transforming the data
df <- df %>%
  mutate(Log_GDP = log(GDP),
         Log_Capital = log(Capital_Stock_N_Bn_USD),
         Log_Labor = log(Labour_Force_Thousands))

# Function to perform regression and calculate TFP, Alpha, Beta for each country
calculate_params <- function(data) {
  model <- lm(Log_GDP ~ Log_Capital + Log_Labor, data = data)
  alpha <- coef(model)["Log_Capital"]
  beta <- coef(model)["Log_Labor"]
  A <- exp(coef(model)["(Intercept)"])
  
  data$TFP <- A
  data$Alpha <- alpha
  data$Beta <- beta
  
  data$Expected_GDP <- data$TFP * (data$Capital_Stock_N_Bn_USD ^ data$Alpha) * (data$Labour_Force_Thousands ^ data$Beta)
  data$Output_Gap <- data$GDP - data$Expected_GDP
  
  return(data)
}

# Apply the function to each country separately
df_aus <- calculate_params(df %>% filter(Country == "AUS"))
df_us <- calculate_params(df %>% filter(Country == "US"))

# Combine the results back into a single dataframe
df_combined <- bind_rows(df_aus, df_us)


#Plot Results
library(scales)

# Function to plot and save the graph
plot_and_save <- function(data, country, x_var, y_var1, y_var2, file_name) {
  # Filter data for the specified country
  country_data <- data %>% filter(Country == country)
  
  # Reshape data for plotting
  plot_data <- country_data %>%
    select({{x_var}}, {{y_var1}}, {{y_var2}}) %>%
    pivot_longer(cols = c({{y_var1}}, {{y_var2}}), names_to = "Variable", values_to = "Value")
  
  # Plot
  p <- ggplot(plot_data, aes_string(x = rlang::as_string(ensym(x_var)), y = "Value", color = "Variable")) +
    geom_line() +
    scale_y_continuous(labels = scales::label_number()) +  # No scientific notation
    scale_x_continuous(breaks = scales::pretty_breaks(n = 10)) + # Adjust number of X-axis labels
    theme_minimal() +
    labs(x = rlang::as_string(ensym(x_var)), y = "Value", color = "Variable") +
    ggtitle(paste("Plot of", rlang::as_string(ensym(y_var1)), "and", rlang::as_string(ensym(y_var2)), "for", country))
  
  print(p)
  
  # Save plot
  ggsave(file_name, plot = p, width = 10, height = 6)
}

# Example usage
plot_and_save(df_combined, "AUS", Year, GDP, Expected_GDP, "GDP_vs_Expected_GDP_AUS.png")
plot_and_save(df_combined, "US", Year, GDP, Expected_GDP, "GDP_vs_Expected_GDP_US.png")



# Difference-in-Difference analysis

library(plm)
library(broom)
library(sandwich)
library(lmtest)

# Convert the data frame to a pdata.frame (panel data frame)
df_panel <- pdata.frame(df_combined, index = c("Country", "Year"))
df_panel$Year <- as.numeric(as.character(df_panel$Year))

# Function to run fixed effects model and return tidy results and models
run_fixed_effects_pandemic <- function(data, dvs, controls = NULL) {
  results <- list()
  models <- list()
  
  for (dv in dvs) {
    formula <- as.formula(paste(dv, "~ Pandemic_Period * Country_Indicator", if (!is.null(controls)) paste("+", controls)))
    model <- plm(formula, data = data, model = "within")
    
    # Applying robust standard errors to address autocorrelation
    coeftest(model, vcov = vcovHC(model, type = "HC1"))
    
    tidy_model <- tidy(model, robust = vcovHC(model, type = "HC1"))
    results[[dv]] <- tidy_model
    models[[dv]] <- model
  }
  
  # Combine results into a single dataframe
  combined_results <- do.call(rbind, results)
  # Adding a column for dependent variables
  combined_results$DV <- rep(dvs, each = nrow(combined_results) / length(dvs))
  
  return(list("results" = combined_results, "models" = models))
}

#Log-transformation

log_transform_all_except <- function(data, exclude_cols) {
  # Identifying all columns except the ones to exclude
  dvs <- setdiff(names(data), exclude_cols)
  
  for (dv in dvs) {
    # Creating a new column name for the log-transformed variable
    new_col_name <- paste("log", dv, sep = "_")
    
    # Applying the log transformation
    # Adding 1 to avoid log(0) which is undefined
    data[[new_col_name]] <- log(data[[dv]] + 1)
  }
  return(data)
}

# Exclude 'Year' and 'Country' columns
exclude_cols <- c("Year", "Country", "Country_Indicator", "GDP_Growth")

# Applying the function to your data
df_panel_transformed <- log_transform_all_except(df_panel, exclude_cols)


# Add a pandemic indicator (0 for pre-pandemic, 1 for post-pandemic)
df_panel_transformed <- df_panel_transformed  %>%
  mutate(Pandemic_Period = ifelse(Year >= 2020, 1, 0),
         Country_Indicator = ifelse(Country == "US", 1, 0)) # assuming 1 for US, 0 for AUS

log_dvs <- c("log_Unemployed_Thousands", "log_Unemp_Rate_Percent", "log_CPI_ex_FE_percentage", "log_GDP", "log_Output_Gap","GDP_Growth")

# Run the function without control variables
did_results <- run_fixed_effects_pandemic(df_panel_transformed, log_dvs)

# Accessing the results
pandemic_model_results <- did_results$results

# Accessing the models (for diagnostics, etc.)
pandemic_models <- did_results$models

# Example to access a specific model
gdp_model <- pandemic_models[["log_GDP"]]
outputgap_model <- pandemic_models[["log_Output_Gap"]]
unemployment_model <- pandemic_models[["log_Unemployed_Thousands"]]

# Breusch-Pagan Test for Homoscedasticity
library(lmtest)

# Apply the Breusch-Pagan test
bptest(gdp_model, ~ fitted(gdp_model), data = df_panel_transformed)

# Wooldridge Test for Autocorrelation in Panel Data
pbgtest(gdp_model)

library(car)

# Extract residuals from the plm model
gdp_model_resid <- residuals(gdp_model)
unemp_model_resid <- residuals(unemployment_model)

# Function to generate and save QQ plots
generate_qq_plots <- function(models, dv_names) {
  for (dv in dv_names) {
    # Extract the model
    model <- models[[dv]]
    
    # Check if the model exists
    if (!is.null(model)) {
      # Extract residuals
      model_resid <- residuals(model)
      
      # Open a new PNG file to save the plot
      png(filename = paste0(dv, "_qq_plot.png"), width = 700, height = 700)
      
      # Generate QQ plot
      qqPlot(model_resid, main = paste("QQ Plot of", dv, "Model Residuals"), 
             ylab = "Residuals", xlab = "Theoretical Quantiles")
      
      # Close the PNG device
      dev.off()
    }
  }
}

# Generate and save QQ plots in the working directory
generate_qq_plots(pandemic_models, log_dvs)


library(ggplot2)
library(dplyr)

# Line Plots
plot_time_series <- function(data, x_var, category_var, y_var) {
  # Ensure variables are correctly converted for plotting
  x_var <- enquo(x_var)
  category_var <- enquo(category_var)
  y_var <- enquo(y_var)
  
  # Filter out rows with NA in the specified variable
  filtered_data <- data %>% 
    filter(!is.na(!!y_var))
  
  # Prepare data
  plot_data <- filtered_data %>%
    select(!!x_var, !!category_var, !!y_var) %>%
    group_by(!!x_var, !!category_var) %>%
    summarize(mean_value = mean(!!y_var, na.rm = TRUE), .groups = 'drop')
  
  # Create the line plot with rotated x-axis labels
  ggplot(plot_data, aes(x = !!x_var, y = mean_value, color = !!category_var, group = !!category_var)) +
    geom_line() +
    labs(x = as.character(rlang::as_label(x_var)), 
         y = as.character(rlang::as_label(y_var)), 
         color = as.character(rlang::as_label(category_var))) +
    ggtitle(paste("Time Series of", as.character(rlang::as_label(y_var)), "by", as.character(rlang::as_label(category_var)))) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}

# Example usage
plot_time_series(df_combined, Year, Country, Output_Gap)

vars <- c("log_Unemployed_Thousands", "log_Unemp_Rate_Percent", "log_CPI_ex_FE_percentage", "log_GDP", "Output_Gap")

# Looping through each DV to create and save plots
for (var in vars) {
  
  # Convert the string to a symbol
  y_var_sym <- sym(var)
  
  # Generate the plot using !! to unquote the symbol
  plot <- plot_time_series(df_panel_transformed, Year, Country, !!y_var_sym)
  
  # Display the plot
  print(plot)
  
  
}

## CORRELATION ANALYSIS

calculate_correlation_matrix <- function(data, variables, method = "pearson") {
  # Ensure the method is valid
  if (!method %in% c("pearson", "spearman")) {
    stop("Method must be either 'pearson' or 'spearman'")
  }
  
  # Subset the data for the specified variables
  data_subset <- data[variables]
  
  # Calculate the correlation matrix
  corr_matrix <- cor(data_subset, method = method, use = "complete.obs")
  
  # Initialize a matrix to hold formatted correlation values
  corr_matrix_formatted <- matrix(nrow = nrow(corr_matrix), ncol = ncol(corr_matrix))
  rownames(corr_matrix_formatted) <- colnames(corr_matrix)
  colnames(corr_matrix_formatted) <- colnames(corr_matrix)
  
  # Calculate p-values and apply formatting with stars
  for (i in 1:nrow(corr_matrix)) {
    for (j in 1:ncol(corr_matrix)) {
      if (i == j) {  # Diagonal elements (correlation with itself)
        corr_matrix_formatted[i, j] <- format(round(corr_matrix[i, j], 3), nsmall = 3)
      } else {
        p_value <- cor.test(data_subset[[i]], data_subset[[j]], method = method)$p.value
        stars <- ifelse(p_value < 0.01, "***", 
                        ifelse(p_value < 0.05, "**", 
                               ifelse(p_value < 0.1, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}

vars <- c("log_Wages_Salaries"      ,            "log_WPI"        ,                    "log_BOP",
          "log_Capital_Stock_N_Bn_USD"         ,  "log_Exports_N_Bn_USD"       ,          "log_Imports_N_Bn_USD",
          "log_Labour_Force_Thousands" , "log_Unemp_Rate_Percent",
          "log_PCE_deflator_percent"      ,       "log_PCE_less_food_and_energy_percent","log_CPI_percentage", "Pandemic_Period")

correlation_matrix <- calculate_correlation_matrix(df_panel_transformed, vars, method = "pearson")



## MODEL TO EVALUATE OUTPUT GAP

# Function to run fixed effects model and return tidy results and models
library(plm)
library(lmtest)

run_fixed_effects <- function(data, dv, ivs, country = NULL) {
  # Adjusting the dataset based on the country if specified
  if (!is.null(country)) {
    data <- subset(data, Country == country)
  }
  
  # Create the formula
  formula <- as.formula(paste(dv, "~", paste(ivs, collapse = " + "), if (is.null(country)) "+ Country_Indicator"))
  
  # Run the model
  model <- plm(formula, data = data, model = "within")
  
  # Applying robust standard errors
  tidy_model <- tidy(model, robust = vcovHC(model, type = "HC1"))
  
  return(list("results" = tidy_model, "model" = model))
}

# Define your Independent Variables
ivs <- c(                                      
         "log_Capital_Stock_N_Bn_USD"   ,        "log_Exports_N_Bn_USD"      ,           "log_Imports_N_Bn_USD",
         "log_Labour_Force_Thousands" , "log_Unemp_Rate_Percent",
         "log_CPI_percentage", "Pandemic_Period")
         

# Running the function for each scenario
results_AUS <- run_fixed_effects(df_panel_transformed, "log_Output_Gap", ivs, "AUS")
results_US <- run_fixed_effects(df_panel_transformed, "log_Output_Gap", ivs, "US")

df_results_AUS <- results_AUS$results
df_results_US <- results_US$results

ivs2 <- c("Country_Indicator", "log_Capital_Stock_N_Bn_USD"   ,        "log_Exports_N_Bn_USD"      ,           "log_Imports_N_Bn_USD",
         "log_Labour_Force_Thousands" , "log_Unemp_Rate_Percent",
         "log_CPI_percentage", "Pandemic_Period")

results_Both <- run_fixed_effects(df_panel_transformed, "log_Output_Gap", ivs2)
df_results_Both <- results_Both$results


#QQ plots

library(car)
# Plotting QQ plot for the Australian model
qq_plot_AUS <- autoplot(qqnorm(residuals(results_AUS$model)))

# Saving the plot
ggsave("qq_plot_AUS.png", qq_plot_AUS)

# Similarly, you can create QQ plots for the US and Both models
qq_plot_US <- autoplot(qqnorm(residuals(results_US$model)))
ggsave("qq_plot_US.png", qq_plot_US)

qq_plot_Both <- autoplot(qqnorm(residuals(results_Both$model)))
ggsave("qq_plot_Both.png", qq_plot_Both)








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
  "Cleaned Data" = df_combined,
  "DiD Analysis - Pandemic" = pandemic_model_results,
  "Correlations" = correlation_matrix,
  "OUtput Gap - US" = df_results_US,
  "OUtput Gap - AUS" = df_results_AUS,
  "OUtput Gap - Combined" = df_results_Both
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Results.xlsx")









# Function to plot graph for a single country

plot_time_series_country <- function(data, x_var, category_var, y_var, countries) {
  # Ensure variables are correctly converted for plotting
  x_var <- enquo(x_var)
  category_var <- enquo(category_var)
  y_var <- enquo(y_var)
  
  # Filter out rows with NA in the specified variable and for specified countries
  filtered_data <- data %>% 
    filter(!is.na(!!y_var), !!category_var %in% countries)
  
  # Prepare data
  plot_data <- filtered_data %>%
    select(!!x_var, !!category_var, !!y_var) %>%
    group_by(!!x_var, !!category_var) %>%
    summarize(mean_value = mean(!!y_var, na.rm = TRUE), .groups = 'drop')
  
  # Create the line plot with rotated x-axis labels
  ggplot(plot_data, aes(x = !!x_var, y = mean_value, color = !!category_var, group = !!category_var)) +
    geom_line() +
    labs(x = as.character(rlang::as_label(x_var)), 
         y = as.character(rlang::as_label(y_var)), 
         color = as.character(rlang::as_label(category_var))) +
    ggtitle(paste("Time Series of", as.character(rlang::as_label(y_var)))) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}


# Example usage for Australia and the United States
plot_time_series_country(df_combined, Year, Country, Output_Gap, countries = c("US"))
