library(quantmod)
library(eurostat)
library(tidyverse)
library(data.table)

# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/frank")

calculate_quarterly_returns <- function(prices) {
  log(prices / lag(prices))
}

start_date <- Sys.Date() - years(30)
end_date <- Sys.Date()

# Energy Data: Gold, Silver, Oil, Gas
symbols_energy <- c("GC=F", "SI=F", "CL=F", "NG=F") # Gold, Silver, Oil, Natural Gas

library(quantmod)
library(dplyr)

# Initialize an empty xts object
aligned_energy_data <- NULL

# Define readable names for commodities
commodity_names <- c("Gold", "Silver", "Oil", "Gas")

# Process each symbol and merge into a single xts object
for (i in seq_along(symbols_energy)) {
  symbol <- symbols_energy[i]
  
  # Retrieve and process data for the symbol
  symbol_data <- getSymbols(symbol, src = "yahoo", from = start_date, to = end_date, auto.assign = FALSE) %>%
    na.locf() %>%  # Fill missing values
    to.quarterly(indexAt = "lastof", OHLC = FALSE) %>%
    Ad() %>%
    calculate_quarterly_returns()
  
  # Rename the column to a readable name
  colnames(symbol_data) <- commodity_names[i]
  
  # Merge the new symbol data into the aligned dataframe
  if (is.null(aligned_energy_data)) {
    aligned_energy_data <- symbol_data
  } else {
    aligned_energy_data <- merge(aligned_energy_data, symbol_data, all = TRUE)
  }
}

# Convert the xts object to a dataframe
aligned_energy_df <- data.frame(Date = index(aligned_energy_data), coredata(aligned_energy_data))



# Stock Data: DAX, Euro Stoxx, MSCI Europe
symbols_stocks <- c("^GDAXI", "^STOXX50E", "IEUR") # DAX, Euro Stoxx, MSCI Europe ETF
stocks_data <- list()

# Function to calculate quarterly returns and standardize column names
process_stock_data <- function(symbol) {
  data <- getSymbols(symbol, src = "yahoo", from = start_date, to = end_date, auto.assign = FALSE) %>%
    na.locf() %>%  # Fill missing values
    to.quarterly(indexAt = "lastof", OHLC = FALSE) %>%
    Ad() %>%
    calculate_quarterly_returns()
  
  # Convert to data frame and standardize column names
  data <- data.frame(Date = index(data), Return = coredata(data))
  data$Stock <- symbol  # Add Stock column for identification
  return(data)
}


# Combine all stock data
stock_list <- lapply(symbols_stocks, process_stock_data)

# Standardize column names for all stocks
stock_list <- lapply(stock_list, function(df) {
  colnames(df) <- c("Date", "Return", "Stock")
  return(df)
})

# Combine all stocks into a single data frame
stocks_data <- do.call(rbind, stock_list)

# Pivot the data wider so each stock is a separate column
stocks_wide <- stocks_data %>%
  
  pivot_wider(names_from = Stock, values_from = Return)

# View the reshaped data
print(head(stocks_wide))

# Cryptocurrency Data: Bitcoin
btc_data <- getSymbols("BTC-USD", src = "yahoo", from = start_date, to = end_date, auto.assign = FALSE) %>%
  na.locf() %>%  # Handle missing values
  to.quarterly(indexAt = "lastof", OHLC = FALSE) %>%
  Ad() %>%
  calculate_quarterly_returns()
btc_data <- data.frame(Date = index(btc_data), Return = coredata(btc_data))


real_estate_data_raw <- get_eurostat("prc_hpi_q", cache = FALSE)

real_estate_data <- real_estate_data_raw %>%
  filter(geo == "EU27_2020", purchase == "TOTAL", unit == "RCH_Q") %>%
  filter(TIME_PERIOD >= as.Date(start_date) & TIME_PERIOD <= as.Date(end_date)) %>%
  mutate(
    Quarter = paste(year(TIME_PERIOD), quarter(TIME_PERIOD), sep = "-"),
    Quarterly_Change = values / 100  # Convert percentages to proportions
  )

# Inflation Data: European Inflation (Eurostat)
euro_inflation <- get_eurostat("prc_hicp_midx", filters = list(geo = "EU27_2020", coicop = "CP00")) %>%
  filter(unit == "I15", time >= as.Date(start_date), time <= as.Date(end_date)) %>%
  mutate(Quarter = paste(year(time), quarter(time), sep = "-")) %>%
  group_by(Quarter) %>%
  summarise(Quarterly_Inflation = (values / lag(values)) - 1) %>%
  na.omit()

# Combine All Data
all_data <- list(
  Energy = energy_data,
  Stocks = stocks_data,
  Bitcoin = btc_data,
  RealEstate = real_estate_data,
  Inflation = euro_inflation
)

saveRDS(all_data, file = "inflation_hedging_data_updated.rds")

print("Data collection and processing complete.")

# Function to print structure of all data frames
print_dataframes_structure <- function() {
  # Get all objects in the global environment
  objs <- ls(envir = .GlobalEnv)
  
  # Iterate over objects and check if they are data frames
  for (obj in objs) {
    if (is.data.frame(get(obj))) {
      cat("Structure of", obj, ":\n")
      str(get(obj))
      cat("\n")  # Add spacing for readability
    }
  }
}

# Call the function
print_dataframes_structure()



btc_data <- btc_data %>%
  rename(Return = BTC.USD.Adjusted) %>%
  mutate(Quarter = paste(year(Date), quarter(Date), sep = "-"))

euro_inflation <- euro_inflation %>%
  group_by(Quarter) %>%
  summarise(Quarterly_Inflation = mean(Quarterly_Inflation, na.rm = TRUE)) %>%
  ungroup()

stocks_data <- stocks_data %>%
  mutate(Quarter = paste(year(Date), quarter(Date), sep = "-"))


# Combine data for correlation analysis with missing data handling
combined_data <- real_estate_data %>%
  select(Quarter, Real_Estate = Quarterly_Change) %>%
  full_join(
    btc_data %>% select(Quarter, Bitcoin = Return),
    by = "Quarter"
  ) %>%
  full_join(
    euro_inflation %>% select(Quarter, Inflation = Quarterly_Inflation),
    by = "Quarter"
  ) %>%
  full_join(
    stocks_wide %>%
      mutate(Quarter = paste(year(Date), quarter(Date), sep = "-")) %>%
      select(-Date),  # Remove 'Date' as it's redundant with 'Quarter'
    by = "Quarter"
  ) %>%
  full_join(
    aligned_energy_df %>%
      mutate(Quarter = paste(year(Date), quarter(Date), sep = "-")) %>%
      select(-Date),  # Remove 'Date' to avoid redundancy
    by = "Quarter"
  ) %>%
  arrange(Quarter)  # Order by Quarter

str(combined_data)


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

vars <- c("Real_Estate", "Bitcoin" ,   "Inflation",  "^GDAXI"  ,    "^STOXX50E",   "IEUR"      ,  "Gold"  ,      "Silver"    ,  "Oil"   ,      "Gas")

# Example usage with your dataframe and list of variables
df_descriptive_stats <- calculate_descriptive_stats(combined_data, vars)


library(ggplot2)
library(tidyr)

# Reshape combined_data to long format
combined_long <- combined_data %>%
  pivot_longer(
    cols = -Quarter,          # Keep Quarter as-is
    names_to = "Category",    # New column for variable names
    values_to = "Value"       # Values for each variable
  ) %>%
  drop_na(Value)              # Remove rows with NA values


# Fix Quarter to proper Date format
combined_long <- combined_long %>%
  mutate(Quarter = as.Date(paste0(substr(Quarter, 1, 4), "-", 
                                  (as.integer(substr(Quarter, 6, 6)) - 1) * 3 + 1, "-01")))

line_plot <- ggplot(combined_long, aes(x = Quarter, y = Value, color = Category, group = Category)) +
  geom_line() +
  labs(
    title = "Quarterly Changes Over Time",
    x = "Quarter",
    y = "Change"
  ) +
  facet_wrap(~ Category, scales = "free_y") +  # Each category gets its own chart
  theme_minimal() +
  theme(
    axis.text.x = element_text(hjust = 1),
    legend.position = "none"  # Remove legend (redundant with facets)
  )


heatmap_plot <- ggplot(combined_long, aes(x = Quarter, y = Category, fill = Value)) +
  geom_tile(color = "white") +  # Add borders for better readability
  scale_fill_gradient2(low = "red", mid = "white", high = "blue", midpoint = 0, name = "Change") +
  geom_text(aes(label = scales::percent(Value, accuracy = 0.1)), size = 3, angle = 90) +  # Add vertical percentage labels
  labs(
    title = "Heatmap of Quarterly Changes with Percentage Labels",
    x = "Quarter",
    y = "Category"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(hjust = 1),  # Rotate x-axis labels
    axis.text.y = element_text(size = 10),  # Adjust y-axis text size
    legend.position = "bottom"
  )


## BOXPLOTS OF MULTIPLE VARIABLES (vars as different columns)
library(ggplot2)
library(tidyr)
library(stringr)
library(dplyr)
library(ggrepel)

create_boxplots_with_outliers <- function(df, vars) {
  # Ensure that vars are in the dataframe
  df <- df[, c("Quarter", vars), drop = FALSE]  # Include Quarter for labeling
  
  # Reshape the data to a long format
  long_df <- df %>%
    pivot_longer(cols = -Quarter, names_to = "Variable", values_to = "Value")
  
  # Identify outliers
  long_df <- long_df %>%
    group_by(Variable) %>%
    mutate(
      Q1 = quantile(Value, 0.25, na.rm = TRUE),
      Q3 = quantile(Value, 0.75, na.rm = TRUE),
      IQR = Q3 - Q1,
      Lower = Q1 - 1.5 * IQR,
      Upper = Q3 + 1.5 * IQR,
      Is_Outlier = Value < Lower | Value > Upper
    )
  
  # Create side-by-side boxplots with outlier labels using ggrepel
  p <- ggplot(long_df, aes(x = Variable, y = Value)) +
    geom_boxplot() +
    stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
    geom_text_repel(
      data = filter(long_df, Is_Outlier), 
      aes(label = Quarter), 
      size = 3, 
      color = "blue",
      box.padding = 0.5,  # Space around labels
      max.overlaps = 10   # Limit maximum number of overlaps
    ) +
    labs(x = "Variable", y = "Value") +
    theme(axis.text.x = element_text(hjust = 1, vjust = 1)) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 10))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = "side_by_side_boxplots_with_outliers.png", plot = p, width = 10, height = 6)
}

# Call the function
create_boxplots_with_outliers(combined_data, vars)


# Outlier Assessment

library(car)

# Check Multicollinearity
check_multicollinearity <- function(df, dependent_var, ivs) {
  # Ensure all variables exist in the dataframe
  df <- df[, c(dependent_var, ivs), drop = FALSE]
  
  # Exclude rows with missing values
  df <- na.omit(df)
  
  # Fit the linear model
  formula <- as.formula(paste(dependent_var, "~", paste(ivs, collapse = " + ")))
  lm_model <- lm(formula, data = df)
  
  # Calculate VIF
  vif_values <- vif(lm_model)
  
  # Return VIF values as a dataframe
  return(data.frame(Variable = names(vif_values), VIF = vif_values))
}


# Clean variable names automatically
#colnames(combined_data) <- make.names(colnames(combined_data), unique = TRUE)
colnames(combined_data)

# Add Inflation to the list if needed
dependent_var <- "Inflation"  # Define the dependent variable
ivs <- c("Bitcoin", "X.GDAXI", "X.STOXX50E", "IEUR", "Gold", "Silver", "Oil", "Gas", "Real_Estate")

# Run VIF check
vif_results <- check_multicollinearity(combined_data, dependent_var, ivs)

# Print the results
print("Variance Inflation Factor (VIF) Results:")
print(vif_results)

ivs2 <- c("Bitcoin", "X.GDAXI", "IEUR"      , "Gold", "Silver", "Oil", "Gas", "Real_Estate")

# Run VIF check
vif_results2 <- check_multicollinearity(combined_data, dependent_var, ivs2)

# Print the results
print("Variance Inflation Factor (VIF) Results:")
print(vif_results2)


calculate_correlations <- function(df, vars, method = "pearson") {
  # Ensure only the selected variables are in the dataframe
  df <- df[vars]
  
  # Calculate pairwise correlations
  correlation_matrix <- cor(df, use = "pairwise.complete.obs", method = method)
  
  # Return the correlation matrix
  return(correlation_matrix)
}

vars <- c("Real_Estate", "Bitcoin" ,   "Inflation",  "X.GDAXI"  ,    "X.STOXX50E",   "IEUR"      ,  "Gold"  ,      "Silver"    ,  "Oil"   ,      "Gas")


correlation_matrix <- calculate_correlations(combined_data, vars, method = "pearson")
print("Correlation Matrix:")
print(correlation_matrix)



library(lmtest)

# Fit a simple linear model
lm_model <- lm(Inflation ~ Bitcoin + X.GDAXI + X.STOXX50E + IEUR + Gold + Silver + Oil + Gas + Real_Estate, 
               data = combined_data)

# Perform Durbin-Watson Test
dw_test <- dwtest(lm_model)
print(dw_test)


# OLS Regression

library(broom)

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
  data = combined_data,
  predictors = ivs2,  # Replace with actual predictor names
  response_vars = "Inflation",       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults <- bind_rows(ols_results_list_both_groups)

ivs3 <- c("X.GDAXI",  "Gold", "Silver", "Oil", "Gas", "Real_Estate")

ols_results_list_nobtc <- fit_ols_and_format(
  data = combined_data,
  predictors = ivs3,  # Replace with actual predictor names
  response_vars = "Inflation",       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults_nobtc <- bind_rows(ols_results_list_nobtc)


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
  "Full Data" = combined_data, 
  "Descriptive Stats" = df_descriptive_stats, 
  "Correlation" = correlation_matrix,
  "VIF Results" = vif_results,
  "Model" = df_modelresults,
  "Model - No BTC" = df_modelresults_nobtc
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")

# Save the Line Plot
ggsave("quarterly_changes_line_plot.png", plot = line_plot, width = 10, height = 6)
ggsave("quarterly_changes_heatmap.png", plot = heatmap_plot, width = 18, height = 10)
