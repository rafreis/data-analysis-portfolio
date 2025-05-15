setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/SPOTIO")

df_overall <- read.csv('Data_ROI_Overall.csv')
df_perCustomer <- read.csv('Data_ROI_perCustomer.csv')

# Get rid of special characters

names(df_overall) <- gsub(" ", "_", names(df_overall))
names(df_overall) <- gsub("\\(", "_", names(df_overall))
names(df_overall) <- gsub("\\)", "_", names(df_overall))
names(df_overall) <- gsub("\\-", "_", names(df_overall))
names(df_overall) <- gsub("/", "_", names(df_overall))
names(df_overall) <- gsub("\\\\", "_", names(df_overall)) 
names(df_overall) <- gsub("\\?", "", names(df_overall))
names(df_overall) <- gsub("\\'", "", names(df_overall))
names(df_overall) <- gsub("\\,", "_", names(df_overall))

names(df_perCustomer) <- gsub(" ", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\(", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\)", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\-", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("/", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\\\", "_", names(df_perCustomer)) 
names(df_perCustomer) <- gsub("\\?", "", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\'", "", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\,", "_", names(df_perCustomer))

# New dataset with averages

library(dplyr)

# Calculate the average activities per month
df_avg_per_month <- df_perCustomer %>%
  group_by(Month.Number) %>%
  summarize(Average_Activity = mean(Count.of.Activity.Type), Average_perSalesRep = mean(Average.Count.per.SalesRep), Average_InPerson = mean(ActivityCountinPerson))

# Inspect Plot

library(ggplot2)

ggplot(df_overall, aes(x = Month.Number, y = CountActivityPerUserByMonth)) +
  geom_point() +
  labs(x = "Month Number", y = "Average Activity", title = "Plot of Average Activities per SalesRep per Month")

ggplot(df_overall, aes(x = Month.Number, y = CountActivityPerUserByMonth)) +
  geom_col() +  # Creates the bar plot
  geom_text(aes(label = round(CountActivityPerUserByMonth, 0)), # Adds rounded data labels
            position = position_stack(vjust = 1.05), # Adjusts the vertical position
            size = 3) + # Adjusts the text size
  labs(x = "Month Number", y = "Average Activity", title = "Plot of Average Activities per SalesRep per Month")


# Calculation of growth Rates

# Ensure the columns are numeric
df_overall$Average.Count.per.Customer <- as.numeric(df_overall$Average.Count.per.Customer)
df_overall$CountActivityPerUserByMonth <- as.numeric(df_overall$CountActivityPerUserByMonth)

# Calculate Monthly Growth Rates
df_overall <- df_overall %>%
  mutate(
    Growth_AvgCountPerCustomer = (Average.Count.per.Customer - lag(Average.Count.per.Customer)) / lag(Average.Count.per.Customer) * 100,
    Growth_AvgCountPerSalesRep = (CountActivityPerUserByMonth - lag(CountActivityPerUserByMonth)) / lag(CountActivityPerUserByMonth) * 100
  )

# Remove NA values that result from the lag in the first row
df_overall <- na.omit(df_overall)

# View a few rows to check the growth rates
head(df_overall)

# Calculate a Fixed Expected Growth Rate
fixed_growth_rate <- df_overall %>%
  summarise(
    FixedGrowthRate_Customer = mean(Growth_AvgCountPerCustomer, na.rm = TRUE),
    FixedGrowthRate_SalesRep = mean(Growth_AvgCountPerSalesRep, na.rm = TRUE)
  )

fixed_growth_rate

#Calculate Proportion of In-Person
df_avg_per_month <- df_avg_per_month %>%
  mutate(Proportion_InPerson = (Average_InPerson / Average_Activity) * 100)

# Predictive Model

# Creation of Dummy variables

library(dplyr)
library(lmtest)
library(broom)

df_perCustomer <- df_perCustomer %>%
  mutate_at(vars(Customer), as.factor) %>%
  mutate(CustomerID = as.numeric(as.factor(Customer)))  # Convert 'Customer' to a numeric ID

# Create dummy variables for each customer
df_perCustomer_dummies <- model.matrix(~ Customer - 1, data = df_perCustomer)

df_perCustomer <- cbind(df_perCustomer, df_perCustomer_dummies)

#Clean names

names(df_perCustomer) <- gsub(" ", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\(", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\)", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\-", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("/", "_", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\\\", "_", names(df_perCustomer)) 
names(df_perCustomer) <- gsub("\\?", "", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\'", "", names(df_perCustomer))
names(df_perCustomer) <- gsub("\\,", "_", names(df_perCustomer))


# Model for Total Activities per SalesRep

library(lme4)
library(broom)

model_activities_persalesrep <- lm(CountActivityPerUserByMonth ~ Month.Number, data = df_overall)
summary(model_activities_persalesrep)

# use 'tidy' from broom package for a cleaner output
model_activities_avg_tidy <- tidy(model_activities_persalesrep)
model_activities_avg_tidy

# Calculate residuals and fitted values
residuals <- resid(model_activities_persalesrep)
fitted_values <- fitted(model_activities_persalesrep)

# Create a dataframe for plotting
model_df_activitiespersalesrep <- data.frame(Residuals = residuals, Fitted = fitted_values, Month = df_overall$Month.Number)

# Residuals vs Month Number Plot
ggplot(model_df_activitiespersalesrep, aes(x = Month, y = Residuals)) +
  geom_point() +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
  labs(x = "Month Number", y = "Residuals", title = "Residuals vs Month Number")


library(alr4)

# Fit a simple linear regression model
model_linear <- lm(CountActivityPerUserByMonth ~ Month.Number, data = df_overall)

# Fit a more complex model, such as a polynomial model
model_poly <- lm(CountActivityPerUserByMonth ~ poly(Month.Number, 2, raw=TRUE), data = df_overall)

# Use anova to compare the two models
anova_linear_poly <- anova(model_linear, model_poly)
anova_linear_poly

library(gvlma)

# Perform the global validation test
gvmodel <- gvlma(model_linear)
summary(gvmodel) # Model is linear!

# Calculate predictions based on the model
new_data <- data.frame(Month.Number = 1:34)
prediction <- predict(model_activities_persalesrep, newdata = new_data, interval = "confidence")

# Extract the predicted values and confidence intervals
df_overall$Predicted_Activity <- prediction[, "fit"]
df_overall$Predicted_Activity_Lower <- prediction[, "lwr"]
df_overall$Predicted_Activity_Upper <- prediction[, "upr"]


# Calculate the growth rate percentage (Month over Month)
# Add NA at the beginning of the diff() vector to align the lengths
growth_rates <- c(NA, diff(df_overall$Predicted_Activity))

# The growth rate is the change in predicted activity divided by the previous month's predicted activity
df_overall$Growth_Rate_Percentage <- growth_rates / lag(df_overall$Predicted_Activity, default = first(df_overall$Predicted_Activity)) * 100

# Replace any potential NaN (resulting from division by zero) with zero
df_overall$Growth_Rate_Percentage[is.nan(df_overall$Growth_Rate_Percentage)] <- 0

df_overall <- df_overall %>%
  mutate(Growth_Rate_Percentage = ifelse(is.infinite(Growth_Rate_Percentage) | is.nan(Growth_Rate_Percentage), 0, Growth_Rate_Percentage))

# Calculate Compound Monthly Growth Rate for a given period
# CMGR = (Ending Value / Beginning Value)^(1 / Number of Months) - 1
calculate_cmgr <- function(df, start_month, end_month) {
  begin_value <- df$Predicted_Activity[df$Month.Number == start_month]
  end_value <- df$Predicted_Activity[df$Month.Number == end_month]
  num_months <- end_month - start_month
  cmgr <- (end_value / begin_value)^(1 / num_months) - 1
  return(cmgr * 100)
}

# calculate CMGR from Month 1 to Month 6
cmgr_1_to_6 <- calculate_cmgr(df_overall, 1, 6)
cmgr_1_to_6

# calculate CMGR from Month 1 to Month 12
cmgr_1_to_12 <- calculate_cmgr(df_overall, 1, 12)
cmgr_1_to_12

# calculate CMGR from Month 12 to Month 24
cmgr_12_to_24 <- calculate_cmgr(df_overall, 12, 24)
cmgr_12_to_24

# calculate CMGR from Month 1 to Month 24
cmgr_12_to_24 <- calculate_cmgr(df_overall, 12, 24)
cmgr_12_to_24

# calculate CMGR from Month 1 to Month 34
cmgr_1_to_34 <- calculate_cmgr(df_overall, 1, 34)
cmgr_1_to_34


# We offset the predicted values by 1 to align with the 'diff' output
df_overall <- df_overall %>%
  mutate(
    Growth_Rate_Percentage = c(NA, diff(Predicted_Activity) / head(Predicted_Activity, -1) * 100),
    
  )

# Replace any potential NaN or infinite values with zero
df_overall$Growth_Rate_Percentage[is.nan(df_overall$Growth_Rate_Percentage) | is.infinite(df_overall$Growth_Rate_Percentage)] <- 0


# Function to calculate CMGR with confidence intervals
calculate_cmgr_with_ci <- function(df, start_month, end_month) {
  begin_value <- df$Predicted_Activity[df$Month.Number == start_month]
  end_value <- df$Predicted_Activity[df$Month.Number == end_month]
  begin_value_lower <- df$Predicted_Activity_Lower[df$Month.Number == start_month]
  end_value_lower <- df$Predicted_Activity_Lower[df$Month.Number == end_month]
  begin_value_upper <- df$Predicted_Activity_Upper[df$Month.Number == start_month]
  end_value_upper <- df$Predicted_Activity_Upper[df$Month.Number == end_month]
  
  num_months <- end_month - start_month
  cmgr <- (end_value / begin_value)^(1 / num_months) - 1
  cmgr_lower <- (end_value_lower / begin_value_lower)^(1 / num_months) - 1
  cmgr_upper <- (end_value_upper / begin_value_upper)^(1 / num_months) - 1
  
  return(list(CMGR = cmgr * 100, Lower_CI = cmgr_lower * 100, Upper_CI = cmgr_upper * 100))
}

# Calculate CMGR with CI from Month 1 to Month 12
cmgr_1_to_12_ci <- calculate_cmgr_with_ci(df_overall, 1, 12)
cmgr_1_to_12_ci

# Get mean and SD of Activities per Win per User

# Calculate mean of ActivitiesPerWonPerUser
mean_ActivitiesPerWonPerUser <- mean(df_perCustomer$ActivitiesPerWonPerUser, na.rm = TRUE)

# Calculate standard deviation of ActivitiesPerWonPerUser
sd_ActivitiesPerWonPerUser <- sd(df_perCustomer$ActivitiesPerWonPerUser, na.rm = TRUE)

# Monte Carlo Simulation for Conversion

monte_carlo_simulation <- function(predicted_activity, mean_percent, sd_percent, num_simulations) {
  # Simulate proportions
  simulations <- rnorm(num_simulations, mean = mean_percent, sd = sd_percent) / 100
  # Calculate closed deals
  closed_deals <- predicted_activity * simulations
  # Return lower, median, and upper limits
  return(c(lower = quantile(closed_deals, probs = 0.025), 
           median = median(closed_deals),
           upper = quantile(closed_deals, probs = 0.975)))
}

# Parameters for Monte Carlo
mean_percent <- 1.0196    # Estimation based on 217,584 activities that generated 2,134 wins
sd_percent <- 1    # 1%
num_simulations <- 1000  # Number of simulations

# Apply the Monte Carlo simulation
df_overall <- df_overall %>%
  rowwise() %>%
  mutate(
    Closed_Deals_Lower = monte_carlo_simulation(Predicted_Activity, mean_percent, sd_percent, num_simulations)["lower.2.5%"],
    Closed_Deals_Median = monte_carlo_simulation(Predicted_Activity, mean_percent, sd_percent, num_simulations)["median"],
    Closed_Deals_Upper = monte_carlo_simulation(Predicted_Activity, mean_percent, sd_percent, num_simulations)["upper.97.5%"]
  ) %>%
  ungroup()

# Sales/Win estimation

salesperwin <- 13355000/2134

# Monte Carlos Simulation for Revenue

monte_carlo_simulation_revenue <- function(closed_deals, sale_mean, sale_sd, num_simulations) {
  # Simulate sale values
  sale_values_simulations <- rnorm(num_simulations, mean = sale_mean, sd = sale_sd)
  # Calculate revenue
  revenue_simulations <- closed_deals * sale_values_simulations
  # Return lower, median, and upper limits
  return(c(lower = quantile(revenue_simulations, probs = 0.025),
           median = median(revenue_simulations),
           upper = quantile(revenue_simulations, probs = 0.975)))
}


# Parameters for sale value Monte Carlo
sale_mean <- salesperwin  # Example mean sale value
sale_sd <- 2000     # Example standard deviation of sale value

# Apply the Monte Carlo simulation for revenue
df_overall <- df_overall %>%
  rowwise() %>%
  mutate(
    Revenue_Lower = monte_carlo_simulation_revenue(Closed_Deals_Median, sale_mean, sale_sd, num_simulations)["lower.2.5%"],
    Revenue_Median = monte_carlo_simulation_revenue(Closed_Deals_Median, sale_mean, sale_sd, num_simulations)["median"],
    Revenue_Upper = monte_carlo_simulation_revenue(Closed_Deals_Median, sale_mean, sale_sd, num_simulations)["upper.97.5%"]
  ) %>%
  ungroup()


# Calculate cumulative sales
df_overall <- df_overall %>%
  mutate(
    Cumulative_Sales_Lower = cumsum(Revenue_Lower),
    Cumulative_Sales_Median = cumsum(Revenue_Median),
    Cumulative_Sales_Upper = cumsum(Revenue_Upper)
  )

# Define the annual fee
annual_fee <- 16560  

# Function to calculate ROI
calculate_roi <- function(cumulative_sales, annual_fee, months) {
  # Extract the cumulative sales at the end of the year
  end_of_year_sales <- cumulative_sales[months]
  # Calculate ROI
  roi <- (end_of_year_sales - annual_fee) / annual_fee * 100
  return(roi)
}

# Calculate ROI for years 1, 2, and 3
df_overall$ROI_Year1_Lower <- calculate_roi(df_overall$Cumulative_Sales_Lower, annual_fee, 12)
df_overall$ROI_Year2_Lower <- calculate_roi(df_overall$Cumulative_Sales_Lower, annual_fee, 24)


df_overall$ROI_Year1_Median <- calculate_roi(df_overall$Cumulative_Sales_Median, annual_fee, 12)
df_overall$ROI_Year2_Median <- calculate_roi(df_overall$Cumulative_Sales_Median, annual_fee, 24)


df_overall$ROI_Year1_Upper <- calculate_roi(df_overall$Cumulative_Sales_Upper, annual_fee, 12)
df_overall$ROI_Year2_Upper <- calculate_roi(df_overall$Cumulative_Sales_Upper, annual_fee, 24)


#CLosed Deals Graph

library(ggplot2)

# Create the bar graph
ggplot(df_overall, aes(x = Month.Number, y = Closed_Deals_Median)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  labs(title = "Estimated Closed Deals per Sales Rep by Month",
       x = "Month",
       y = "Estimated Closed Deals",
       caption = "Data represents the median estimated closed deals.") +
  theme_minimal()

# Create the bar graph
ggplot(df_overall, aes(x = Month.Number, y = Revenue_Median)) +
  geom_bar(stat = "identity", fill = "darkgreen") +
  labs(title = "Estimated Revenue per Sales Rep by Month",
       x = "Month",
       y = "Estimated revenue",
       caption = "Data represents the median estimated closed deals.") +
  theme_minimal()


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
  "Cleaned Data" = df_overall
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

