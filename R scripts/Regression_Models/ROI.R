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

ggplot(df_avg_per_month, aes(x = Month.Number, y = Average_Activity)) +
  geom_point() +
  labs(x = "Month Number", y = "Average Activity", title = "Plot of Average Activities per Month")


# Calculation of growth Rates

# Ensure the columns are numeric
df_overall$Average.Count.per.Customer <- as.numeric(df_overall$Average.Count.per.Customer)
df_overall$Average.Count.per.SalesRep <- as.numeric(df_overall$Average.Count.per.SalesRep)

# Calculate Monthly Growth Rates
df_overall <- df_overall %>%
  mutate(
    Growth_AvgCountPerCustomer = (Average.Count.per.Customer - lag(Average.Count.per.Customer)) / lag(Average.Count.per.Customer) * 100,
    Growth_AvgCountPerSalesRep = (Average.Count.per.SalesRep - lag(Average.Count.per.SalesRep)) / lag(Average.Count.per.SalesRep) * 100
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


# Model for Total Activities (with Dummy customers)
library(lme4)
library(broom)

model_activities <- lm(Count.of.Activity.Type ~ Month.Number + CustomerCustomer1_activities_encrypted.csv + CustomerCustomer11_activities_encrypted.csv
                         + CustomerCustomer12_activities_encrypted.csv + CustomerCustomer2_activities_encrypted.csv
                         + CustomerCustomer5_activities_encrypted.csv + CustomerCustomer6_activities_encrypted.csv, data = df_perCustomer)
summary(model_activities)

# use 'tidy' from broom package for a cleaner output
model_activities <- tidy(model_activities)
model_activities


# Model for Total Activities (with average values)

model_activities_avg <- lm(Average_Activity ~ Month.Number, data = df_avg_per_month)
summary(model_activities_avg)

# use 'tidy' from broom package for a cleaner output
model_activities_avg_tidy <- tidy(model_activities_avg)
model_activities_avg_tidy

# Calculate residuals and fitted values
residuals <- resid(model_activities_avg)
fitted_values <- fitted(model_activities_avg)

# Create a dataframe for plotting
residuals_df <- data.frame(Residuals = residuals, Fitted = fitted_values, Month = df_avg_per_month$Month.Number)

# Residuals vs Month Number Plot
ggplot(residuals_df, aes(x = Month, y = Residuals)) +
  geom_point() +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
  labs(x = "Month Number", y = "Residuals", title = "Residuals vs Month Number")


library(alr4)

# Fit a simple linear regression model
model_linear <- lm(Average_Activity ~ Month.Number, data = df_avg_per_month)

# Fit a more complex model, such as a polynomial model
model_poly <- lm(Average_Activity ~ poly(Month.Number, 2, raw=TRUE), data = df_avg_per_month)

# Use anova to compare the two models
anova_linear_poly <- anova(model_linear, model_poly)
anova_linear_poly

library(gvlma)

# Perform the global validation test
gvmodel <- gvlma(model_linear)
summary(gvmodel)

# In-Person Activities

# Inspect Plot

library(ggplot2)

ggplot(df_avg_per_month, aes(x = Month.Number, y = Average_InPerson)) +
  geom_point() +
  labs(x = "Month Number", y = "Average In-Person Activity", title = "Plot of Average In-Person Activities per Month")

#Plot proportion

ggplot(df_avg_per_month, aes(x = Month.Number, y = Proportion_InPerson)) +
  geom_point() +
  labs(x = "Month Number", y = "% In-Person Activity", title = "Plot of % In-Person Activities per Month")


# Piece-wise Regression for Activities

# Assuming breakpoints at Month 10 and Month 20, as an example
break1 <- 6
break2 <- 23

df_avg_per_month$Segment <- with(df_avg_per_month, ifelse(Month.Number <= break1, "Segment1",
                                                          ifelse(Month.Number <= break2, "Segment2", "Segment3")))

# Fit separate linear models for each segment
model_seg1 <- lm(Average_Activity ~ Month.Number, data = df_avg_per_month, subset = Segment == "Segment1")
model_seg2 <- lm(Average_Activity ~ Month.Number, data = df_avg_per_month, subset = Segment == "Segment2")
model_seg3 <- lm(Average_Activity ~ Month.Number, data = df_avg_per_month, subset = Segment == "Segment3")

summary(model_seg1)
summary(model_seg2)
summary(model_seg3)


# Define a function to apply the correct model based on the segment
predict_by_segment_with_CI <- function(segment, month_number) {
  if (segment == "Segment1") {
    prediction <- predict(model_seg1, newdata = data.frame(Month.Number = month_number), interval = "confidence")
  } else if (segment == "Segment2") {
    prediction <- predict(model_seg2, newdata = data.frame(Month.Number = month_number), interval = "confidence")
  } else {  # For Segment 3
    prediction <- predict(model_seg3, newdata = data.frame(Month.Number = month_number), interval = "confidence")
  }
  return(prediction)
}

# Apply the function to the entire dataframe to get predicted values and confidence intervals
predictions_and_CIs <- mapply(predict_by_segment_with_CI, df_avg_per_month$Segment, df_avg_per_month$Month.Number, SIMPLIFY = FALSE)

# Extract the predicted values and confidence intervals
df_avg_per_month$Predicted_Activity <- sapply(predictions_and_CIs, "[", 1)
df_avg_per_month$Predicted_Activity_Lower <- sapply(predictions_and_CIs, "[", 2)
df_avg_per_month$Predicted_Activity_Upper <- sapply(predictions_and_CIs, "[", 3)


# Calculate the growth rate percentage (Month over Month)
# Add NA at the beginning of the diff() vector to align the lengths
growth_rates <- c(NA, diff(df_avg_per_month$Predicted_Activity))

# The growth rate is the change in predicted activity divided by the previous month's predicted activity
df_avg_per_month$Growth_Rate_Percentage <- growth_rates / lag(df_avg_per_month$Predicted_Activity, default = first(df_avg_per_month$Predicted_Activity)) * 100

# Replace any potential NaN (resulting from division by zero) with zero
df_avg_per_month$Growth_Rate_Percentage[is.nan(df_avg_per_month$Growth_Rate_Percentage)] <- 0

df_avg_per_month <- df_avg_per_month %>%
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
cmgr_1_to_6 <- calculate_cmgr(df_avg_per_month, 1, 6)
cmgr_1_to_6

# calculate CMGR from Month 1 to Month 6
cmgr_24_to_36 <- calculate_cmgr(df_avg_per_month, 25, 35)
cmgr_24_to_36


# We offset the predicted values by 1 to align with the 'diff' output
df_avg_per_month <- df_avg_per_month %>%
  mutate(
    Growth_Rate_Percentage = c(NA, diff(Predicted_Activity) / head(Predicted_Activity, -1) * 100),
    Growth_Rate_Percentage_Lower = c(NA, diff(Predicted_Activity_Lower) / head(Predicted_Activity_Lower, -1) * 100),
    Growth_Rate_Percentage_Upper = c(NA, diff(Predicted_Activity_Upper) / head(Predicted_Activity_Upper, -1) * 100)
  )

# Replace any potential NaN or infinite values with zero
df_avg_per_month$Growth_Rate_Percentage[is.nan(df_avg_per_month$Growth_Rate_Percentage) | is.infinite(df_avg_per_month$Growth_Rate_Percentage)] <- 0
df_avg_per_month$Growth_Rate_Percentage_Lower[is.nan(df_avg_per_month$Growth_Rate_Percentage_Lower) | is.infinite(df_avg_per_month$Growth_Rate_Percentage_Lower)] <- 0
df_avg_per_month$Growth_Rate_Percentage_Upper[is.nan(df_avg_per_month$Growth_Rate_Percentage_Upper) | is.infinite(df_avg_per_month$Growth_Rate_Percentage_Upper)] <- 0


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

# Calculate CMGR with CI from Month 1 to Month 6
cmgr_1_to_6 <- calculate_cmgr_with_ci(df_avg_per_month, 2, 6)
cmgr_1_to_6

# Calculate CMGR with CI from Month 24 to Month 36
cmgr_24_to_36 <- calculate_cmgr_with_ci(df_avg_per_month, 26, 35)
cmgr_24_to_36


# Monte Carlo Simulation for Converstion

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
mean_percent <- 1    # 1%
sd_percent <- 0.2    # 0.2%
num_simulations <- 1000  # Number of simulations

# Apply the Monte Carlo simulation
df_avg_per_month <- df_avg_per_month %>%
  rowwise() %>%
  mutate(
    Closed_Deals_Lower = monte_carlo_simulation(Predicted_Activity, mean_percent, sd_percent, num_simulations)["lower.2.5%"],
    Closed_Deals_Median = monte_carlo_simulation(Predicted_Activity, mean_percent, sd_percent, num_simulations)["median"],
    Closed_Deals_Upper = monte_carlo_simulation(Predicted_Activity, mean_percent, sd_percent, num_simulations)["upper.97.5%"]
  ) %>%
  ungroup()

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
sale_mean <- 1000  # Example mean sale value
sale_sd <- 200     # Example standard deviation of sale value

# Apply the Monte Carlo simulation for revenue
df_avg_per_month <- df_avg_per_month %>%
  rowwise() %>%
  mutate(
    Revenue_Lower = monte_carlo_simulation_revenue(Closed_Deals_Median, sale_mean, sale_sd, num_simulations)["lower.2.5%"],
    Revenue_Median = monte_carlo_simulation_revenue(Closed_Deals_Median, sale_mean, sale_sd, num_simulations)["median"],
    Revenue_Upper = monte_carlo_simulation_revenue(Closed_Deals_Median, sale_mean, sale_sd, num_simulations)["upper.97.5%"]
  ) %>%
  ungroup()


# Calculate cumulative sales
df_avg_per_month <- df_avg_per_month %>%
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
df_avg_per_month$ROI_Year1_Lower <- calculate_roi(df_avg_per_month$Cumulative_Sales_Lower, annual_fee, 12)
df_avg_per_month$ROI_Year2_Lower <- calculate_roi(df_avg_per_month$Cumulative_Sales_Lower, annual_fee, 24)
df_avg_per_month$ROI_Year3_Lower <- calculate_roi(df_avg_per_month$Cumulative_Sales_Lower, annual_fee, 36)

df_avg_per_month$ROI_Year1_Median <- calculate_roi(df_avg_per_month$Cumulative_Sales_Median, annual_fee, 12)
df_avg_per_month$ROI_Year2_Median <- calculate_roi(df_avg_per_month$Cumulative_Sales_Median, annual_fee, 24)
df_avg_per_month$ROI_Year3_Median <- calculate_roi(df_avg_per_month$Cumulative_Sales_Median, annual_fee, 36)

df_avg_per_month$ROI_Year1_Upper <- calculate_roi(df_avg_per_month$Cumulative_Sales_Upper, annual_fee, 12)
df_avg_per_month$ROI_Year2_Upper <- calculate_roi(df_avg_per_month$Cumulative_Sales_Upper, annual_fee, 24)
df_avg_per_month$ROI_Year3_Upper <- calculate_roi(df_avg_per_month$Cumulative_Sales_Upper, annual_fee, 36)


# Revenue Graph

library(ggplot2)

# Create the line graph
ggplot(df_avg_per_month, aes(x = Month.Number)) +
  geom_line(aes(y = Revenue_Median), color = "blue") +
  geom_ribbon(aes(ymin = Revenue_Lower, ymax = Revenue_Upper), fill = "blue", alpha = 0.3) +
  labs(title = "Monthly Revenue Estimates",
       x = "Month",
       y = "Revenue",
       caption = "Shaded area represents the range between lower and upper revenue estimates.") +
  theme_minimal()

#CLosed Deals Graph

library(ggplot2)

# Create the bar graph
ggplot(df_avg_per_month, aes(x = Month.Number, y = Closed_Deals_Median)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  labs(title = "Estimated Closed Deals by Month",
       x = "Month",
       y = "Estimated Closed Deals",
       caption = "Data represents the median estimated closed deals.") +
  theme_minimal()

