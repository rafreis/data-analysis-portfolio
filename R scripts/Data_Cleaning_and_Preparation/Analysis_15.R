setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/evers_e")

library(readxl)
df_annual <- read_xlsx('Data total_clean.xlsx', sheet='Total_annualised')
df_quarterly <- read_xlsx('Data total_clean.xlsx', sheet='Rental price_quarterly')

# Get rid of spaces in column names
  
names(df_annual) <- gsub(" ", "_", names(df_annual))
names(df_quarterly) <- gsub(" ", "_", names(df_quarterly))

names(df_annual) <- gsub("\\(", "", names(df_annual))  # Remove opening brackets
names(df_annual) <- gsub("\\)", "", names(df_annual)) 
names(df_quarterly) <- gsub("\\(", "", names(df_quarterly))
names(df_quarterly) <- gsub("\\)", "", names(df_quarterly))

names(df_annual) <- gsub("\\²", "2", names(df_annual)) 
names(df_quarterly) <- gsub("\\²", "2", names(df_quarterly))

# Eliminate last row of data (2023)

df_annual <- df_annual[-nrow(df_annual), ]

# Load necessary libraries
library(ggplot2)
library(tidyr)
library(dplyr)

annual_vars = c("Rental_price_CBS", "Rental_price_Korevaar", 
                "Market_Rental_Value_Growth_MSCI", "Gross_Rent_Passing_Growth_MSCI", "Rental_price_increase_contracts_liberalised_IVBN_huurenquete",
                "Rental_price_increase_excluding_turnover,_liberalised_CBS","Rental_price_increase_VGM_NL")

# Ensure all columns are numeric
df_annual <- df_annual %>% 
  mutate(across(all_of(annual_vars), ~ as.numeric(as.character(.))))

# Ensure all columns are numeric
df_annual <- df_annual %>% 
  mutate(across(all_of(annual_price_vars), ~ as.numeric(as.character(.))))

# Reshape data to long format
df_annual_long <- df_annual %>% 
  pivot_longer(
    cols = annual_vars,
    names_to = "Variable",
    values_to = "Value"
  )

# Rename the 'Rijlabels' column to 'Quarters'
names(df_quarterly)[names(df_quarterly) == "Rijlabels"] <- "Quarters"

# Unite the 'Year' and 'Quarters' columns
df_quarterly <- df_quarterly %>%
  unite("YearQuarter", c("Year", "Quarters"), sep = ".")

quarter_vars = c("Market_Rental_Value_per_m2,_per_quarter,_value_per_month_MSCI", "Transactionmonitor_price_sq._M.IVBN", "Rental_price_per_m2_VGM_NL", "Unfurnished_rental_price_per_sq_m_Pararius")
  
# Reshape the data to long format
df_quarterly_long <- df_quarterly %>% 
  pivot_longer(
    cols = quarter_vars,
    names_to = "Variable",
    values_to = "Value"
  )

## COMPOSITE RENTAL PRICE INDEX AND GROWTH
library(FactoMineR)

# Function to perform PCA and return factor loadings along with PCA details
perform_pca <- function(data) {
  pca_result <- PCA(data, graph = FALSE)
  # Print PCA results
  print(pca_result)
  # Return factor loadings of the first principal component
  return(pca_result$var$coord[, 1])
}

# Initialize a list to store composite indices
composite_indices <- list()

# Function to calculate weighted average considering only positive loadings
weighted_average <- function(data, loadings) {
  # Set negative loadings to zero
  positive_loadings <- ifelse(loadings > 0, loadings, 0)
  
  # Calculate weighted average using positive loadings
  weighted_data <- sweep(data, 2, positive_loadings, `*`)
  rowSums(weighted_data, na.rm = TRUE) / sum(positive_loadings, na.rm = TRUE)
}

# Function to calculate simple average
simple_average <- function(data) {
  rowMeans(data, na.rm = TRUE)
}

# Subset 1: 1961-1994
subset1 <- df_annual %>% 
  filter(Year >= 1961 & Year <= 1994) %>%
  select(Year, Rental_price_CBS, Rental_price_Korevaar) %>%
  drop_na()
loadings1 <- perform_pca(select(subset1, -Year))
composite_indices[["1961-1994"]] <- data.frame(Year = subset1$Year, Composite_Index = weighted_average(select(subset1, -Year), loadings1))

# Subset 2: 1995-2008
subset2 <- df_annual %>% 
  filter(Year >= 1995 & Year <= 2008) %>% 
  select(Year, annual_vars[1:4]) %>%
  drop_na()
loadings2 <- perform_pca(select(subset2, -Year))
composite_indices[["1995-2008"]] <- data.frame(Year = subset2$Year, Composite_Index = weighted_average(select(subset2, -Year), loadings2))

# Subset 3: 2009-2014
subset3 <- df_annual %>% 
  filter(Year >= 2009 & Year <= 2014) %>% 
  select(Year, annual_vars[1:5]) %>%
  drop_na()
loadings3 <- perform_pca(select(subset3, -Year))
composite_indices[["2009-2014"]] <- data.frame(Year = subset3$Year, Composite_Index = weighted_average(select(subset3, -Year), loadings3))

# Subset 4: 2015-2022
subset4 <- df_annual %>% 
  filter(Year >= 2015 & Year <= 2022) %>% 
  select(Year, annual_vars[-2]) %>%
  drop_na()
loadings4 <- perform_pca(select(subset4, -Year))
composite_indices[["2015-2022"]] <- data.frame(Year = subset4$Year, Composite_Index = weighted_average(select(subset4, -Year), loadings4))

# Combine all the composite indices
final_composite_index <- do.call(rbind, composite_indices)

# Resulting data frame has Year and Composite_Index for each subset
print(final_composite_index)

## ADJUSTED FOR INFLATION

# Add the Inflation column to the final_composite_index for corresponding years
final_composite_index <- final_composite_index %>%
  left_join(df_annual %>% select(Year, "Inflation,_lagged_a_year_CPI,_CBS"), by = "Year")

# Calculate the real composite index by subtracting inflation
final_composite_index$real_composite_index <- final_composite_index$Composite_Index - final_composite_index$"Inflation,_lagged_a_year_CPI,_CBS"

# View the result
print(final_composite_index)

## ANNUAL RENT PRICES (NOT GROWTH)

annual_price_vars = c("Market_Rental_Value_per_m2_MSCI", "Unfurnished_annual_Pararius", "Transactionmonitor_price_sq._M.IVBN", "Rental_price_per_m2_VGM_NL")

# Initialize a list to store composite indices
composite_indices_price <- list()

# Subset 1: 1995-2008
subset2 <- df_annual %>% 
  filter(Year >= 1995 & Year <= 2007) %>% 
  select(Year, "Market_Rental_Value_per_m2_MSCI") %>%
  drop_na()
loadings2 <- perform_pca(select(subset2, -Year))
composite_indices_price[["1995-2007"]] <- data.frame(Year = subset2$Year, Composite_Index = weighted_average(select(subset2, -Year), loadings2))

# Subset 2: 2009-2022
subset3 <- df_annual %>% 
  filter(Year >= 2008 & Year <= 2015) %>% 
  select(Year, "Market_Rental_Value_per_m2_MSCI", "Transactionmonitor_price_sq._M.IVBN") %>%
  drop_na()
loadings3 <- perform_pca(select(subset3, -Year))
composite_indices_price[["2008-2015"]] <- data.frame(Year = subset3$Year, Composite_Index = weighted_average(select(subset3, -Year), loadings3))

# Subset 3: 2009-2022
subset4 <- df_annual %>% 
  filter(Year >= 2016 & Year <= 2022) %>% 
  select(Year, all_of(annual_price_vars)) %>%
  drop_na()
loadings4 <- perform_pca(select(subset4, -Year))
composite_indices_price[["2016-2022"]] <- data.frame(Year = subset4$Year, Composite_Index = weighted_average(select(subset4, -Year), loadings4))

# Combine the composite indices for the subsets
final_composite_index_rentprice <- do.call(rbind, composite_indices_price)

## QUARTERLY RENTAL PRICE INDEX

# Remove the last two rows from df_quarterly
df_quarterly <- head(df_quarterly, -2)

# Initialize a list to store composite indices
composite_indices <- list()

# Subset 1: Until YearQuarter "2013.Q4"
index_2013Q4 <- which(df_quarterly$YearQuarter == "2013.Q4")
subset1 <- df_quarterly[1:index_2013Q4, ]
subset1_cols <- c("Market_Rental_Value_per_m2,_per_quarter,_value_per_month_MSCI", "Transactionmonitor_price_sq._M.IVBN")
loadings1 <- perform_pca(subset1[, subset1_cols])
composite_indices[["Until 2013.Q4"]] <- data.frame(YearQuarter = subset1$YearQuarter, Composite_Index = weighted_average(subset1[, subset1_cols], loadings1))

# Subset 2: From "2014.Q1" to "2015.Q4"
index_2014Q1 <- which(df_quarterly$YearQuarter == "2014.Q1")
index_2015Q4 <- which(df_quarterly$YearQuarter == "2015.Q4")
subset2 <- df_quarterly[index_2014Q1:index_2015Q4, ]
subset2_cols <- c("Market_Rental_Value_per_m2,_per_quarter,_value_per_month_MSCI", "Transactionmonitor_price_sq._M.IVBN", "Rental_price_per_m2_VGM_NL")
loadings2 <- perform_pca(subset2[, subset2_cols])
composite_indices[["2014.Q1-2015.Q4"]] <- data.frame(YearQuarter = subset2$YearQuarter, Composite_Index = weighted_average(subset2[, subset2_cols], loadings2))

# Subset 3: From "2016.Q1"
index_2016Q1 <- which(df_quarterly$YearQuarter == "2016.Q1")
subset3 <- df_quarterly[index_2016Q1:nrow(df_quarterly), ]
subset3_cols <- c("Market_Rental_Value_per_m2,_per_quarter,_value_per_month_MSCI", "Transactionmonitor_price_sq._M.IVBN", "Rental_price_per_m2_VGM_NL", "Unfurnished_rental_price_per_sq_m_Pararius")
loadings3 <- perform_pca(subset3[, subset3_cols])
composite_indices[["From 2016.Q1"]] <- data.frame(YearQuarter = subset3$YearQuarter, Composite_Index = weighted_average(subset3[, subset3_cols], loadings3))

# Combine all the composite indices
final_composite_index_quarter <- do.call(rbind, composite_indices)

# Resulting data frame has YearQuarter and Composite_Index for each subset
print(final_composite_index_quarter)


## QUARTERLY CHANGE
# Initialize a list to store composite indices
composite_indices <- list()

# Subset 1: Until YearQuarter "2014.Q1"
index_2014Q1 <- which(df_quarterly$YearQuarter == "2014.Q1")
subset1 <- df_quarterly[2:index_2014Q1, ]
subset1_cols <- c("Market_Rental_Value_Growth_quarterly_MSCI", "Gross_Rent_Passing_Growth_MSCI", "Transactionmonitor_growth_QoQ_IVBN")
loadings1 <- perform_pca(subset1[, subset1_cols])
composite_indices[["Until 2014.Q1"]] <- data.frame(YearQuarter = subset1$YearQuarter, Composite_Index = weighted_average(subset1[, subset1_cols], loadings1))

# Subset 2: From "2014.Q2" to "2016.Q1"
index_2014Q2 <- which(df_quarterly$YearQuarter == "2014.Q2")
index_2016Q1 <- which(df_quarterly$YearQuarter == "2016.Q1")
subset2 <- df_quarterly[index_2014Q2:index_2016Q1, ]
subset2_cols <- c("Market_Rental_Value_Growth_quarterly_MSCI", "Gross_Rent_Passing_Growth_MSCI", "Transactionmonitor_growth_QoQ_IVBN", "Rental_price_increase_VGM_NL")
loadings2 <- perform_pca(subset2[, subset2_cols])
composite_indices[["2014.Q2-2016.Q1"]] <- data.frame(YearQuarter = subset2$YearQuarter, Composite_Index = weighted_average(subset2[, subset2_cols], loadings2))

# Subset 3: From "2016.Q2"
index_2016Q2 <- which(df_quarterly$YearQuarter == "2016.Q2")
subset3 <- df_quarterly[index_2016Q2:nrow(df_quarterly), ]
subset3_cols <- c("Market_Rental_Value_Growth_quarterly_MSCI", "Gross_Rent_Passing_Growth_MSCI", "Transactionmonitor_growth_QoQ_IVBN", "Rental_price_increase_VGM_NL", "Unfurnished_rental_price_growth_average_Pararius")
loadings3 <- perform_pca(subset3[, subset3_cols])
composite_indices[["From 2016.Q2"]] <- data.frame(YearQuarter = subset3$YearQuarter, Composite_Index = weighted_average(subset3[, subset3_cols], loadings3))

# Combine all the composite indices
final_composite_index_quarter_growth <- do.call(rbind, composite_indices)

# Resulting data frame has YearQuarter and Composite_Index for each subset
print(final_composite_index_quarter_growth)

## ANNUAL PLOTS

library(ggplot2)

# Define the variable categories
contract_prices_vars <- c("Gross_Rent_Passing_Growth_MSCI", 
                          "Rental_price_increase_excluding_turnover,_liberalised_CBS",
                          "Rental_price_increase_contracts_liberalised_IVBN_huurenquete")

market_prices_vars <- setdiff(annual_vars, contract_prices_vars)

# Reshape data to long format
df_annual_long <- df_annual %>%
  pivot_longer(
    cols = annual_vars,
    names_to = "Variable",
    values_to = "Value"
  )

# Split the data into contract and market prices
df_contract_prices <- df_annual_long %>% filter(Variable %in% contract_prices_vars)
df_market_prices <- df_annual_long %>% filter(Variable %in% market_prices_vars)

# Plot for Contract Rent Prices
plot_contract_prices <- ggplot(df_contract_prices, aes(x = Year, y = Value, color = Variable)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Historical Annual Trends of Contract Rent Price Changes (%)",
       x = "Year",
       y = "Value",
       color = "Variable") +
  scale_color_brewer(palette = "Set1") +
  theme(legend.text = element_text(size = 7),
        legend.title = element_text(size = 10)) +
  theme(plot.margin = margin(t = 50, r = 0, b = 50, l = 0)) +
  scale_y_continuous(limits = c(0, 10))  # Set y-axis limits from 0 to 10

# Save the plot
ggsave("contract_rent_prices_plot.png", plot_contract_prices, width = 10, height = 6)


# Plot for Market Rent Prices Combined
plot_market_prices <- ggplot(df_market_prices, aes(x = Year, y = Value, color = Variable)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Historical Annual Trends of Market Rent Price Changes (%)",
       x = "Year",
       y = "Value",
       color = "Variable") +
  scale_color_brewer(palette = "Set1") +
  theme(legend.text = element_text(size = 7),
        legend.title = element_text(size = 10)) +
  theme(plot.margin = margin(t = 50, r = 0, b = 50, l = 0))

# Save the plot
ggsave("market_rent_prices_plot.png", plot_market_prices)


# Set a larger plot size
options(repr.plot.width = 10, repr.plot.height = 6)

# Plot Historical Trends of Rental Price Changes
plot9 <- ggplot(df_annual_long, aes(x = Year, y = Value, color = Variable)) +
  geom_line() +
  theme_minimal() +
  labs(title = "Historical Annual Trends of Rental Price Changes (%)",
       x = "Year",
       y = "Value",
       color = "Variable") +
  scale_color_brewer(palette = "Set1") +  # Adjust color palette as needed
  theme(legend.text = element_text(size = 7),  # Adjust legend text size
        legend.title = element_text(size = 10)) +  # Adjust legend title size
  theme(plot.margin = margin(t = 50, r = 0, b = 50, l = 0))  # Adjust margins

plot9
ggsave("anual_rental_changes_plot.png", plot9)


# Plot for final_composite_index
plot2 <- ggplot(final_composite_index, aes(x = Year, y = Composite_Index)) + 
  geom_line() +
  theme_minimal() +
  labs(title = "Annual Composite Rental Price Growth Index", x = "Year", y = "Index")
plot2
ggsave("final_annual_composite_index_growth_plot.png", plot2)

# Plot for final_composite_index and real_composite_index
plot2 <- ggplot(final_composite_index, aes(x = Year)) +
  geom_line(aes(y = Composite_Index, color = "Composite Index")) +
  geom_line(aes(y = real_composite_index, color = "Real Composite Index")) +
  theme_minimal() +
  labs(title = "Annual Composite Rental Price Index", 
       x = "Year", 
       y = "Index",
       color = "Index Type") +
  scale_color_manual(values = c("Composite Index" = "blue", "Real Composite Index" = "red"))

# Save the plot
ggsave("annual_growth_composite_index_plot.png", plot2, width = 10, height = 6)

#Plot for composite index of rent price
plot99 <- ggplot(final_composite_index_rentprice, aes(x = Year, y = Composite_Index)) + 
  geom_line() +
  theme_minimal() +
  labs(title = "Annual Composite Rental Price Index", x = "Year", y = "Index")

ggsave("annual_price_composite_index_plot.png", plot99)


##QUARTERLY PLOTS

# Reshape df_quarterly into long format for the specified variables
df_quarterly_long_2 <- df_quarterly %>%
  pivot_longer(cols = c("Market_Rental_Value_Growth_quarterly_MSCI", 
                        "Gross_Rent_Passing_Growth_MSCI", 
                        "Transactionmonitor_growth_QoQ_IVBN", 
                        "Rental_price_increase_VGM_NL", 
                        "Unfurnished_rental_price_growth_average_Pararius"),
               names_to = "Variable", values_to = "Value")

# Set a larger plot size
options(repr.plot.width = 10, repr.plot.height = 6)

# PLot Historical Trends of Rental Price (quarterly)
plot10 <- ggplot(df_quarterly_long, aes(x = YearQuarter, y = Value, color = Variable, group = Variable)) +
  geom_line() +
  scale_x_discrete("Quarters") +
  scale_y_continuous("Value") +
  theme_minimal() +
  labs(title = "Historical Quarterly Trends of Rental Price",
       color = "Variable") +
  scale_color_brewer(palette = "Set1") +
  theme(axis.text.x = element_text(angle = 45, vjust = 0.5, hjust = 0.5),
        legend.position = "bottom") +  # Place the legend below the graph
  theme(plot.margin = margin(t = 20, r = 10, b = 70, l = 10))  # Adjust margins

plot10
ggsave("quarterly_rental_prices_plot.png", plot10, width = 10, height = 6)  # Save the larger plot

# Create the line plot
plot11 <- ggplot(df_quarterly_long_2, aes(x = YearQuarter, y = Value, color = Variable, group = Variable)) +
  geom_line() +
  scale_x_discrete("Quarters") +
  scale_y_continuous("Value") +
  theme_minimal() +
  labs(title = "Historical Quarterly Trends of Rental Price Changes (%)",
       color = "Variable") +
  scale_color_brewer(palette = "Set1") +
  theme(axis.text.x = element_text(angle = 45, vjust = 0.5, hjust = 0.5),
        legend.position = "bottom") +  # Place the legend below the graph
  theme(plot.margin = margin(t = 20, r = 10, b = 50, l = 10))  # Adjust margins

plot11
ggsave("quarterly_rental_price_changes_plot.png", plot11, width = 10, height = 6)  # Save the larger plot


# Convert YearQuarter to a factor with levels in the correct order
final_composite_index_quarter$YearQuarter <- factor(final_composite_index_quarter$YearQuarter, levels = unique(final_composite_index_quarter$YearQuarter))
final_composite_index_quarter_growth$YearQuarter <- factor(final_composite_index_quarter_growth$YearQuarter, levels = unique(final_composite_index_quarter_growth$YearQuarter))

# Plot for final_composite_index_quarter
plot3 <- ggplot(final_composite_index_quarter, aes(x = YearQuarter, y = Composite_Index, group = 1)) + 
  geom_line() +
  theme_minimal() +
  labs(title = "Quarterly Composite Rental Price Index", x = "Year Quarter", y = "Index") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1))
plot3
ggsave("final_quarterly_composite_index_plot.png", plot3)

# Plot for final_composite_index_quarter_growth
plot4 <- ggplot(final_composite_index_quarter_growth, aes(x = YearQuarter, y = Composite_Index, group = 1)) + 
  geom_line() +
  theme_minimal() +
  labs(title = "Quarterly Composite Rental Price Growth Index", x = "Year Quarter", y = "Growth") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1))
plot4
ggsave("final_quarterly_growth_composite_index_plot.png", plot4)

##EXPORT TABLES

library(writexl)

# Create a list of data frames
list_of_dfs <- list(
  Composite_Index_growth = final_composite_index,
  Composite_Index = final_composite_index_rentprice,
  Composite_Index_Quarter = final_composite_index_quarter,
  Composite_Index_Quarter_Growth = final_composite_index_quarter_growth
)

# Write to Excel
write_xlsx(list_of_dfs, "Composite_Indexes.xlsx")
