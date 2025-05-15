library(dplyr)
library(openxlsx)
library(lubridate)
library(stringr)
library(ggplot2)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/kevinh75")

df <- read.xlsx("ClusterData_9_6_23.xlsx")

## DATA CLEANING ##

# Define the function to convert Excel date to R date
convert_excel_date <- function(excel_date) {
  as.Date(excel_date, origin="1899-12-30")
}

# Apply the function to the Job.Date and Sold.Date columns
df <- df %>%
  mutate(Job.Date = convert_excel_date(Job.Date),
         Sold.Date = convert_excel_date(Sold.Date))

# Calculate the difference in months and add as a new column
df <- df %>%
  mutate(Months.Since.Purchase = as.numeric(round(interval(Sold.Date, Job.Date) / months(1), 0)))

# Removing duplicates by address

# Group by address and sort by Months.Since.Purchase
df <- df %>%
  arrange(Zillow.Full.Address, Months.Since.Purchase)

# Remove duplicates by keeping the first entry for each address group
df_unique <- df %>%
  group_by(Zillow.Full.Address) %>%
  slice_head(n = 1) %>%
  ungroup()

# Add a column for the number of occurrences for each address
count_occurrences <- df %>%
  group_by(Zillow.Full.Address) %>%
  summarise(Num_Occurrences = n(), .groups = 'drop')

# Merge the occurrence counts into df_unique
df_unique <- left_join(df_unique, count_occurrences, by = "Zillow.Full.Address")

# Print the number of cases with 1, 2, 3, and more occurrences
occurrence_count <- table(count_occurrences$Num_Occurrences)
print(occurrence_count)

#Recode Num_Occurrences
df_unique$Num_Occurrences <- ifelse(df_unique$Num_Occurrences == 1, "No", "Yes")

# Rename the column
colnames(df_unique)[colnames(df_unique) == "Num_Occurrences"] <- "Returning.Customer"

## Extracting the State

# Extract the state as a new column
df_unique <- df_unique %>%
  mutate(State = str_extract(Zillow.Full.Address, "(?<=, )([A-Z]{2})(?=,)"))

# View the modified dataframe to confirm
head(df_unique)

## DIAGNOSE MISSING VALUES

# Count missing values in specific columns
missing_values_summary <- df_unique %>%
  summarise(
    missing_bedrooms = sum(is.na(Bedrooms)),
    missing_bathrooms = sum(is.na(Bathrooms)),
    missing_price_per_sqft = sum(is.na(Price.Per.sq.ft)),
    missing_has_hoa = sum(is.na(Has.HOA)),
    missing_property_age = sum(is.na(Property.Age)),
    missing_total_structure_area_sqft = sum(is.na(`Total.Structure.Area.(sqft)`)),
    missing_has_garage = sum(is.na(Has.Garage)),
    missing_months_since_purchase = sum(is.na(Months.Since.Purchase)),
    missing_num_occurrences = sum(is.na(Returning.Customer)),
    missing_state = sum(is.na(State))
  )

# Print the summary to the console
print(missing_values_summary)

# Rename columns
df_unique <- df_unique %>% rename(Total.Structure.Area.sqft = `Total.Structure.Area.(sqft)`)


# Calculate the frequency table for the Discovery_Confidence_Level column
freq_table <- table(df_unique$Discovery.Confidence.Level)


# Print the frequency table
print(freq_table) ## LOW NUMBER OF LOW AND MEDIUM CONFIDENCE

# Subset the dataframe to only include rows where Discovery_Confidence_Level is 'High'
df_unique_high <- subset(df_unique, Discovery.Confidence.Level == "High")

# Subset data with relevant variables
df_unique_selected <- df_unique_high[, c('Bedrooms', 'Bathrooms', 'Price.Per.sq.ft', 'Has.HOA', 'Property.Age', 
                                    'Total.Structure.Area.sqft', 'Has.Garage', 'Months.Since.Purchase', 
                                    'Returning.Customer', 'State')]
# Get Complete Cases Only

# Find the complete cases
complete_cases_index <- complete.cases(df_unique_selected)

# Count the number of complete cases
num_complete_cases <- sum(complete_cases_index)

# Create a new data frame with only complete cases
df_complete_cases <- df_unique_selected[complete_cases_index, ]

# Create a new whole data frame with only complete cases
df_unique_high <- df_unique_high[complete_cases_index, ]

## CLUSTERING ALGORITHM

library(cluster)
library(fpc)

# Convert to factor for categorical variables
df_complete_cases$Has.HOA <- as.factor(df_complete_cases$Has.HOA)
df_complete_cases$Has.Garage <- as.factor(df_complete_cases$Has.Garage)
df_complete_cases$State <- as.factor(df_complete_cases$State)
df_complete_cases$Returning.Customer <- as.factor(df_complete_cases$Returning.Customer)

# Convert to numeric for numerical variables
df_complete_cases$Bedrooms <- as.numeric(df_complete_cases$Bedrooms)
df_complete_cases$Bathrooms <- as.numeric(df_complete_cases$Bathrooms)
df_complete_cases$Price.Per.sq.ft <- as.numeric(df_complete_cases$Price.Per.sq.ft)
df_complete_cases$Property.Age <- as.numeric(df_complete_cases$Property.Age)
df_complete_cases$Total.Structure.Area.sqft <- as.numeric(df_complete_cases$Total.Structure.Area.sqft)
df_complete_cases$Months.Since.Purchase <- as.numeric(df_complete_cases$Months.Since.Purchase)

## DEFINING OPTIMAL NUMBER OF CLUSTERS

# Calculate Gower distance matrix for df_complete_cases
gower_dist <- daisy(df_complete_cases, metric = "gower")

# Initialize the vector to store silhouette widths
silhouette_width <- numeric(7)

# Loop to calculate silhouette widths for k = 2 to k = 7
for (k in 2:7) {
  pam_fit <- pam(gower_dist, diss = TRUE, k = k)
  s_width <- silhouette(pam_fit)
  silhouette_width[k] <- mean(s_width[, 3])
}

# Plotting the silhouette widths
plot(2:7, silhouette_width[2:7], type = 'b',
     xlab = "Number of clusters", ylab = "Average silhouette width",
     main = "Average silhouette for varying number of clusters")

#Run Solution with 4 clusters
set.seed(123)
pam_fit <- pam(gower_dist, diss = TRUE, k = 4)


# Adding the cluster assignments to your original data frame
df_complete_cases$Cluster <- pam_fit$clustering

## CLUSTER PROFILING

library(dplyr)

# CLUSTER PROFILES

# Calculate cluster sizes and relative percentages
cluster_sizes <- df_complete_cases %>%
  group_by(Cluster) %>%
  summarize(
    Size = n(),
    Relative_Percentage = (n() / nrow(.)) * 100
  )

# Display the dataframe
print(cluster_sizes)


cluster_profile <- df_complete_cases %>% 
  group_by(Cluster) %>% 
  summarize(
    Avg_Price_Per_sqft = mean(Price.Per.sq.ft, na.rm = TRUE),
    Avg_Property_Age = mean(Property.Age, na.rm = TRUE),
    Proportion_with_HOA = mean(Has.HOA == "Yes", na.rm = TRUE),
    Proportion_with_Garage = mean(Has.Garage == "Yes", na.rm = TRUE),
    Avg_Bedrooms = mean(Bedrooms, na.rm = TRUE),
    Avg_Bathrooms = mean(Bathrooms, na.rm = TRUE),
    Avg_Total_Structure_Area = mean(Total.Structure.Area.sqft, na.rm = TRUE),
    Avg_Months_Since_Purchase = mean(Months.Since.Purchase, na.rm = TRUE),
    Proportion_ReturningCustomer = mean(Returning.Customer == "Yes", na.rm = TRUE),
    Most_Common_State = names(sort(table(State), decreasing = TRUE)[1])
  )

# Count the occurrences of each state in each cluster and calculate total per cluster
common_states <- df_complete_cases %>% 
  group_by(Cluster, State) %>% 
  summarise(n = n()) %>% 
  mutate(total_per_cluster = sum(n, na.rm = TRUE))

# Calculate the percentage for each state within its cluster
common_states <- common_states %>% 
  mutate(percentage = (n / total_per_cluster) * 100) %>% 
  arrange(Cluster, desc(percentage))

# Get the top 5 states for each cluster, with percentages
top_states_per_cluster <- common_states %>% 
  group_by(Cluster) %>% 
  slice_head(n = 5) %>% 
  ungroup()

library(ggplot2)
library(tidyverse)

# Create a long-form version of the cluster_profile dataframe for easier plotting
cluster_profile_long <- cluster_profile %>% 
  select(-Most_Common_State) %>% 
  pivot_longer(cols = -Cluster, names_to = "Variable", values_to = "Value")


# Generate the charts
ggplot(cluster_profile_long, aes(x = factor(Cluster), y = Value, fill = factor(Cluster))) +
  geom_bar(stat = "identity", position = "dodge") +
  geom_text(aes(label=round(Value, 2)), vjust=-0.3, size = 3) +
  facet_wrap(~ Variable, scales = "free_y") + # Notice the "free_y" to allow for a dynamic y-axis
  theme_minimal() +
  theme(text=element_text(size=12)) +
  xlab("Cluster") +
  ylab("Value") +
  labs(fill = "Cluster")

#BoxPlots

numerical_vars <- c("Price.Per.sq.ft", "Property.Age", "Bedrooms", "Bathrooms", 
                    "Total.Structure.Area.sqft", "Months.Since.Purchase")

## Remove Extreme outliers
df_filtered_boxplots <- df_complete_cases %>%
  filter(
    between(Price.Per.sq.ft, quantile(Price.Per.sq.ft, 0.01), quantile(Price.Per.sq.ft, 0.99)),
    between(Property.Age, quantile(Property.Age, 0.01), quantile(Property.Age, 0.99)),
    between(Total.Structure.Area.sqft, quantile(Total.Structure.Area.sqft, 0.01), quantile(Total.Structure.Area.sqft, 0.99))
  )

df_complete_cases_long <- df_filtered_boxplots %>% 
  select(Cluster, all_of(numerical_vars)) %>%
  pivot_longer(cols = -Cluster, names_to = "Variable", values_to = "Value")

# Create the boxplots
ggplot(df_complete_cases_long, aes(x = factor(Cluster), y = Value, fill = factor(Cluster))) +
  geom_boxplot() +
  facet_wrap(~ Variable, scales = "free_y") +
  theme_minimal() +
  theme(text=element_text(size=12)) +
  xlab("Cluster") +
  ylab("Value") +
  labs(fill = "Cluster")

# Pie Charts
# Calculate state frequency within each cluster
state_frequency <- df_complete_cases %>%
  group_by(Cluster, State) %>%
  summarise(n = n()) %>%
  mutate(percentage = n / sum(n) * 100)

# Group states with less than 5% representation within each cluster
state_frequency_grouped <- state_frequency %>%
  mutate(State = if_else(percentage < 5, "Other", as.character(State))) %>%
  group_by(Cluster, State) %>%
  summarise(percentage = sum(percentage), .groups = 'drop')

# Generate pie charts
ggplot(state_frequency_grouped, aes(x = "", y = percentage, fill = State)) +
  geom_bar(stat = "identity", width = 1) +
  geom_text(aes(label = sprintf("%.1f%%", percentage)), position = position_stack(vjust = 0.5), size = 3) +
  coord_polar("y") +
  facet_wrap(~ Cluster) +
  scale_fill_brewer(palette = "Set3") +
  labs(title = "State Composition by Cluster", x = NULL, y = NULL, fill = "State") +
  theme_minimal() +
  theme(axis.text.x = element_blank(), legend.position = "right")


# COMPARE NON CLUSTERING VARIBLES

# Initialize a new column in df for cluster labels, setting it to NA initially
df_unique_high$Cluster <- NA

# Assign cluster labels to the original df using the same indices as in df_complete_cases
df_unique_high$Cluster[complete_cases_index] <- df_complete_cases$Cluster

# Remove rows where Cluster is NA
df_unique_high_filtered <- df_unique_high %>% filter(!is.na(Cluster))

df_unique_high_filtered$Aggregate.Sub.Total <- as.numeric(df_unique_high_filtered$Aggregate.Sub.Total)

# Order Service Statistic

df_sales_stats_perCluster <- df_unique_high_filtered %>%
  group_by(Cluster) %>%
  summarise(mean = mean(Aggregate.Sub.Total, na.rm = TRUE),
            median = median(Aggregate.Sub.Total, na.rm = TRUE),
            min = min(Aggregate.Sub.Total, na.rm = TRUE),
            max = max(Aggregate.Sub.Total, na.rm = TRUE),
            .groups = 'drop')

# Create the boxplot
ggplot(df_unique_high_filtered, aes(x = as.factor(Cluster), y = Aggregate.Sub.Total, color = as.factor(Cluster))) +
  geom_boxplot(outlier.shape = NA) +  # exclude outliers for better visualization
  theme_minimal() +
  xlab("Cluster") +
  ylab("Aggregate Sub Total") +
  labs(color = "Cluster")


### A PRIOR SEGMENTATION

# Create columns for regions
df_unique_high <- df_unique_high %>%
  mutate(Region = case_when(
    State %in% c("CT", "ME", "MA", "NH", "RI", "VT") ~ "New England",
    State %in% c("DE", "MD", "NJ", "NY", "PA") ~ "Mid-Atlantic",
    State %in% c("IL", "IN", "MI", "OH", "WI") ~ "East North Central",
    State %in% c("IA", "KS", "MN", "MO", "NE", "ND", "SD") ~ "West North Central",
    State %in% c("DC", "FL", "GA", "MD", "NC", "SC", "VA", "WV") ~ "South Atlantic",
    State %in% c("AL", "KY", "MS", "TN") ~ "East South Central",
    State %in% c("AR", "LA", "OK", "TX") ~ "West South Central",
    State %in% c("AZ", "CO", "ID", "MT", "NV", "NM", "UT", "WY") ~ "Mountain",
    State %in% c("AK", "CA", "HI", "OR", "WA") ~ "Pacific",
    TRUE ~ "Other"
  ))

df_unique_high <- df_unique_high %>%
  mutate(Macro_Region = case_when(
    State %in% c("CT", "ME", "MA", "NH", "RI", "VT", "DE", "MD", "NJ", "NY", "PA") ~ "Northeast",
    State %in% c("IL", "IN", "MI", "OH", "WI", "IA", "KS", "MN", "MO", "NE", "ND", "SD") ~ "Midwest",
    State %in% c("DC", "FL", "GA", "MD", "NC", "SC", "VA", "WV", "AL", "KY", "MS", "TN", "AR", "LA", "OK", "TX") ~ "South",
    State %in% c("AZ", "CO", "ID", "MT", "NV", "NM", "UT", "WY", "AK", "CA", "HI", "OR", "WA") ~ "West",
    TRUE ~ "Other"
  ))

# Create a dataframe

# Convert to factor for categorical variables
df_unique_high$Has.HOA <- as.factor(df_unique_high$Has.HOA)
df_unique_high$Has.Garage <- as.factor(df_unique_high$Has.Garage)
df_unique_high$State <- as.factor(df_unique_high$State)
df_unique_high$Returning.Customer <- as.factor(df_unique_high$Returning.Customer)

# Convert to numeric for numerical variables
df_unique_high$Bedrooms <- as.numeric(df_unique_high$Bedrooms)
df_unique_high$Bathrooms <- as.numeric(df_unique_high$Bathrooms)
df_unique_high$Price.Per.sqft <- as.numeric(df_unique_high$Price.Per.sq.ft)
df_unique_high$Property.Age <- as.numeric(df_unique_high$Property.Age)
df_unique_high$Total.Structure.Area.sqft <- as.numeric(df_unique_high$Total.Structure.Area.sqft)
df_unique_high$Months.Since.Purchase <- as.numeric(df_unique_high$Months.Since.Purchase)
df_unique_high$Aggregate.Sub.Total <- as.numeric(df_unique_high$Aggregate.Sub.Total)

# Grouping by Region, Macro Region and State
df_by_region_all <- df_unique_high %>%
  group_by(Macro_Region, Region, State) %>%
  summarise(
    Avg_Bedrooms = mean(Bedrooms, na.rm = TRUE),
    Avg_Bathrooms = mean(Bathrooms, na.rm = TRUE),
    Avg_Price_Per_sqft = mean(Price.Per.sq.ft, na.rm = TRUE),
    Avg_Property_Age = mean(Property.Age, na.rm = TRUE),
    Avg_Total_Structure_Area = mean(Total.Structure.Area.sqft, na.rm = TRUE),
    Avg_Months_Since_Purchase = mean(Months.Since.Purchase, na.rm = TRUE),
    Avg_Aggregate_Sub_Total = mean(Aggregate.Sub.Total, na.rm = TRUE),
    .groups = 'drop'
  )

# Summarize by State
df_by_state <- df_unique_high %>%
  group_by(State) %>%
  summarise(
    Sample_Size = n(),
    Avg_Bedrooms = mean(Bedrooms, na.rm = TRUE),
    Avg_Bathrooms = mean(Bathrooms, na.rm = TRUE),
    Avg_Price_Per_sqft = mean(Price.Per.sqft, na.rm = TRUE),
    Avg_Property_Age = mean(Property.Age, na.rm = TRUE),
    Avg_Total_Structure_Area = mean(Total.Structure.Area.sqft, na.rm = TRUE),
    Avg_Months_Since_Purchase = mean(Months.Since.Purchase, na.rm = TRUE),
    Avg_Aggregate_Sub_Total = mean(Aggregate.Sub.Total, na.rm = TRUE),
    Prop_Has_HOA = sum(Has.HOA == "Yes") / n(),
    Prop_Has_Garage = sum(Has.Garage == "Yes") / n(),
    Prop_Returning_Customer = sum(Returning.Customer == "Yes") / n(),
    .groups = 'drop'
  )

# Filter cases with sample_size >= 5
df_by_state_filtered <- df_by_state %>% filter(Sample_Size >= 5)

# Summarize by Region
df_by_region <- df_unique_high %>%
  group_by(Region) %>%
  summarise(
    Sample_Size = n(),
    Avg_Bedrooms = mean(Bedrooms, na.rm = TRUE),
    Avg_Bathrooms = mean(Bathrooms, na.rm = TRUE),
    Avg_Price_Per_sqft = mean(Price.Per.sqft, na.rm = TRUE),
    Avg_Property_Age = mean(Property.Age, na.rm = TRUE),
    Avg_Total_Structure_Area = mean(Total.Structure.Area.sqft, na.rm = TRUE),
    Avg_Months_Since_Purchase = mean(Months.Since.Purchase, na.rm = TRUE),
    Avg_Aggregate_Sub_Total = mean(Aggregate.Sub.Total, na.rm = TRUE),
    Prop_Has_HOA = sum(Has.HOA == "Yes", na.rm = TRUE) / n(),
    Prop_Has_Garage = sum(Has.Garage == "Yes",na.rm = TRUE) / n(),
    Prop_Returning_Customer = sum(Returning.Customer == "Yes", na.rm = TRUE) / n(),
    .groups = 'drop'
  )

# Summarize by Macro Region
df_by_macro_region <- df_unique_high %>%
  group_by(Macro_Region) %>%
  summarise(
    Sample_Size = n(),
    Avg_Bedrooms = mean(Bedrooms, na.rm = TRUE),
    Avg_Bathrooms = mean(Bathrooms, na.rm = TRUE),
    Avg_Price_Per_sqft = mean(Price.Per.sqft, na.rm = TRUE),
    Avg_Property_Age = mean(Property.Age, na.rm = TRUE),
    Avg_Total_Structure_Area = mean(Total.Structure.Area.sqft, na.rm = TRUE),
    Avg_Months_Since_Purchase = mean(Months.Since.Purchase, na.rm = TRUE),
    Avg_Aggregate_Sub_Total = mean(Aggregate.Sub.Total, na.rm = TRUE),
    Prop_Has_HOA = sum(Has.HOA == "Yes", na.rm = TRUE) / n(),
    Prop_Has_Garage = sum(Has.Garage == "Yes",na.rm = TRUE) / n(),
    Prop_Returning_Customer = sum(Returning.Customer == "Yes", na.rm = TRUE) / n(),
    .groups = 'drop'
  )

## Visualizations by Region

df_by_macro_region <- df_by_macro_region %>% rename("N.of.Customers" = Sample_Size)
df_by_region <- df_by_region %>% rename("N.of.Customers" = Sample_Size)
df_by_state_filtered <- df_by_state_filtered %>% rename("N.of.Customers" = Sample_Size)

# List of variables to plot
vars_to_plot <- c("N.of.Customers", "Avg_Bedrooms", "Avg_Bathrooms", "Avg_Price_Per_sqft", 
                  "Avg_Property_Age", "Avg_Total_Structure_Area", "Avg_Months_Since_Purchase", 
                  "Avg_Aggregate_Sub_Total", "Prop_Has_HOA", "Prop_Has_Garage", "Prop_Returning_Customer")

# Function to create bar plots
create_sorted_bar_plot <- function(df, var, group_var) {
  df <- df %>% 
    arrange(!!sym(var)) %>% 
    mutate(!!sym(group_var) := factor(!!sym(group_var), levels = unique(!!sym(group_var))))
  
  p <- ggplot(df, aes_string(x = var, y = group_var, fill = var)) +
    geom_bar(stat = "identity") +
    scale_fill_gradient(low = "lightblue", high = "darkblue") +
    theme_minimal() +
    ylab(group_var) +
    xlab(var) +
    ggtitle(paste("Sorted Bar Plot of", var, "by", group_var))
  
  return(p)
}

plot_list_by_state <- list()
plot_list_by_region <- list()
plot_list_by_macro_region <- list()

for (var in vars_to_plot) {
  plot_list_by_state[[var]] <- create_sorted_bar_plot(df_by_state_filtered, var, "State")
  plot_list_by_region[[var]] <- create_sorted_bar_plot(df_by_region, var, "Region")
  plot_list_by_macro_region[[var]] <- create_sorted_bar_plot(df_by_macro_region, var, "Macro_Region")
}

save_plots <- function(plot_list, directory) {
  if (!dir.exists(directory)) {
    dir.create(directory)
  }
  
  for (var in names(plot_list)) {
    ggsave(paste0(directory, "/", var, ".png"), plot = plot_list[[var]], width = 10, height = 8)
  }
}

save_plots(plot_list_by_state, "State_Plots")
save_plots(plot_list_by_region, "Region_Plots")
save_plots(plot_list_by_macro_region, "Macro_Region_Plots")


# Export Tables
# Create a new workbook
wb <- createWorkbook()

# Add data frames as sheets
addWorksheet(wb, "Cluster Profile")
writeData(wb, "Cluster Profile", cluster_profile)

addWorksheet(wb, "Cluster Sizes")
writeData(wb, "Cluster Sizes", cluster_sizes)

addWorksheet(wb, "Common States")
writeData(wb, "Common States", common_states)

addWorksheet(wb, "Sales Stats per Cluster")
writeData(wb, "Sales Stats per Cluster", df_sales_stats_perCluster)

addWorksheet(wb, "State Frequency")
writeData(wb, "State Frequency", state_frequency)

addWorksheet(wb, "State Frequency Grouped")
writeData(wb, "State Frequency Grouped", state_frequency_grouped)

addWorksheet(wb, "Top States per Cluster")
writeData(wb, "Top States per Cluster", top_states_per_cluster)

# Save workbook to a file
saveWorkbook(wb, "Cluster_Analysis_Output.xlsx", overwrite = TRUE)

write.csv(df_unique_high_filtered, "Data_wClusters.csv")


# Export Data with Regions

# Create a new workbook
wb <- createWorkbook()

addWorksheet(wb, "By State")
writeData(wb, "By State", df_by_state_filtered)

addWorksheet(wb, "By Macro Region")
writeData(wb, "By Macro Region", df_by_macro_region)

addWorksheet(wb, "By Region")
writeData(wb, "By Region", df_by_region)

# Save workbook to a file
saveWorkbook(wb, "Region_Analysis_Output.xlsx", overwrite = TRUE)



## CREATE SEGMENT VALUE INDEX

# Convert 'Yes' to 1 and 'No' to 0 in the Prop_Returning_Customer column
df_unique_high$Returning.Customer <- ifelse(df_unique_high$Returning.Customer == "Yes", 1, 
                                                     ifelse(df_unique_high$Returning.Customer == "No", 0, df_unique_high$Returning.Customer))

# Calculate the min and max for each variable, ignoring NAs
min_sub_total <- min(df_unique_high$Aggregate.Sub.Total, na.rm = TRUE)
max_sub_total <- max(df_unique_high$Aggregate.Sub.Total, na.rm = TRUE)

min_months_since_purchase <- min(df_unique_high$Months.Since.Purchase)
max_months_since_purchase <- max(df_unique_high$Months.Since.Purchase)

min_returning_customer <- min(df_unique_high$Returning.Customer)
max_returning_customer <- max(df_unique_high$Returning.Customer)

# Standardize the variables
std_sub_total <- scale(df_unique_high$Aggregate.Sub.Total, center = TRUE, scale = TRUE)
std_months_since_purchase <- scale(df_unique_high$Months.Since.Purchase, center = TRUE, scale = TRUE)
std_returning_customer <- scale(df_unique_high$Returning.Customer, center = TRUE, scale = TRUE)

# Function to standardize a vector, handling NAs
standardize_with_na <- function(x) {
  na_mask <- is.na(x)
  x_mean <- mean(x, na.rm = TRUE)
  x_sd <- sd(x, na.rm = TRUE)
  x[!na_mask] <- (x[!na_mask] - x_mean) / x_sd
  return(x)
}

# Standardize the variables
std_sub_total <- standardize_with_na(df_unique_high$Aggregate.Sub.Total)
std_months_since_purchase <- standardize_with_na(df_unique_high$Months.Since.Purchase)
std_returning_customer <- standardize_with_na(df_unique_high$Returning.Customer)

# Calculate CVI, handling NAs
df_unique_high$CVI <- (1/3) * std_sub_total + 
  (1/3) * (-std_months_since_purchase) + 
  (1/3) * std_returning_customer

# Rescale to a 0-100 scale if needed, again handling NAs
if (any(!is.na(df_unique_high$CVI))) {
  df_unique_high$CVI <- (df_unique_high$CVI - min(df_unique_high$CVI, na.rm = TRUE)) / 
    (max(df_unique_high$CVI, na.rm = TRUE) - min(df_unique_high$CVI, na.rm = TRUE)) * 100
}

# Calculate CVI2, handling NAs
df_unique_high$CVI2 <- 0.5 * std_sub_total +
  0.5 * (-std_months_since_purchase)

# Rescale to a 0-100 scale if needed, again handling NAs
if (any(!is.na(df_unique_high$CVI2))) {
  df_unique_high$CVI2 <- (df_unique_high$CVI2 - min(df_unique_high$CVI2, na.rm = TRUE)) / 
    (max(df_unique_high$CVI2, na.rm = TRUE) - min(df_unique_high$CVI2, na.rm = TRUE)) * 100
}


## Add mean to the state summary CVI and cluster summary

# Calculate mean CVI and CVI2 for each state in df_unique_high
state_level_means <- df_unique_high %>%
  group_by(State) %>%
  summarise(Mean_CVI = mean(CVI, na.rm = TRUE),
            Mean_CVI2 = mean(CVI2, na.rm = TRUE))

# Add state-level mean values back to df_by_state_filtered
df_by_state_filtered <- left_join(df_by_state_filtered, state_level_means, by = "State")

### F-TEST

# Filter out states with more than 5 cases
state_counts <- table(df_unique_high$State)
states_with_more_than_5 <- names(state_counts[state_counts > 5])
df_filtered_states <- df_unique_high %>% filter(State %in% states_with_more_than_5)

# Change only if the value is 'Yes' or 'No' for df_filtered_states
df_filtered_states$Has.HOA <- ifelse(df_filtered_states$Has.HOA == "Yes", 1, 
                                     ifelse(df_filtered_states$Has.HOA == "No", 0, df_filtered_states$Has.HOA))

df_filtered_states$Has.Garage <- ifelse(df_filtered_states$Has.Garage == "Yes", 1, 
                                        ifelse(df_filtered_states$Has.Garage == "No", 0, df_filtered_states$Has.Garage))

# Change only if the value is 'Yes' or 'No' for df_unique_high
df_unique_high$Has.HOA <- ifelse(df_unique_high$Has.HOA == "Yes", 1, 
                                          ifelse(df_unique_high$Has.HOA == "No", 0, df_unique_high$Has.HOA))

df_unique_high$Has.Garage <- ifelse(df_unique_high$Has.Garage == "Yes", 1, 
                                             ifelse(df_unique_high$Has.Garage == "No", 0, df_unique_high$Has.Garage))


# Mapping between df_by_state_filtered and df_unique_high column names
state_col_mapping <- list(
  Avg_Bedrooms = "Bedrooms",
  Avg_Bathrooms = "Bathrooms",
  Avg_Price_Per_sqft = "Price.Per.sq.ft",
  Avg_Property_Age = "Property.Age",
  Avg_Total_Structure_Area = "Total.Structure.Area.sqft",
  Avg_Months_Since_Purchase = "Months.Since.Purchase",
  Avg_Aggregate_Sub_Total = "Aggregate.Sub.Total",
  Prop_Has_HOA = "Has.HOA",
  Prop_Has_Garage = "Has.Garage",
  Prop_Returning_Customer = "Returning.Customer"
)

# Add CVI and CVI2 to the mapping list
state_col_mapping <- c(state_col_mapping, list(
  Mean_CVI = "CVI",
  Mean_CVI2 = "CVI2"
))

# Initialize data frames to store F-values and p-values for df_by_state_filtered
f_values_state_df <- data.frame(Variable = names(state_col_mapping), F_value = NA, P_value = NA)

# Calculate F-values and p-values for df_by_state_filtered
for (col in names(state_col_mapping)) {
  # Filter out rows where either the variable or 'State' is NA
  filtered_data <- df_filtered_states[!is.na(df_filtered_states[[state_col_mapping[[col]]]]) & !is.na(df_filtered_states$State), ]
  
  # Run the ANOVA on the filtered data
  fit <- aov(filtered_data[[state_col_mapping[[col]]]] ~ filtered_data$State)
  summary_fit <- summary(fit)
  
  # Store F-value and p-value if they exist
  if (length(summary_fit[[1]]) > 0) {
    f_values_state_df[f_values_state_df$Variable == col, "F_value"] <- summary_fit[[1]][["F value"]][1]
    f_values_state_df[f_values_state_df$Variable == col, "P_value"] <- summary_fit[[1]][["Pr(>F)"]][1]
  }
}

# Show the F-values and p-values data frames
print("F-values and p-values for df_by_state_filtered:")
print(f_values_state_df)

library(dplyr)
library(purrr)

# Function to calculate the required statistics for a given column
calc_stats <- function(col) {
  data <- df_unique_high[[col]]
  data <- na.omit(data)  # Remove NA values
  return(c(
    mean = mean(data),
    sd = sd(data),
    p10 = quantile(data, 0.1),
    p15 = quantile(data, 0.15),
    p25 = quantile(data, 0.25),
    p50 = quantile(data, 0.5),
    p75 = quantile(data, 0.75),
    p85 = quantile(data, 0.85),
    p90 = quantile(data, 0.9)
  ))
}

# Calculate the statistics for each mapped variable
stats_df <- map_df(state_col_mapping, calc_stats)

# Transpose the data frame for better readability and set row names to the mapped variables
stats_df <- as.data.frame(t(stats_df))
colnames(stats_df) <- names(state_col_mapping)

print(stats_df)

library(openxlsx)

# Load the existing workbook
wb <- loadWorkbook("Region_Analysis_Output.xlsx")

# Create a new worksheet for f_values_state_df and add the data
addWorksheet(wb, "F_Values_State")
writeData(wb, sheet = "F_Values_State", x = f_values_state_df)

# Create a new worksheet for stats_df and add the data
addWorksheet(wb, "Stats")
writeData(wb, sheet = "Stats", x = stats_df, rowNames = TRUE)

# Save the workbook
saveWorkbook(wb, "Region_Analysis_Output.xlsx", overwrite = TRUE)


