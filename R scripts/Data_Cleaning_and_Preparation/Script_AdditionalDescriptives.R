library(dplyr)
library(lubridate)

# Filter for members active in January 2018
cohort_2021 <- df %>%
  filter(year(reporting_month) == 2021 & month(reporting_month) == 8) %>%
  dplyr::select(case_mbr_key) %>%
  distinct()

# Filter the main dataset to include only these members and data from 2021 onwards
df_cohort <- df %>%
  filter(case_mbr_key %in% cohort_2021$case_mbr_key & year(reporting_month) >= 2021)

# Create a summary of active customers by year and month for the cohort
active_customers_summary_cohort <- df_cohort %>%
  group_by(year, month) %>%
  summarise(Active_Customers = n_distinct(case_mbr_key), .groups = 'drop')

# Join this summary back to the main df_cohort based on year and month
df_cohort <- df_cohort %>%
  left_join(active_customers_summary_cohort, by = c("year", "month"))

df_cohort <- df_cohort %>%
  mutate(month = as.numeric(month))

# Adding flags for communication changes in the cohort
df_cohort <- df_cohort %>%
  arrange(case_mbr_key, year, month) %>%
  group_by(case_mbr_key) %>%
  mutate(
    Change_2020 = as.integer((year == 2020 & month >= 8) | (year == 2021 & month < 8)),
    Change_2021 = as.integer((year == 2021 & month >= 8) | (year == 2022 & month < 8)),
    Change_2022 = as.integer((year == 2022 & month >= 8) | (year == 2023 & month < 8)),
    Change_2023 = as.integer(year >= 2023 & month >= 8)
  ) %>%
  ungroup()


# Log transformation for the cohort
df_cohort <- df_cohort %>%
  mutate(log_retirement_contribution_rate = log(retirement_contribution_rate + 1),
         log_annual_pensionable_salary = log(annual_pensionable_salary + 1))


# Create flags for contribution rate change: increased, decreased, or stayed the same with hierarchical checks
df_cohort <- df_cohort %>%
  group_by(case_mbr_key) %>%
  arrange(case_mbr_key, reporting_month) %>%
  mutate(
    # Calculate the change in contribution rate
    contribution_rate_change = retirement_contribution_rate - lag(retirement_contribution_rate),
    # Identify increases
    contribution_increased = ifelse(contribution_rate_change > 0, 1, 0),
    # Include the identification in the row before and after the change
    contribution_increased_lag1 = lag(contribution_increased, default = 0),
    contribution_increased_lead1 = lead(contribution_increased, default = 0),
    contribution_increased_final = pmax(contribution_increased, contribution_increased_lag1, contribution_increased_lead1)
  ) %>%
  mutate(
    # Identify decreases if not increased
    contribution_decreased = ifelse(contribution_increased_final == 0 & contribution_rate_change < 0, 1, 0),
    contribution_decreased_lag1 = lag(contribution_decreased, default = 0),
    contribution_decreased_lead1 = lead(contribution_decreased, default = 0),
    contribution_decreased_final = pmax(contribution_decreased, contribution_decreased_lag1, contribution_decreased_lead1)
  ) %>%
  mutate(
    # Identify stayed same if neither increased nor decreased
    contribution_stayed_same_final = ifelse(contribution_increased_final == 0 & contribution_decreased_final == 0, 1, 0)
  ) %>%
  mutate(
    # Calculate final contribution rate change considering possible lags
    contribution_rate_change_lag1 = lag(contribution_rate_change, default = 0),
    contribution_rate_change_lead1 = lead(contribution_rate_change, default = 0),
    contribution_rate_change_final = ifelse(
      contribution_increased_final == 1 | contribution_decreased_final == 1,
      pmax(contribution_rate_change, contribution_rate_change_lag1, contribution_rate_change_lead1, na.rm = TRUE),
      0
    )
  ) %>%
  ungroup()

# Replace NAs with 0 for the first observation of each group where there is no lag value
df_cohort <- df_cohort %>%
  mutate(
    contribution_increased_final = ifelse(is.na(contribution_increased_final), 0, contribution_increased_final),
    contribution_decreased_final = ifelse(is.na(contribution_decreased_final), 0, contribution_decreased_final),
    contribution_stayed_same_final = ifelse(is.na(contribution_stayed_same_final), 0, contribution_stayed_same_final),
    contribution_rate_change_final = ifelse(is.na(contribution_rate_change_final), 0, contribution_rate_change_final)
  )

# Create the Change_ dummy variables
df_cohort <- df_cohort %>%
  mutate(
    Change_2021 = ifelse((year == 2019 & month == 8) | (year == 2020 & month == 4), 1, 0),
    Change_2022 = ifelse((year == 2022 & month == 8) | (year == 2023 & month == 4), 1, 0),
    Change_2023 = ifelse((year == 2023 & month == 8) | (year == 2024 & month == 4), 1, 0)
  )

# Create separate Change dummy variables for April and August
df_cohort <- df_cohort %>%
  mutate(
    Change_April_2021 = ifelse(year == 2020 & month == 4, 1, 0),
    Change_August_2021 = ifelse(year == 2019 & month == 8, 1, 0),
    Change_April_2022 = ifelse(year == 2023 & month == 4, 1, 0),
    Change_August_2022 = ifelse(year == 2022 & month == 8, 1, 0),
    Change_April_2023 = ifelse(year == 2024 & month == 4, 1, 0),
    Change_August_2023 = ifelse(year == 2023 & month == 8, 1, 0)
  )

# Filter to keep only the data for April and August
df_august_april <- df_cohort %>%
  filter(month %in% c(4, 8))


# Filter out August 2021 from df_cohort
df_cohort_no_apr_2021 <- df_cohort %>%
  filter(!(year(reporting_month) == 2021 & month(reporting_month) == 4))

# Filter to keep only the data for April and August (excluding August 2021)
df_august_april <- df_cohort_no_apr_2021 %>%
  filter(month(reporting_month) %in% c(4, 8))


library(dplyr)

# Filter data to include only those who increased their contribution rate
df_increased <- df_august_april %>%
  filter(contribution_increased_final == 1)

library(dplyr)

newdescriptive_table <- df_august_april %>%
  group_by(reporting_month) %>%
  summarise(
    total = n(),
    increased = sum(contribution_increased_final),
    decreased = sum(contribution_decreased_final),
    stayed_same = sum(contribution_stayed_same_final)
  ) %>%
  left_join(
    df_increased %>%
      group_by(reporting_month) %>%
      summarise(
        N = n(),
        mean_contribution_change = mean(contribution_rate_change_final, na.rm = TRUE),
        median_contribution_change = median(contribution_rate_change_final, na.rm = TRUE),
        sd_contribution_change = sd(contribution_rate_change_final, na.rm = TRUE),
        iqr_contribution_change = IQR(contribution_rate_change_final, na.rm = TRUE),
        min_contribution_change = min(contribution_rate_change_final, na.rm = TRUE),
        max_contribution_change = max(contribution_rate_change_final, na.rm = TRUE)
      ),
    by = "reporting_month"
  ) %>%
  mutate(
    perc_increased = (increased / total) * 100,
    perc_decreased = (decreased / total) * 100,
    perc_stayed_same = (stayed_same / total) * 100
  ) %>%
  dplyr::select(reporting_month, N, perc_increased, perc_decreased, perc_stayed_same,
                mean_contribution_change, median_contribution_change, sd_contribution_change,
                iqr_contribution_change, min_contribution_change, max_contribution_change)


# Create boxplots of the contribution changes for those that increased their contribution rates
boxplot <- ggplot(df_increased, aes(x = as.factor(reporting_month), y = contribution_rate_change_final)) +
  geom_boxplot() +
  labs(title = "Boxplots of Contribution Changes for Increased Contributions",
       x = "Reporting Month",
       y = "Contribution Rate Change (%)") +
  theme_minimal() +
  ylim(0, 15)

# Display the boxplot
print(boxplot)










# Adding flags for communication changes in the cohort
df <- df %>%
  arrange(case_mbr_key, year, month) %>%
  group_by(case_mbr_key) %>%
  mutate(
    Change_2020 = as.integer((year == 2020 & month >= 8) | (year == 2021 & month < 8)),
    Change_2021 = as.integer((year == 2021 & month >= 8) | (year == 2022 & month < 8)),
    Change_2022 = as.integer((year == 2022 & month >= 8) | (year == 2023 & month < 8)),
    Change_2023 = as.integer(year >= 2023 & month >= 8)
  ) %>%
  ungroup()


# Log transformation for the cohort
df <- df %>%
  mutate(log_retirement_contribution_rate = log(retirement_contribution_rate + 1),
         log_annual_pensionable_salary = log(annual_pensionable_salary + 1))


# Create flags for contribution rate change: increased, decreased, or stayed the same with hierarchical checks
df <- df %>%
  group_by(case_mbr_key) %>%
  arrange(case_mbr_key, reporting_month) %>%
  mutate(
    # Calculate the change in contribution rate
    contribution_rate_change = retirement_contribution_rate - lag(retirement_contribution_rate),
    # Identify increases
    contribution_increased = ifelse(contribution_rate_change > 0, 1, 0),
    # Include the identification in the row before and after the change
    contribution_increased_lag1 = lag(contribution_increased, default = 0),
    contribution_increased_lead1 = lead(contribution_increased, default = 0),
    contribution_increased_final = pmax(contribution_increased, contribution_increased_lag1, contribution_increased_lead1)
  ) %>%
  mutate(
    # Identify decreases if not increased
    contribution_decreased = ifelse(contribution_increased_final == 0 & contribution_rate_change < 0, 1, 0),
    contribution_decreased_lag1 = lag(contribution_decreased, default = 0),
    contribution_decreased_lead1 = lead(contribution_decreased, default = 0),
    contribution_decreased_final = pmax(contribution_decreased, contribution_decreased_lag1, contribution_decreased_lead1)
  ) %>%
  mutate(
    # Identify stayed same if neither increased nor decreased
    contribution_stayed_same_final = ifelse(contribution_increased_final == 0 & contribution_decreased_final == 0, 1, 0)
  ) %>%
  mutate(
    # Calculate final contribution rate change considering possible lags
    contribution_rate_change_lag1 = lag(contribution_rate_change, default = 0),
    contribution_rate_change_lead1 = lead(contribution_rate_change, default = 0),
    contribution_rate_change_final = ifelse(
      contribution_increased_final == 1 | contribution_decreased_final == 1,
      pmax(contribution_rate_change, contribution_rate_change_lag1, contribution_rate_change_lead1, na.rm = TRUE),
      0
    )
  ) %>%
  ungroup()

# Replace NAs with 0 for the first observation of each group where there is no lag value
df <- df %>%
  mutate(
    contribution_increased_final = ifelse(is.na(contribution_increased_final), 0, contribution_increased_final),
    contribution_decreased_final = ifelse(is.na(contribution_decreased_final), 0, contribution_decreased_final),
    contribution_stayed_same_final = ifelse(is.na(contribution_stayed_same_final), 0, contribution_stayed_same_final),
    contribution_rate_change_final = ifelse(is.na(contribution_rate_change_final), 0, contribution_rate_change_final)
  )

# Create the Change_ dummy variables
df <- df %>%
  mutate(
    Change_2021 = ifelse((year == 2019 & month == 8) | (year == 2020 & month == 4), 1, 0),
    Change_2022 = ifelse((year == 2022 & month == 8) | (year == 2023 & month == 4), 1, 0),
    Change_2023 = ifelse((year == 2023 & month == 8) | (year == 2024 & month == 4), 1, 0)
  )

# Create separate Change dummy variables for April and August
df <- df %>%
  mutate(
    Change_April_2021 = ifelse(year == 2020 & month == 4, 1, 0),
    Change_August_2021 = ifelse(year == 2019 & month == 8, 1, 0),
    Change_April_2022 = ifelse(year == 2023 & month == 4, 1, 0),
    Change_August_2022 = ifelse(year == 2022 & month == 8, 1, 0),
    Change_April_2023 = ifelse(year == 2024 & month == 4, 1, 0),
    Change_August_2023 = ifelse(year == 2023 & month == 8, 1, 0)
  )

# Filter to keep only the data for April and August
df_august_april <- df %>%
  filter(month %in% c(4, 8))


library(dplyr)

# Filter data to include only those who increased their contribution rate
df_increased <- df_august_april %>%
  filter(contribution_increased_final == 1)

# Create a summary table with additional statistics for those who increased their contribution rate
newdescriptive_table_all <- df_august_april %>%
  group_by(reporting_month) %>%
  summarise(
    total = n(),
    increased = sum(contribution_increased_final),
    decreased = sum(contribution_decreased_final),
    stayed_same = sum(contribution_stayed_same_final)
  ) %>%
  left_join(
    df_increased %>%
      group_by(reporting_month) %>%
      summarise(
        N = n(),
        mean_contribution_change = mean(contribution_rate_change_final, na.rm = TRUE),
        median_contribution_change = median(contribution_rate_change_final, na.rm = TRUE),
        sd_contribution_change = sd(contribution_rate_change_final, na.rm = TRUE),
        iqr_contribution_change = IQR(contribution_rate_change_final, na.rm = TRUE),
        min_contribution_change = min(contribution_rate_change_final, na.rm = TRUE),
        max_contribution_change = max(contribution_rate_change_final, na.rm = TRUE)
      ),
    by = "reporting_month"
  ) %>%
  mutate(
    perc_increased = (increased / total) * 100,
    perc_decreased = (decreased / total) * 100,
    perc_stayed_same = (stayed_same / total) * 100
  ) %>%
  select(reporting_month, N, perc_increased, perc_decreased, perc_stayed_same,
         mean_contribution_change, median_contribution_change, sd_contribution_change,
         iqr_contribution_change, min_contribution_change, max_contribution_change)



# Create boxplots of the contribution changes for those that increased their contribution rates
boxplot <- ggplot(df_increased, aes(x = as.factor(reporting_month), y = contribution_rate_change_final)) +
  geom_boxplot() +
  labs(title = "Boxplots of Contribution Changes for Increased Contributions - All Sample",
       x = "Reporting Month",
       y = "Contribution Rate Change") +
  theme_minimal() +
  ylim(0, 15)

# Display the boxplot
print(boxplot)


library(ggplot2)

# Create histograms of the contribution changes for those that increased their contribution rates, faceted by reporting month
histogram <- ggplot(df_increased, aes(x = contribution_rate_change_final)) +
  geom_histogram(binwidth = 0.5, fill = "blue", color = "black", alpha = 0.7) +
  labs(title = "Histograms of Contribution Changes for Increased Contributions by Reporting Month",
       x = "Contribution Rate Change (%)",
       y = "Frequency") +
  theme_minimal() +
  scale_x_continuous(breaks = seq(0, 15, by = 1)) +
  facet_wrap(~ reporting_month, scales = "free_y") +
  ylim(0, 20) # Set the y-axis limits

# Display the histogram
print(histogram)


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
data_list_cohort <- list(
  "Cohort Sample" = newdescriptive_table
)

# Save to Excel with APA formatting for the cohort
save_apa_formatted_excel(data_list_cohort, "APA_Formatted_Tables_v7.xlsx")

