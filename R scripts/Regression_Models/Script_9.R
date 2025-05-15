# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/aminsakha")

# Load the necessary library
library(openxlsx)
library(readxl)

# List of files to import
files <- list(
  "Max Temperature" = "1. Max Temperature 2000-2024.xls",
  "Min Temperature" = "2. Min Temperature 2000-2024.xls",
  "Average Temperature" = "3. Averg Temperature 2000-2024.xls",
  "Rainfall" = "4. Rainfall 2000-2024.xlsx",
  "Relative Humidity" = "5. Relative Humidity 2000-2024.xls",
  "Wind Speed" = "6. Wind speed 2000-2024.xls"
)

library(dplyr)
library(tidyr)
library(purrr)

datasets <- list(
  "Max Temperature" = read_excel(files[["Max Temperature"]]) %>%
    rename(Station = STATION, Date = Date, Max_Temp = `MAX TEMP`) %>%
    select(Station, Date, Max_Temp),
  "Min Temperature" = read_excel(files[["Min Temperature"]]) %>%
    rename(Station = STATION, Date = Date, Min_Temp = `MIN TEMP`) %>%
    select(Station, Date, Min_Temp),
  "Average Temperature" = read_excel(files[["Average Temperature"]]) %>%
    rename(Station = `Station Code`, Date = Date, Avg_Temp = `AVG_DAILY_TEMP`) %>%
    select(Station, Date, Avg_Temp),
  "Rainfall" = read_excel(files[["Rainfall"]]) %>%
    rename(Station = STATION, Date = Date, Rainfall_mm = `Rainfall (mm)`, Rainy_Days = `No. of Rainy Days`) %>%
    select(Station, Date, Rainfall_mm, Rainy_Days),
  "Relative Humidity" = read_excel(files[["Relative Humidity"]]) %>%
    rename(Station = Station, Date = Date, Daily_RH = `Daily RH(%)`) %>%
    select(Station, Date, Daily_RH),
  "Wind Speed" = read_excel(files[["Wind Speed"]]) %>%
    rename(Station = Station, Date = Date, Wind_Speed = `Wind Speed(kt)`) %>%
    select(Station, Date, Wind_Speed)
)

# Merge datasets by 'Date' and 'Station'
merged_data <- reduce(datasets, function(x, y) {
  full_join(x, y, by = c("Station", "Date"))
})

# Convert 'Date' column to Date type
merged_data <- merged_data %>%
  mutate(Date = as.Date(Date, format = "%d/%m/%Y"))

# View the consolidated dataset
print(head(merged_data))


# Replace NA with 0 in Rainfall_mm and Rainy_Days columns
merged_data <- merged_data %>%
  mutate(
    Rainfall_mm = ifelse(is.na(Rainfall_mm), 0, Rainfall_mm),
    Rainy_Days = ifelse(is.na(Rainy_Days), 0, Rainy_Days)
  )

# Calculate the percentage of filled values by station
filled_values_summary <- merged_data %>%
  group_by(Station) %>%
  summarise(
    Max_Temp_Filled = 100 * mean(!is.na(Max_Temp)),
    Min_Temp_Filled = 100 * mean(!is.na(Min_Temp)),
    Avg_Temp_Filled = 100 * mean(!is.na(Avg_Temp)),
    Rainfall_mm_Filled = 100 * mean(!is.na(Rainfall_mm)),
    Rainy_Days_Filled = 100 * mean(!is.na(Rainy_Days)),
    Daily_RH_Filled = 100 * mean(!is.na(Daily_RH)),
    Wind_Speed_Filled = 100 * mean(!is.na(Wind_Speed))
  )


library(dplyr)
library(zoo)

# Impute missing data for non-rainfall variables
impute_by_neighbors <- function(df, columns) {
  df %>%
    group_by(Station) %>%
    mutate(across(
      all_of(columns),
      ~ na.approx(., rule = 2, na.rm = FALSE), # Linear interpolation between neighbors
      .names = "imputed_{col}"
    )) %>%
    ungroup()
}

# Impute missing rainfall data using monthly means
impute_rainfall_monthly_mean <- function(df) {
  df %>%
    mutate(Month = format(Date, "%m")) %>%  # Extract month from Date
    group_by(Station, Month) %>%
    mutate(
      Rainfall_mm = ifelse(
        is.na(Rainfall_mm), 
        mean(Rainfall_mm, na.rm = TRUE),  # Replace NA with monthly mean
        Rainfall_mm
      ),
      Rainy_Days = ifelse(
        is.na(Rainy_Days),
        mean(Rainy_Days, na.rm = TRUE),  # Replace NA with monthly mean for rainy days
        Rainy_Days
      )
    ) %>%
    ungroup() %>%
    select(-Month)  # Remove temporary Month column
}

# Identify substantial gaps
report_gaps <- function(df, columns) {
  df %>%
    group_by(Station) %>%
    summarise(across(
      all_of(columns),
      ~ {
        runs <- rle(is.na(.)) # Identify runs of NA values
        max_gap <- ifelse(any(runs$values), max(runs$lengths[runs$values]), 0) # Largest gap
        total_gaps <- sum(runs$values * runs$lengths) # Total number of missing rows
        list(Max_Gap_Length = max_gap, Total_Missing = total_gaps)
      },
      .names = "{col}_{.fn}"
    )) %>%
    pivot_longer(
      cols = -Station,
      names_to = c("Variable", "Metric"),
      names_sep = "_",
      values_to = "Value"
    ) %>%
    pivot_wider(
      names_from = Metric,
      values_from = Value
    )
}

# Impute rainfall using monthly means
imputed_rainfall_data <- impute_rainfall_monthly_mean(merged_data)

# Impute other variables using neighbor interpolation
columns_to_impute <- c("Max_Temp", "Min_Temp", "Avg_Temp", "Daily_RH")
imputed_data <- impute_by_neighbors(imputed_rainfall_data, columns_to_impute)

# Generate report on gaps
gaps_report <- report_gaps(merged_data, c("Max_Temp", "Min_Temp", "Avg_Temp", "Rainfall_mm", "Rainy_Days", "Daily_RH"))

# Function to calculate and impute Avg_Temp
impute_avg_temp <- function(df) {
  df %>%
    group_by(Station) %>%
    mutate(
      # Step 1: Calculate Avg_Temp from Min_Temp and Max_Temp if available
      ImputedAvg_Temp = ifelse(
        is.na(Avg_Temp) & !is.na(Min_Temp) & !is.na(Max_Temp),
        (Min_Temp + Max_Temp) / 2,
        Avg_Temp
      ),
      # Step 2: Fill remaining NA values using neighbor interpolation
      Avg_Temp = na.approx(Avg_Temp, na.rm = FALSE)
    ) %>%
    ungroup()
}

# Update the imputed_data dataset
imputed_data <- impute_avg_temp(imputed_data)

# Validate imputation by checking summary statistics
validation_summary <- imputed_data %>%
  summarise(
    Total_Rainfall = sum(Rainfall_mm, na.rm = TRUE),
    Mean_Rainfall = mean(Rainfall_mm, na.rm = TRUE),
    Missing_Rainfall_Percentage = 100 * mean(is.na(Rainfall_mm)),
    Total_Rainy_Days = sum(Rainy_Days, na.rm = TRUE),
    Missing_Rainy_Days_Percentage = 100 * mean(is.na(Rainy_Days))
  )
print(validation_summary)


library(dplyr)
library(lubridate)
library(ggplot2)

# Add 'Year' and 'Season' columns based on exact seasonal dates
imputed_data <- imputed_data %>%
  mutate(
    Date = as.Date(Date),
    Year = year(Date),
    Season = case_when(
      Date >= as.Date(paste(Year, "12-21", sep = "-")) | 
        Date <= as.Date(paste(Year, "03-20", sep = "-")) ~ "Winter",
      Date >= as.Date(paste(Year, "03-21", sep = "-")) & 
        Date <= as.Date(paste(Year, "06-20", sep = "-")) ~ "Spring",
      Date >= as.Date(paste(Year, "06-21", sep = "-")) & 
        Date <= as.Date(paste(Year, "09-22", sep = "-")) ~ "Summer",
      Date >= as.Date(paste(Year, "09-23", sep = "-")) & 
        Date <= as.Date(paste(Year, "12-20", sep = "-")) ~ "Autumn"
    )
  )

# Filter data up to 2021
imputed_data <- imputed_data %>%
  filter(Date <= as.Date("2021-12-11"))

daily_avg_rainfall <- imputed_data %>%
  group_by(Date) %>%
  summarise(
    Avg_Rainfall_mm = mean(Rainfall_mm, na.rm = TRUE), # Average rainfall across stations
    Rainy_Day = ifelse(any(Rainfall_mm > 0, na.rm = TRUE), 1, 0), # At least one station with rainfall
    .groups = "drop"
  )

# Step 2: Merge the averaged rainfall back with the original dataset
imputed_data <- imputed_data %>%
  left_join(daily_avg_rainfall, by = "Date")


# Update annual and seasonal summaries to include humidity
seasonal_summary <- imputed_data %>%
  group_by(Year, Season) %>%
  summarise(
    Avg_Max_Temp = mean(Max_Temp, na.rm = TRUE),
    Avg_Min_Temp = mean(Min_Temp, na.rm = TRUE),
    Avg_Avg_Temp = mean(Avg_Temp, na.rm = TRUE),
    Total_Rainfall = sum(Avg_Rainfall_mm, na.rm = TRUE) / 4, # Divide total by 4 (number of stations)
    Rainy_Days = sum(Rainy_Day, na.rm = TRUE)/4,              # Count days with at least one rainy station
    Avg_Humidity = mean(Daily_RH, na.rm = TRUE),            # Average Humidity
    .groups = "drop"
  )

# Plot with shaded area, thinner lines for Min and Max, and data labels for Avg
ggplot(seasonal_summary, aes(x = Year, group = Season, fill = Season)) +
  # Shaded area between Min and Max
  geom_ribbon(aes(ymin = Avg_Min_Temp, ymax = Avg_Max_Temp), alpha = 0.2) +
  # Line for average temperature
  geom_line(aes(y = Avg_Avg_Temp, color = Season), size = 1.2) +
  # Thinner lines for Min and Max temperatures
  geom_line(aes(y = Avg_Min_Temp, color = Season), size = 0.5, linetype = "dashed") +
  geom_line(aes(y = Avg_Max_Temp, color = Season), size = 0.5, linetype = "dashed") +
  # Add data labels to the average temperature line
  geom_text(aes(y = Avg_Avg_Temp, label = round(Avg_Avg_Temp, 1)), 
            color = "black", size = 3, vjust = -1, check_overlap = TRUE) +
  # Customize X-axis with yearly breaks
  scale_x_continuous(breaks = seq(2000, 2024, by = 1)) +
  # Labels and title
  labs(
    title = "Seasonal Temperature Trends (2000–2021)",
    x = "Year",
    y = "Temperature (°C)",
    color = "Season",
    fill = "Season"
  ) +
  # Theme for professional appearance
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey90"),
    plot.title = element_text(hjust = 0.5, size = 16)
  )



# Plot seasonal total rainfall trends
ggplot(seasonal_summary, aes(x = factor(Year), y = Total_Rainfall, fill = Season)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), alpha = 0.8) +
  # Add data labels at the center top of the bars
  geom_text(
    aes(
      y = Total_Rainfall, 
      label = round(Total_Rainfall, 1)
    ), 
    position = position_dodge(width = 0.8), 
    vjust = -0.5, # Slightly above the top of the bar
    color = "black", 
    size = 3
  ) +
  scale_x_discrete(breaks = seq(2000, 2024, by = 1)) + # Year as discrete axis for cleaner layout
  labs(
    title = "Seasonal Total Rainfall Trends (2000–2021)",
    x = "Year",
    y = "Total Rainfall (mm)",
    fill = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey90"),
    plot.title = element_text(hjust = 0.5, size = 16),
    axis.text.x = element_text(hjust = 1) # Rotate x-axis labels for readability
  )

# Plot seasonal average humidity trends
ggplot(seasonal_summary, aes(x = Year, y = Avg_Humidity, color = Season, group = Season)) +
  geom_line(size = 1.2) +
  geom_point(size = 2) +
  labs(
    title = "Seasonal Average Humidity Trends (2000–2021)",
    x = "Year",
    y = "Average Humidity (%)",
    color = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey90"),
    plot.title = element_text(hjust = 0.5, size = 16)
  )

# Plot temperature and humidity trends on the same plot for comparison
ggplot(seasonal_summary, aes(x = Year, group = Season)) +
  geom_line(aes(y = Avg_Avg_Temp, color = Season), size = 1.2) +
  geom_line(aes(y = Avg_Humidity, linetype = Season), size = 1.2, color = "grey50") +
  labs(
    title = "Seasonal Temperature and Humidity Trends (2000–2021)",
    x = "Year",
    y = "Value",
    color = "Season",
    linetype = "Humidity"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey90"),
    plot.title = element_text(hjust = 0.5, size = 16)
  )

# Plot seasonal total rainfall and average humidity trends
ggplot(seasonal_summary, aes(x = factor(Year), group = Season)) +
  geom_bar(aes(y = Total_Rainfall, fill = Season), stat = "identity", position = position_dodge(width = 0.8), alpha = 0.8) +
  geom_line(aes(y = Avg_Humidity * 10, color = Season, group = Season), size = 1.2, linetype = "dashed") +
  scale_y_continuous(
    name = "Total Rainfall (mm)",
    sec.axis = sec_axis(~./10, name = "Average Humidity (%)")
  ) +
  labs(
    title = "Seasonal Total Rainfall and Average Humidity Trends (2000–2021)",
    x = "Year",
    fill = "Season",
    color = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey90"),
    plot.title = element_text(hjust = 0.5, size = 16),
    axis.text.x = element_text(hjust = 1)
  )


# Descriptive statistics by season, including humidity
descriptive_stats <- imputed_data %>%
  group_by(Season) %>%
  summarise(
    Avg_Max_Temp = mean(Max_Temp, na.rm = TRUE),
    Avg_Min_Temp = mean(Min_Temp, na.rm = TRUE),
    Avg_Avg_Temp = mean(Avg_Temp, na.rm = TRUE),
    Total_Rainfall = sum(Rainfall_mm, na.rm = TRUE)/4,
    Mean_Daily_Rainfall = mean(Rainfall_mm, na.rm = TRUE),
    Avg_Humidity = mean(Daily_RH, na.rm = TRUE),  # Average Humidity
    .groups = "drop"
  )

# View the updated descriptive stats
print(descriptive_stats)


# DETECT ANOMALIES IN SEASONS

library(dplyr)
library(ggplot2)
library(car)

# Group data by Year and Season, calculate mean and standard deviation for each year (including Humidity)
seasonal_variability <- imputed_data %>%
  group_by(Year, Season) %>%
  summarise(
    Avg_Avg_Temp = mean(Avg_Temp, na.rm = TRUE),
    Std_Avg_Temp = sd(Avg_Temp, na.rm = TRUE),
    Total_Rainfall = sum(Rainfall_mm, na.rm = TRUE)/4,
    Rainfall_SD = sd(Rainfall_mm, na.rm = TRUE),
    Avg_Humidity = mean(Daily_RH, na.rm = TRUE),
    Humidity_SD = sd(Daily_RH, na.rm = TRUE),
    .groups = "drop"
  )


library(dplyr)
library(car)
library(broom)
library(ggplot2)

# Aggregate by station, year, and season
overall_data_season <- imputed_data %>%
  mutate(
    Winter_Year = if_else(
      Season == "Winter" & month(Date) >= 12, year(Date) + 1, year(Date)
    )
  ) %>%
  group_by(Date, Season, Winter_Year) %>%
  summarise(
    Avg_Avg_Temp = mean(Avg_Temp, na.rm = TRUE),
    Total_Rainfall = sum(Rainfall_mm, na.rm = TRUE)/4,
    Avg_Humidity = mean(Daily_RH, na.rm = TRUE),
    .groups = "drop"
  )


# Plot seasonal temperature trends
ggplot(overall_data_season, aes(x = Date, y = Avg_Avg_Temp, color = Season)) +
  geom_line(
    aes(group = if_else(Season == "Winter", interaction(Season, Winter_Year), interaction(Season, year(Date)))),
    alpha = 0.6
  ) +
  geom_smooth(method = "loess", se = FALSE, linetype = "dashed", size = 1) +
  facet_wrap(~Season, scales = "free_y") +
  labs(
    title = "Daily Temperature Trends by Season (Adjusted Winter)",
    x = "Date",
    y = "Average Temperature (°C)",
    color = "Season"
  ) +
  theme_minimal()


# Plot seasonal humidity trends with adjusted winter grouping
ggplot(overall_data_season, aes(x = Date, y = Avg_Humidity, color = Season)) +
  geom_line(
    aes(group = if_else(Season == "Winter", interaction(Season, Winter_Year), interaction(Season, year(Date)))),
    alpha = 0.6
  ) +
  geom_smooth(method = "loess", se = FALSE, linetype = "dashed", size = 1) +
  facet_wrap(~Season, scales = "free_y") +
  labs(
    title = "Daily Humidity Trends by Season",
    x = "Date",
    y = "Average Humidity (%)",
    color = "Season"
  ) +
  theme_minimal()

# Plot seasonal rainfall trends with adjusted winter grouping
ggplot(overall_data_season, aes(x = Date, y = Total_Rainfall, color = Season)) +
  geom_line(
    aes(group = if_else(Season == "Winter", interaction(Season, Winter_Year), interaction(Season, year(Date)))),
    alpha = 0.6
  ) +
  
  facet_wrap(~Season, scales = "free_y") +
  labs(
    title = "Daily Rainfall Trends by Season",
    x = "Date",
    y = "Total Rainfall (mm)",
    color = "Season"
  ) +
  theme_minimal()



library(lubridate)

# Aggregate by month
monthly_data <- overall_data_season %>%
  mutate(Month = floor_date(Date, "month")) %>%
  group_by(Month, Season) %>%
  summarise(
    Avg_Monthly_Temp = mean(Avg_Avg_Temp, na.rm = TRUE),
    Total_Monthly_Rainfall = sum(Total_Rainfall, na.rm = TRUE),
    Avg_Monthly_Humidity = mean(Avg_Humidity, na.rm = TRUE),
    .groups = "drop"
  )


# Define a common color palette for the seasons
season_colors <- c("Winter" = "#1f77b4", "Spring" = "#2ca02c", "Summer" = "#ff7f0e", "Autumn" = "#d62728")

# Monthly temperature trends with matching colors
ggplot(monthly_data, aes(x = Month, y = Avg_Monthly_Temp, fill = Season)) +
  geom_bar(stat = "identity", position = "dodge", alpha = 0.8) +
  geom_smooth(
    aes(group = Season, color = Season), 
    method = "loess", se = FALSE, linetype = "dashed", size = 1
  ) +
  scale_fill_manual(values = season_colors) + # Apply colors to bars
  scale_color_manual(values = season_colors) + # Apply matching colors to trend lines
  labs(
    title = "Monthly Average Temperature by Season",
    x = "Month",
    y = "Average Temperature (°C)",
    fill = "Season",
    color = "Season"
  ) +
  scale_x_date(date_breaks = "3 months", date_labels = "%b %Y") +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )

# Monthly rainfall trends with matching colors
ggplot(monthly_data, aes(x = Month, y = Total_Monthly_Rainfall, fill = Season)) +
  geom_bar(stat = "identity", position = "dodge", alpha = 0.8) +
  geom_smooth(
    aes(group = Season, color = Season), 
    method = "loess", se = FALSE, linetype = "dashed", size = 1
  ) +
  scale_fill_manual(values = season_colors) + # Apply colors to bars
  scale_color_manual(values = season_colors) + # Apply matching colors to trend lines
  labs(
    title = "Monthly Total Rainfall by Season",
    x = "Month",
    y = "Total Rainfall (mm)",
    fill = "Season",
    color = "Season"
  ) +
  scale_x_date(date_breaks = "3 months", date_labels = "%b %Y") +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )

ggplot(monthly_data, aes(x = Month, y = Avg_Monthly_Humidity, fill = Season)) +
  geom_bar(stat = "identity", position = "dodge", alpha = 0.8) +
  geom_smooth(
    aes(group = Season, color = Season), 
    method = "loess", se = FALSE, linetype = "dashed", size = 1
  ) +
  scale_fill_manual(values = season_colors) + 
  scale_color_manual(values = season_colors) +
  labs(
    title = "Monthly Average Humidity by Season",
    x = "Month",
    y = "Average Humidity (%)",
    fill = "Season",
    color = "Season"
  ) +
  scale_x_date(date_breaks = "3 months", date_labels = "%b %Y") +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )


ggplot(monthly_data, aes(x = Month)) +
  geom_line(aes(y = Avg_Monthly_Temp, color = "Temperature"), size = 1.2) +
  geom_line(aes(y = Avg_Monthly_Humidity, color = "Humidity"), size = 1.2, linetype = "dashed") +
  geom_bar(aes(y = Total_Monthly_Rainfall / 10, fill = "Rainfall"), stat = "identity", alpha = 0.5) +
  scale_y_continuous(
    name = "Temperature (°C) & Humidity (%)",
    sec.axis = sec_axis(~.*10, name = "Rainfall (mm)")
  ) +
  labs(
    title = "Monthly Temperature, Humidity, and Rainfall Trends",
    x = "Month",
    color = "Metric",
    fill = "Metric"
  ) +
  theme_minimal() +
  theme(
    legend.position = "top",
    axis.text.x = element_text(angle = 45, hjust = 1)
  )



# Seasonal temperature and humidity summary
season_summary <- overall_data_season %>%
  group_by(Season, Year = year(Date)) %>%
  summarise(
    Mean_Temp = mean(Avg_Avg_Temp, na.rm = TRUE),
    SE_Temp = sd(Avg_Avg_Temp, na.rm = TRUE) / sqrt(n()),
    Mean_Humidity = mean(Avg_Humidity, na.rm = TRUE),
    SE_Humidity = sd(Avg_Humidity, na.rm = TRUE) / sqrt(n()),
    .groups = "drop"
  )

# Dot and whisker plot for temperature and humidity
ggplot(season_summary, aes(x = factor(Year))) +
  geom_point(aes(y = Mean_Temp, color = Season), position = position_dodge(0.5), size = 2) +
  geom_errorbar(
    aes(ymin = Mean_Temp - SE_Temp, ymax = Mean_Temp + SE_Temp, color = Season),
    position = position_dodge(0.5),
    width = 0.2
  ) +
  geom_line(aes(y = Mean_Humidity, group = Season), color = "grey50", linetype = "dashed") +
  facet_wrap(~Season, scales = "free") +
  labs(
    title = "Seasonal Average Temperature and Humidity by Year",
    x = "Year",
    y = "Mean Temperature (°C) & Humidity (%)",
    color = "Season"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )



# Dot and whisker plot
ggplot(season_summary, aes(x = factor(Year), y = Mean_Temp, color = Season)) +
  geom_point(position = position_dodge(0.5), size = 2) +
  geom_errorbar(
    aes(ymin = Mean_Temp - SE_Temp, ymax = Mean_Temp + SE_Temp),
    position = position_dodge(0.5),
    width = 0.2
  ) +
  facet_wrap(~Season, scales = "free") +
  scale_color_manual(values = season_colors) + # Use seasonal colors
  labs(
    title = "Average Temperature by Year (with Standard Errors)",
    x = "Year",
    y = "Mean Temperature (°C)",
    color = "Season"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )


# Fit LOESS models for daily data
library(ggplot2)

# Temperature trend
ggplot(overall_data_season, aes(x = Date, y = Avg_Avg_Temp)) +
  geom_point(alpha = 0.3) +
  geom_smooth(method = "loess", se = FALSE, color = "blue") +
  labs(
    title = "Non-Linear Daily Temperature Trends",
    x = "Date",
    y = "Average Temperature (°C)"
  ) +
  theme_minimal()



library(dplyr)
# Aggregate data across all stations for overall trends
overall_data <- imputed_data %>%
  group_by(Date) %>%
  summarise(
    Max_Temp = mean(Max_Temp, na.rm = TRUE),
    Min_Temp = mean(Min_Temp, na.rm = TRUE),
    Avg_Temp = mean(Avg_Temp, na.rm = TRUE)
  )

ggplot(overall_data, aes(x = Date)) +
  # Dots for Max, Min, and Avg Temperatures with increased transparency
  geom_point(aes(y = Max_Temp, color = "Max Temp"), size = 1.5, alpha = 0.3) +
  geom_point(aes(y = Min_Temp, color = "Min Temp"), size = 1.5, alpha = 0.3) +
  geom_point(aes(y = Avg_Temp, color = "Avg Temp"), size = 1.5, alpha = 0.3) +
  # LOESS trend lines with distinct line types and colors matching dots
  geom_smooth(aes(y = Max_Temp, linetype = "Max Temp"), method = "loess", se = FALSE, size = 1.5, color = "black") +
  geom_smooth(aes(y = Min_Temp, linetype = "Min Temp"), method = "loess", se = FALSE, size = 1.5, color = "black") +
  geom_smooth(aes(y = Avg_Temp, linetype = "Avg Temp"), method = "loess", se = FALSE, size = 1.5, color = "black") +
  # Custom color and linetype mappings
  scale_color_manual(values = c("Max Temp" = "red", "Min Temp" = "blue", "Avg Temp" = "green")) +
  scale_linetype_manual(values = c("Max Temp" = "dotted", "Min Temp" = "dotted", "Avg Temp" = "solid")) +
  # Customize legend appearance
  guides(
    color = guide_legend(title = "Temperature Type", override.aes = list(size = 3, alpha = 1)),
    linetype = guide_legend(title = "Line Type", override.aes = list(size = 1.5))
  ) +
  # Titles and labels
  labs(
    title = "Historical Temperature Trends (2000–2021)",
    x = "Year",
    y = "Temperature (°C)"
  ) +
  # Minimal theme
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top",
    plot.title = element_text(hjust = 0.5, size = 16)
  )



# Aggregate daily rainfall and count rainy days by month
monthly_rainfall <- imputed_data %>%
  mutate(Month = floor_date(Date, "month")) %>%
  group_by(Month) %>%
  summarise(
    Total_Rainfall = sum(Rainfall_mm, na.rm = TRUE)/4,
    Rainy_Days = sum(Rainfall_mm > 0, na.rm = TRUE)/4
  )


# Plot total rainfall and number of rainy days over time
ggplot(monthly_rainfall, aes(x = Month)) +
  geom_line(aes(y = Total_Rainfall, color = "Total Rainfall (mm)"), size = 1) +
  geom_line(aes(y = Rainy_Days, color = "Number of Rainy Days"), size = 1) +
  scale_x_date(date_breaks = "1 year", date_labels = "%Y") + # Yearly labels
  scale_color_manual(values = c("Total Rainfall (mm)" = "blue", "Number of Rainy Days" = "green")) +
  labs(
    title = "Monthly Rainfall and Rainy Days Trends (2000–2021)",
    x = "Year",
    y = "Value",
    color = "Metric"
  ) +
  theme_minimal()


#Check for trends

# Fit linear and non-linear models for temperature
temp_trends <- lm(Avg_Temp ~ Date, data = daily_combined_data)
summary(temp_trends)

# GAM for non-linear trends
library(mgcv)
gam_temp <- gam(Avg_Temp ~ s(as.numeric(Date)), data = daily_combined_data)
summary(gam_temp)

# Visualize the trends
ggplot(daily_combined_data, aes(x = Date, y = Avg_Temp)) +
  geom_point(alpha = 0.4) +
  geom_smooth(method = "lm", se = TRUE, color = "red", size = 1) +
  geom_smooth(method = "gam", formula = y ~ s(x), se = TRUE, color = "blue", size = 1) +
  labs(title = "Temperature Trends (Linear and Non-linear)", y = "Average Temperature (°C)", x = "Date") +
  theme_minimal()


library(forecast)
library(ggplot2)
library(tseries)

# Prepare data for SARIMA
avg_temp_ts <- ts(daily_combined_data$Avg_Temp, frequency = 365, start = c(2000, 1))

# Step 1: Check Stationarity
adf_test <- adf.test(avg_temp_ts, alternative = "stationary")
print(adf_test)

# Apply differencing if necessary
if (adf_test$p.value > 0.05) {
  avg_temp_diff <- diff(avg_temp_ts)
  print("Data differenced to achieve stationarity.")
} else {
  avg_temp_diff <- avg_temp_ts
  print("Data is already stationary.")
}

# Step 2: Identify SARIMA parameters using auto.arima
sarima_model <- auto.arima(avg_temp_diff, seasonal = TRUE)

# Display SARIMA model summary
summary(sarima_model)

# Step 3: Forecast future values (e.g., next 365 days)
forecast_temp <- forecast(sarima_model, h = 365)

# Plot the forecast
autoplot(forecast_temp) +
  labs(
    title = "SARIMA Forecast for Average Temperature",
    x = "Time",
    y = "Average Temperature (°C)"
  ) +
  theme_minimal()

# Step 4: Diagnostics
checkresiduals(sarima_model)


# Step 3: Forecast future values (5 years = 1825 days)
forecast_temp_long <- forecast(sarima_model, h = 1825)

# Plot the 5-year forecast
autoplot(forecast_temp_long) +
  labs(
    title = "5-Year SARIMA Forecast for Average Temperature",
    x = "Time",
    y = "Average Temperature (°C)"
  ) +
  theme_minimal()



# SARIMA MODEL WITH VALIDATION SET

library(dplyr)

# Determine the split point
n <- nrow(daily_combined_data)
split_point <- round(0.9 * n)

# Split the data
train_data <- daily_combined_data[1:split_point, ]  # First 80% of the data
test_data <- daily_combined_data[(split_point + 1):n, ]  # Remaining 20% of the data

# Inspect the splits
print(paste("Training set length:", nrow(train_data)))
print(paste("Test set length:", nrow(test_data)))

library(forecast)

# Convert the training data into a time series object
train_ts <- ts(train_data$Avg_Temp, frequency = 365, start = c(2000, 1))  # Assuming data starts in 2000

# Fit SARIMA model to training data
sarima_model_train <- Arima(train_ts, order = c(5, 0, 0), seasonal = c(0, 1, 0))

# Summary of the model
summary(sarima_model_train)

# Forecast for the length of the test data
forecast_horizon <- nrow(test_data)
sarima_forecast_train <- forecast(sarima_model_train, h = forecast_horizon)

# Inspect the forecast
autoplot(sarima_forecast_train) +
  autolayer(test_data$Avg_Temp, series = "Actual", color = "blue") +
  labs(
    title = "SARIMA Model Forecast vs. Actual (Validation Set)",
    x = "Date",
    y = "Average Temperature"
  ) +
  theme_minimal()

library(forecast)
library(lubridate)

# Convert test_data to a time-series object
test_data_ts <- ts(test_data$Avg_Temp, start = c(year(test_data$Date[1]), month(test_data$Date[1])), frequency = 365)

# Generate forecast for the test data length
sarima_forecast_train <- forecast(sarima_model_train, h = forecast_horizon)

# Plot forecast vs actual data
autoplot(sarima_forecast_train) +
  autolayer(test_data_ts, series = "Actual", color = "blue") +
  labs(
    title = "SARIMA Model Forecast vs. Actual (Validation Set)",
    x = "Date",
    y = "Average Temperature"
  ) +
  theme_minimal()


# Combine test data and forecast into a single data frame for better comparison
forecast_df <- data.frame(
  Date = test_data$Date,
  Actual = test_data$Avg_Temp,
  Forecast = as.numeric(sarima_forecast_train$mean)
)

# Plot Actual vs Forecast
ggplot(forecast_df, aes(x = Date)) +
  geom_line(aes(y = Actual, color = "Actual"), size = 1) +
  geom_line(aes(y = Forecast, color = "Forecast"), size = 1, linetype = "dashed") +
  labs(
    title = "SARIMA Model: Forecast vs Actual (Validation Set)",
    x = "Date",
    y = "Average Temperature (°C)",
    color = "Legend"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    plot.title = element_text(hjust = 0.5, size = 16),
    axis.text.x = element_text(angle = 45, hjust = 1)
  )

# Calculate residuals: difference between actual and forecasted values
forecast_df <- forecast_df %>%
  mutate(Residuals = Actual - Forecast)

# Compute accuracy metrics
model_diagnostics <- forecast_df %>%
  summarise(
    ME = mean(Residuals),  # Mean Error
    RMSE = sqrt(mean(Residuals^2)),  # Root Mean Squared Error
    MAE = mean(abs(Residuals)),  # Mean Absolute Error
    MAPE = mean(abs(Residuals / Actual)) * 100,  # Mean Absolute Percentage Error
    R2 = 1 - (sum(Residuals^2) / sum((Actual - mean(Actual))^2))  # Coefficient of determination
  )

# Display diagnostics
print(model_diagnostics)

# Plot residuals for diagnostics
ggplot(forecast_df, aes(x = Date, y = Residuals)) +
  geom_line(color = "darkred", size = 0.7) +
  geom_hline(yintercept = 0, linetype = "dashed", color = "black") +
  labs(
    title = "Residuals Over Time",
    x = "Date",
    y = "Residuals (Actual - Forecast)"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5, size = 16),
    axis.text.x = element_text(angle = 45, hjust = 1)
  )

# Residual histogram and density plot
ggplot(forecast_df, aes(x = Residuals)) +
  geom_histogram(aes(y = ..density..), bins = 30, fill = "lightblue", color = "black", alpha = 0.7) +
  geom_density(color = "darkblue", size = 1) +
  labs(
    title = "Residuals Distribution",
    x = "Residuals",
    y = "Density"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5, size = 16)
  )

library(tseries)

# Check stationarity of residuals
adf_test_residuals <- adf.test(residuals(sarima_model))
print(adf_test_residuals)

# Step 4: Diagnostics
checkresiduals(sarima_forecast_train)





# Fourier terms due to a 365 lag autocorrelation


library(forecast)

# Add Fourier terms (using 2 harmonics for annual seasonality)
fourier_terms <- fourier(avg_temp_diff, K = 2)

# Fit an ARIMA model with Fourier terms
sarima_fourier_model <- auto.arima(
  avg_temp_diff,
  xreg = fourier_terms,
  seasonal = FALSE  # No explicit seasonal components
)

# Summary and diagnostics
summary(sarima_fourier_model)
checkresiduals(sarima_fourier_model)

# Forecast using Fourier terms for 5 years (1825 days)
future_fourier <- fourier(avg_temp_diff, K = 2, h = 1825)  # 5 years of daily forecasts
sarima_fourier_forecast <- forecast(sarima_fourier_model, xreg = future_fourier)

# Plot the 5-year Fourier-based SARIMA forecast
autoplot(sarima_fourier_forecast) +
  labs(
    title = "5-Year SARIMA Model Forecast with Fourier Terms",
    x = "Time",
    y = "Temperature Differences"
  ) +
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




# Set working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/aminsakha")


# Prepare tables to export
tables_to_export <- list(
  "Filled Values Summary" = filled_values_summary,
  "Final Data" = imputed_data,
  "Seasonal Summary" = seasonal_summary,
  "Descriptive Stats" = descriptive_stats,
  "Gaps Report" = gaps_report,
  "Model Diagnostics (Fourier)" = model_diagnostics,
  "SARIMA Residual Diagnostics" = data.frame(Residuals = residuals(sarima_model)),
  "Seasonal Variability" = seasonal_variability
)

# Export tables to Excel
save_apa_formatted_excel(tables_to_export, "Summary_Tables.xlsx")


# Prepare tables to export
tables_to_export <- list(
  "Filled Values Summary" = filled_values_summary,
  "Final Data" = imputed_data,
  "Seasonal Summary" = seasonal_summary,
  "Descriptive Stats" = descriptive_stats,
  
  "Seasonal Variability" = seasonal_variability
)

# Export tables to Excel
save_apa_formatted_excel(tables_to_export, "Summary_Tables_v3_adjRain.xlsx")


# Save SARIMA residual diagnostics plot
sarima_residual_diagnostics <- checkresiduals(sarima_model, plot = TRUE)  # Generate residual diagnostics
ggsave("SARIMA_Residual_Diagnostics.png", width = 10, height = 6, dpi = 300)

# Save Fourier Model diagnostic residuals
fourier_residual_diagnostics <- checkresiduals(sarima_fourier_model, plot = TRUE)  # Generate residual diagnostics
ggsave("Fourier_Model_Residual_Diagnostics.png", width = 10, height = 6, dpi = 300)

# Ensure test_data$Avg_Temp is a time series
test_data_ts <- ts(
  test_data$Avg_Temp,
  start = c(year(test_data$Date[1]), yday(test_data$Date[1])),
  frequency = 365
)

# Create the plot
sarima_forecast_vs_actual_plot <- autoplot(sarima_forecast_train) +
  autolayer(test_data_ts, series = "Actual") +
  labs(
    title = "SARIMA Model Forecast vs Actual (Validation Set)",
    x = "Date",
    y = "Average Temperature"
  ) +
  theme_minimal()

# Save the plot
ggsave("SARIMA_Forecast_vs_Actual.png", sarima_forecast_vs_actual_plot, width = 10, height = 6, dpi = 300)

# Save forecast vs actual comparison plot (Fourier Model)
fourier_forecast_vs_actual_plot <- ggplot(forecast_df, aes(x = Date)) +
  geom_line(aes(y = Actual, color = "Actual"), size = 1) +
  geom_line(aes(y = Forecast, color = "Forecast"), size = 1, linetype = "dashed") +
  labs(
    title = "Fourier Model: Forecast vs Actual (Validation Set)",
    x = "Date",
    y = "Average Temperature (°C)",
    color = "Legend"
  ) +
  theme_minimal()
ggsave("Fourier_Model_Forecast_vs_Actual.png", fourier_forecast_vs_actual_plot, width = 10, height = 6, dpi = 300)

# Save residual time-series plot
residual_time_series_plot <- ggplot(forecast_df, aes(x = Date, y = Residuals)) +
  geom_line(color = "darkred", size = 0.7) +
  geom_hline(yintercept = 0, linetype = "dashed", color = "black") +
  labs(
    title = "Residuals Over Time",
    x = "Date",
    y = "Residuals (Actual - Forecast)"
  ) +
  theme_minimal()
ggsave("Residual_Time_Series.png", residual_time_series_plot, width = 10, height = 6, dpi = 300)

# Save residual distribution histogram
residual_distribution_plot <- ggplot(forecast_df, aes(x = Residuals)) +
  geom_histogram(aes(y = ..density..), bins = 30, fill = "lightblue", color = "black", alpha = 0.7) +
  geom_density(color = "darkblue", size = 1) +
  labs(
    title = "Residuals Distribution",
    x = "Residuals",
    y = "Density"
  ) +
  theme_minimal()
ggsave("Residual_Distribution.png", residual_distribution_plot, width = 10, height = 6, dpi = 300)


library(ggplot2)
library(dplyr)

# Plot with shaded area, thinner lines for Min and Max, and data labels for Avg
seasonal_temp_trends <- ggplot(seasonal_summary, aes(x = Year, group = Season, fill = Season)) +
  geom_ribbon(aes(ymin = Avg_Min_Temp, ymax = Avg_Max_Temp), alpha = 0.2) +
  geom_line(aes(y = Avg_Avg_Temp, color = Season), size = 1.2) +
  geom_line(aes(y = Avg_Min_Temp, color = Season), size = 0.5, linetype = "dashed") +
  geom_line(aes(y = Avg_Max_Temp, color = Season), size = 0.5, linetype = "dashed") +
  geom_text(aes(y = Avg_Avg_Temp, label = round(Avg_Avg_Temp, 1)), color = "black", size = 3, vjust = -1, check_overlap = TRUE) +
  scale_x_continuous(breaks = seq(2000, 2024, by = 1)) +
  labs(
    title = "Seasonal Temperature Trends (2000–2021)",
    x = "Year",
    y = "Temperature (°C)",
    color = "Season",
    fill = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey90"),
    plot.title = element_text(hjust = 0.5, size = 16)
  )

# Plot seasonal total rainfall trends
seasonal_rainfall_trends <- ggplot(seasonal_summary, aes(x = factor(Year), y = Total_Rainfall, fill = Season)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), alpha = 0.8) +
  geom_text(
    aes(y = Total_Rainfall, label = round(Total_Rainfall, 1)), 
    position = position_dodge(width = 0.8), 
    vjust = -0.5, color = "black", size = 3
  ) +
  scale_x_discrete(breaks = seq(2000, 2024, by = 1)) +
  labs(
    title = "Seasonal Total Rainfall Trends (2000–2022)",
    x = "Year",
    y = "Total Rainfall (mm)",
    fill = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    legend.position = "top",
    panel.grid.minor = element_blank(),
    panel.grid.major = element_line(color = "grey90"),
    plot.title = element_text(hjust = 0.5, size = 16),
    axis.text.x = element_text(hjust = 1)
  )

# Plot seasonal temperature trends
adjusted_winter_temp_trends <- ggplot(overall_data_season, aes(x = Date, y = Avg_Avg_Temp, color = Season)) +
  geom_line(
    aes(group = if_else(Season == "Winter", interaction(Season, Winter_Year), interaction(Season, year(Date)))),
    alpha = 0.6
  ) +
  geom_smooth(method = "loess", se = FALSE, linetype = "dashed", size = 1) +
  facet_wrap(~Season, scales = "free_y") +
  labs(
    title = "Daily Temperature Trends by Season (Adjusted Winter)",
    x = "Date",
    y = "Average Temperature (°C)",
    color = "Season"
  ) +
  theme_minimal()

# Monthly temperature trends with matching colors
monthly_temp_trends <- ggplot(monthly_data, aes(x = Month, y = Avg_Monthly_Temp, fill = Season)) +
  geom_bar(stat = "identity", position = "dodge", alpha = 0.8) +
  geom_smooth(
    aes(group = Season, color = Season), 
    method = "loess", se = FALSE, linetype = "dashed", size = 1
  ) +
  scale_fill_manual(values = season_colors) +
  scale_color_manual(values = season_colors) +
  labs(
    title = "Monthly Average Temperature by Season",
    x = "Month",
    y = "Average Temperature (°C)",
    fill = "Season",
    color = "Season"
  ) +
  scale_x_date(date_breaks = "3 months", date_labels = "%b %Y") +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )

# Dot and whisker plot
seasonal_avg_dot_whisker <- ggplot(season_summary, aes(x = factor(Year), y = Mean_Temp, color = Season)) +
  geom_point(position = position_dodge(0.5), size = 2) +
  geom_errorbar(
    aes(ymin = Mean_Temp - SE_Temp, ymax = Mean_Temp + SE_Temp),
    position = position_dodge(0.5),
    width = 0.2
  ) +
  facet_wrap(~Season, scales = "free") +
  scale_color_manual(values = season_colors) +
  labs(
    title = "Average Temperature by Year (with Standard Errors)",
    x = "Year",
    y = "Mean Temperature (°C)",
    color = "Season"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1),
    legend.position = "top"
  )



# Save plots
ggsave("Seasonal_Temperature_Trends.png", seasonal_temp_trends, width = 10, height = 6, dpi = 300)
ggsave("Seasonal_Rainfall_Trends.png", seasonal_rainfall_trends, width = 10, height = 6, dpi = 300)
ggsave("Adjusted_Winter_Temperature_Trends.png", adjusted_winter_temp_trends, width = 10, height = 6, dpi = 300)
ggsave("Monthly_Temperature_Trends.png", monthly_temp_trends, width = 10, height = 6, dpi = 300)
ggsave("Seasonal_Averages_Dot_Whisker.png", seasonal_avg_dot_whisker, width = 10, height = 6, dpi = 300)


# Prepare tables to export
tables_to_export <- list(
  "Daily Data" = daily_combined_data,
  "Forecasted Temperatures" = forecast_df
)


# Export tables to Excel
save_apa_formatted_excel(tables_to_export, "Additional_Tables.xlsx")

# New requests
# Define traditional seasons based on dates
overall_data <- overall_data %>%
  mutate(
    Year = year(Date),
    Traditional_Season = case_when(
      Date >= as.Date(paste(Year, "12-22", sep = "-")) | Date < as.Date(paste(Year, "03-21", sep = "-")) ~ "Winter",
      Date >= as.Date(paste(Year, "03-21", sep = "-")) & Date < as.Date(paste(Year, "06-22", sep = "-")) ~ "Spring",
      Date >= as.Date(paste(Year, "06-22", sep = "-")) & Date < as.Date(paste(Year, "09-23", sep = "-")) ~ "Summer",
      Date >= as.Date(paste(Year, "09-23", sep = "-")) & Date < as.Date(paste(Year, "12-22", sep = "-")) ~ "Autumn"
    )
  )

# Calculate thresholds for min, max, and avg temperatures
seasonal_thresholds <- overall_data %>%
  group_by(Traditional_Season) %>%
  summarise(
    Min_Temp_Threshold = quantile(Min_Temp, probs = 0.1, na.rm = TRUE),  # 10th percentile
    Max_Temp_Threshold = quantile(Max_Temp, probs = 0.9, na.rm = TRUE),  # 90th percentile
    Avg_Temp_Lower = quantile(Avg_Temp, probs = 0.1, na.rm = TRUE),      # Lower bound for avg temp
    Avg_Temp_Upper = quantile(Avg_Temp, probs = 0.9, na.rm = TRUE),      # Upper bound for avg temp
    .groups = "drop"
  )

# Display the thresholds
print(seasonal_thresholds)


# Standardize Traditional_Season column in both datasets
overall_data <- overall_data %>%
  mutate(Traditional_Season = as.character(Traditional_Season))  # Ensure character type

seasonal_thresholds <- seasonal_thresholds %>%
  mutate(Traditional_Season = as.character(Traditional_Season))  # Ensure character type


# Join thresholds and classify observed seasons
overall_data <- overall_data %>%
  mutate(
    Observed_Season = case_when(
      is.na(Avg_Temp) ~ NA_character_,  # Explicitly assign NA when Avg_Temp is missing
      Avg_Temp <= Min_Temp_Threshold ~ "Winter",
      Avg_Temp >= Max_Temp_Threshold ~ "Summer",
      Avg_Temp > Min_Temp_Threshold & Avg_Temp < Max_Temp_Threshold & Traditional_Season == "Spring" ~ "Spring",
      Avg_Temp > Min_Temp_Threshold & Avg_Temp < Max_Temp_Threshold & Traditional_Season == "Autumn" ~ "Autumn",
      TRUE ~ Traditional_Season
    )
  )



# Generate detailed table by year and season
season_shift_table <- overall_data %>%
  group_by(Year, Traditional_Season, Observed_Season) %>%
  summarise(
    Days_Shifted = n(),  # Count of days
    .groups = "drop"
  ) %>%
  group_by(Year, Traditional_Season) %>%
  mutate(
    Total_Days = sum(Days_Shifted),  # Total days in the traditional season
    Percentage_Shifted = 100 * Days_Shifted / Total_Days  # Percentage of days shifted
  ) %>%
  ungroup() %>%
  complete(
    Year = unique(overall_data$Year),
    Traditional_Season = unique(overall_data$Traditional_Season),
    Observed_Season = unique(overall_data$Observed_Season),
    fill = list(Days_Shifted = 0, Percentage_Shifted = 0)  # Fill missing combinations with 0
  ) %>%
  filter(Traditional_Season != Observed_Season)  # Exclude rows where no shifts occurred

# Calculate total days for each traditional season across the dataset
total_days_per_season <- overall_data %>%
  group_by(Traditional_Season) %>%
  summarise(
    Total_Days = n(),
    .groups = "drop"
  )

# Generate overall table (aggregated across all years)
overall_shift_table <- overall_data %>%
  group_by(Traditional_Season, Observed_Season) %>%
  summarise(
    Days_Shifted = n(),  # Total days shifted across all years
    .groups = "drop"
  ) %>%
  left_join(total_days_per_season, by = "Traditional_Season") %>%
  mutate(
    Percentage_Shifted = 100 * Days_Shifted / Total_Days  # Percentage relative to all days for the season
  ) %>%
  complete(
    Traditional_Season = unique(overall_data$Traditional_Season),
    Observed_Season = unique(overall_data$Observed_Season),
    fill = list(Days_Shifted = 0, Percentage_Shifted = NA)  # Include missing combinations
  ) %>%
  filter(Traditional_Season != Observed_Season)  # Exclude rows where no shifts occurred

# View tables
print(season_shift_table)
print(overall_shift_table)


#New table to export

# Prepare tables to export
tables_to_export <- list(
 "Data" = overall_data,
  "Seasonal Summary" = seasonal_summary,
  "Descriptive Stats" = descriptive_stats,
 
  "Seasonal Variability" = seasonal_variability,
  "Overall Data by Season" = overall_data_season,
  "Seasonal Summary by Year" = season_summary
)

# Export tables to Excel
save_apa_formatted_excel(tables_to_export, "Summary_Tables_v22.xlsx")


# Prepare tables to export
tables_to_export <- list(
  
  "Seasonal Thresholds" = seasonal_thresholds,
  "Seasonal Shifts" = season_shift_table,
  
  "Overall Shifts" = overall_shift_table
)

# Export tables to Excel
save_apa_formatted_excel(tables_to_export, "Seasonal_Shifts.xlsx")


# Calculate total days for each traditional season across the dataset
total_days_per_season <- overall_data %>%
  group_by(Traditional_Season) %>%
  summarise(
    Total_Days = n(),
    .groups = "drop"
  )

# Calculate the total number of years in the dataset
num_years <- overall_data %>%
  summarise(Years = n_distinct(Year)) %>%
  pull(Years)

# Generate overall table (aggregated across all years) with average annual shifted days
overall_shift_table <- overall_data %>%
  group_by(Traditional_Season, Observed_Season) %>%
  summarise(
    Days_Shifted = n(),  # Total days shifted across all years
    .groups = "drop"
  ) %>%
  left_join(total_days_per_season, by = "Traditional_Season") %>%
  mutate(
    Percentage_Shifted = 100 * Days_Shifted / Total_Days,  # Percentage relative to all days for the season
    Avg_Shifted_Days_Per_Year = Days_Shifted / num_years  # Average shifted days per year
  ) %>%
  complete(
    Traditional_Season = unique(overall_data$Traditional_Season),
    Observed_Season = unique(overall_data$Observed_Season),
    fill = list(Days_Shifted = 0, Percentage_Shifted = NA, Avg_Shifted_Days_Per_Year = 0)  # Include missing combinations
  ) %>%
  filter(Traditional_Season != Observed_Season)  # Exclude rows where no shifts occurred

# View the updated overall_shift_table
print(overall_shift_table)



# Shift in rainfall season

library(dplyr)

# Define wet and dry season periods
wet_season_start <- "09-01"
wet_season_end <- "06-30"
dry_season_start <- "07-01"
dry_season_end <- "08-31"

# Add season type (wet or dry) based on the date
rain_season_analysis <- imputed_data %>%
  group_by(Date) %>%
  summarise(
    Avg_Rainfall_mm = mean(Rainfall_mm, na.rm = TRUE),  # Average rainfall across stations
    Rainy_Days = sum(Rainfall_mm > 0, na.rm = TRUE)     # Rainy days count across stations
  ) %>%
  mutate(
    Rain_Year = if_else(month(Date) >= 9, year(Date), year(Date) - 1), # Assign rain year
    Season_Type = case_when(
      between(month(Date), 9, 12) | between(month(Date), 1, 6) ~ "Wet Season",
      between(month(Date), 7, 8) ~ "Dry Season",
      TRUE ~ NA_character_
    )
  ) %>%
  group_by(Rain_Year, Season_Type) %>%
  summarise(
    Total_Rainfall = sum(Avg_Rainfall_mm, na.rm = TRUE), # Total rainfall for each season
    Total_Rainy_Days = sum(Rainy_Days, na.rm = TRUE),    # Total rainy days for each season
    .groups = "drop"
  ) %>%
  pivot_wider(
    names_from = Season_Type,
    values_from = c(Total_Rainfall, Total_Rainy_Days),
    names_sep = "_"
  )

# Replace NAs with zeros for missing values
rain_season_analysis[is.na(rain_season_analysis)] <- 0

# Display the final table
print(rain_season_analysis)












library(dplyr)
library(ggplot2)

# Step 1: Calculate average rainfall across stations for each day
average_rainfall <- imputed_data %>%
  group_by(Date) %>%
  summarise(
    Avg_Rainfall_mm = mean(Rainfall_mm, na.rm = TRUE), # Average rainfall across stations
    Rainy_Days = sum(Rainfall_mm > 0, na.rm = TRUE)   # Count rainy days across stations
  )

# Step 2: Define the rain year (July 1st to June 30th) and calculate cumulative rainfall
rainfall_analysis <- average_rainfall %>%
  mutate(
    Rain_Year = if_else(month(Date) >= 7, year(Date), year(Date) - 1), # Assign rain year based on July
    Rain_Year_Label = paste0(Rain_Year, "-", Rain_Year + 1), # Create rain year label (e.g., 1999-2000)
    Rain_Year_Start = as.Date(paste0(Rain_Year, "-07-01")), # Rain year starts July 1st
    Rain_Year_End = as.Date(paste0(Rain_Year + 1, "-06-30")) # Rain year ends June 30th
  ) %>%
  filter(Date >= Rain_Year_Start & Date <= Rain_Year_End) %>% # Filter for rain year dates only
  group_by(Rain_Year, Rain_Year_Label) %>%
  arrange(Date) %>%
  mutate(
    Cumulative_Rainfall = cumsum(Avg_Rainfall_mm), # Cumulative rainfall starting July 1st
    Total_Rainfall_Rain_Year = sum(Avg_Rainfall_mm, na.rm = TRUE) # Total rainfall for the rain year
  ) %>%
  summarise(
    Observed_Start_Date = min(Date[Cumulative_Rainfall / Total_Rainfall_Rain_Year >= 0.005], na.rm = TRUE), # Start at 5th percentile
    Observed_End_Date = max(Date[Cumulative_Rainfall / Total_Rainfall_Rain_Year <= 0.995], na.rm = TRUE), # End at 98th percentile
    Observed_Duration = as.numeric(Observed_End_Date - Observed_Start_Date),
    Total_Rainfall = Total_Rainfall_Rain_Year,
    Rainy_Days = sum(Avg_Rainfall_mm > 0, na.rm = TRUE), # Count rainy days
    .groups = "drop" # Ensure ungrouped result
  )

# Step 3: Add expected rainy season information and calculate deviations
rainfall_results <- rainfall_analysis %>%
  mutate(
    Expected_Start_Date = as.Date(paste0(Rain_Year, "-09-01")), # Expected start date
    Expected_End_Date = as.Date(paste0(Rain_Year + 1, "-06-30")), # Expected end date
    Expected_Duration = as.numeric(Expected_End_Date - Expected_Start_Date),
    Start_Date_Deviation = as.numeric(Observed_Start_Date - Expected_Start_Date),
    End_Date_Deviation = as.numeric(Observed_End_Date - Expected_End_Date),
    Duration_Deviation = Observed_Duration - Expected_Duration
  )

# Step 4: Format the summary table with one unique row per rain season
rainfall_summary_table <- rainfall_results %>%
  select(
    Rain_Year_Label,
    Observed_Start_Date,
    Observed_End_Date,
    Observed_Duration,
    Expected_Start_Date,
    Expected_End_Date,
    Expected_Duration,
    Start_Date_Deviation,
    End_Date_Deviation,
    Duration_Deviation,
    Total_Rainfall,
    Rainy_Days
  )


# Extract a random single row for each rain season
rainfall_summary_table <- rainfall_summary_table %>%
  filter(!(Rain_Year_Label %in% c("1999-2000", "2021-2022"))) %>% # Remove specified rain years
  group_by(Rain_Year_Label) %>% # Group by rain year
  slice_sample(n = 1) %>% # Randomly select one row per group
  ungroup() # Remove grouping


# Prepare tables to export
tables_to_export <- list(
  "Season Stats" = rain_season_analysis,
  "Season Shifts" = rainfall_summary_table
)

# Export tables to Excel
save_apa_formatted_excel(tables_to_export, "Seasonal_Shifts.xlsx")


library(ggplot2)

ggplot(rainfall_summary_table, aes(x = Rain_Year_Label, y = Duration_Deviation)) +
  geom_bar(stat = "identity", fill = "skyblue") +
  geom_text(
    aes(label = round(Duration_Deviation, 1)), # Display deviation values as labels
    vjust = ifelse(rainfall_summary_table$Duration_Deviation >= 0, -0.5, 1.5), # Position above for positive, below for negative
    size = 3, # Adjust text size
    color = "black" # Text color
  ) +
  labs(
    title = "Rainy Season Duration Deviation from Expected",
    x = "Rain Year",
    y = "Deviation (Days)"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))


ggplot(rainfall_summary_table, aes(x = Rain_Year_Label)) +
  geom_line(aes(y = Total_Rainfall, color = "Total Rainfall"), group = 1, size = 1) +
  geom_point(aes(y = Total_Rainfall, color = "Total Rainfall"), size = 3) +
  geom_line(aes(y = Rainy_Days, color = "Rainy Days"), group = 1, size = 1) +
  geom_point(aes(y = Rainy_Days, color = "Rainy Days"), size = 3) +
  labs(
    title = "Total Rainfall and Rainy Days Trends",
    x = "Rain Year",
    y = "Value",
    color = "Metric"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))


library(dplyr)

# Calculate descriptive statistics for each station
descriptive_table <- imputed_data %>%
  mutate(Year = year(Date)) %>% # Extract the year from the Date column
  group_by(Station, Year) %>%
  summarise(
    Total_Annual_Rainfall = sum(Rainfall_mm, na.rm = TRUE), # Total annual rainfall
    Avg_Annual_Temperature = mean(Avg_Temp, na.rm = TRUE), # Average annual temperature
    Avg_Annual_Humidity = mean(Daily_RH, na.rm = TRUE),    # Average annual relative humidity
    .groups = "drop"                                      # Ungroup after summarising
  ) %>%
  group_by(Station) %>%
  summarise(
    Avg_Total_Annual_Rainfall = mean(Total_Annual_Rainfall, na.rm = TRUE),
    Avg_Temperature = mean(Avg_Annual_Temperature, na.rm = TRUE),
    Avg_Relative_Humidity = mean(Avg_Annual_Humidity, na.rm = TRUE),
    .groups = "drop" # Ensure final table is not grouped
  )

# Print the descriptive table
print(descriptive_table)

library(dplyr)

# Calculate the proportion of missing data for each station
missing_data_summary <- merged_data %>%
  group_by(Station) %>%
  summarise(
    Proportion_Missing_Max_Temp = mean(is.na(Max_Temp)),
    Proportion_Missing_Min_Temp = mean(is.na(Min_Temp)),
    Proportion_Missing_Avg_Temp = mean(is.na(Avg_Temp)),
    Proportion_Missing_Rainfall = mean(is.na(Rainfall_mm)),
    Proportion_Missing_Rainy_Days = mean(is.na(Rainy_Days)),
    Proportion_Missing_Humidity = mean(is.na(Daily_RH)),
    Proportion_Missing_Wind_Speed = mean(is.na(Wind_Speed)),
    .groups = "drop" # Ungroup for a clean summary
  )

# Print the missing data summary
print(missing_data_summary)


# Save the descriptive table to a CSV file
write.csv(descriptive_table, "station_descriptive_table.csv", row.names = FALSE)

write.csv(missing_data_summary, "missing_table.csv", row.names = FALSE)

library(dplyr)

# Calculate the overall proportion of missing data for selected variables
overall_missingness <- merged_data %>%
  summarise(
    Proportion_Missing_Avg_Temp = mean(is.na(Avg_Temp)),
    Proportion_Missing_Humidity = mean(is.na(Daily_RH)),
    Proportion_Missing_Rainfall = mean(is.na(Rainfall_mm))
  )

# Print the overall missingness
print(overall_missingness)

# Save the summary to a CSV file
write.csv(overall_missingness, "overall_missingness_summary.csv", row.names = FALSE)


library(dplyr)

# Calculate the overall proportion of missing data for all variables together at each station
missingness_by_station <- merged_data %>%
  group_by(Station) %>%
  summarise(
    Total_Records = n(),
    Missing_Avg_Temp = sum(is.na(Avg_Temp)),
    Missing_Humidity = sum(is.na(Daily_RH)),
    Missing_Rainfall = sum(is.na(Rainfall_mm)),
    Total_Missing = Missing_Avg_Temp + Missing_Humidity + Missing_Rainfall,
    Proportion_Missing = Total_Missing / (Total_Records * 3) # Considering three variables
  )

# Print the missingness table
print(missingness_by_station)

# Save the results to a CSV file
write.csv(missingness_by_station, "missingness_by_station.csv", row.names = FALSE)



library(lubridate)

# Calculate average observed start and end dates (day of the year)
average_start_end <- rainfall_results %>%
  summarise(
    Avg_Observed_Start = mean(yday(Observed_Start_Date), na.rm = TRUE),
    Avg_Observed_End = mean(yday(Observed_End_Date), na.rm = TRUE)
  )

# Extract averages as day-of-year integers
avg_start_doy <- round(average_start_end$Avg_Observed_Start)
avg_end_doy <- round(average_start_end$Avg_Observed_End)

# Adjust deviations to be relative to average observed dates (day of year)
rainfall_results <- rainfall_results %>%
  mutate(
    Start_Date_DayOfYear = yday(Observed_Start_Date),
    End_Date_DayOfYear = yday(Observed_End_Date),
    Avg_Observed_Start_Date = avg_start_doy,
    Avg_Observed_End_Date = avg_end_doy,
    Start_Date_Deviation = Start_Date_DayOfYear - Avg_Observed_Start_Date,
    End_Date_Deviation = End_Date_DayOfYear - Avg_Observed_End_Date,
    Duration_Deviation = Observed_Duration - (avg_end_doy - avg_start_doy)
  )


# Adjust deviations and filter out specified years
rainfall_results <- rainfall_results %>%
  # Filter out unwanted years
  filter(!(Rain_Year_Label %in% c("1999-2000", "2021-2022"))) %>%
  # Add day-of-year deviations
  mutate(
    Start_Date_DayOfYear = yday(Observed_Start_Date),
    End_Date_DayOfYear = yday(Observed_End_Date),
    Avg_Observed_Start_Date = avg_start_doy,
    Avg_Observed_End_Date = avg_end_doy,
    Start_Date_Deviation = Start_Date_DayOfYear - Avg_Observed_Start_Date,
    End_Date_Deviation = End_Date_DayOfYear - Avg_Observed_End_Date,
    Duration_Deviation = Observed_Duration - (avg_end_doy - avg_start_doy)
  ) %>%
  # Randomly select one row per season
  group_by(Rain_Year_Label) %>%
  slice_sample(n = 1) %>%
  ungroup()



write.csv(rainfall_results, "rainfall_results.csv", row.names = FALSE)


library(ggplot2)
library(lubridate)

# Prepare data for plotting
plot_data <- rainfall_results %>%
  mutate(
    Start_Date_DayOfYear = yday(Observed_Start_Date),
    End_Date_DayOfYear = yday(Observed_End_Date),
    Rain_Year_Label = factor(Rain_Year_Label) # Ensure Rain_Year_Label is a factor for ordering
  )

# Calculate averages for plotting
avg_start_doy <- mean(plot_data$Start_Date_DayOfYear, na.rm = TRUE)
avg_end_doy <- mean(plot_data$End_Date_DayOfYear, na.rm = TRUE)

# Plot observed start and end dates with average lines
ggplot(plot_data) +
  # Points for start dates
  geom_point(aes(x = Rain_Year_Label, y = Start_Date_DayOfYear, color = "Start Date"), size = 3) +
  # Points for end dates
  geom_point(aes(x = Rain_Year_Label, y = End_Date_DayOfYear, color = "End Date"), size = 3) +
  # Average start date line
  geom_hline(aes(yintercept = avg_start_doy, color = "Average Start Date"), linetype = "dashed", size = 1) +
  # Average end date line
  geom_hline(aes(yintercept = avg_end_doy, color = "Average End Date"), linetype = "dashed", size = 1) +
  # Labels and customization
  labs(
    title = "Observed Start and End Dates of Wet Seasons",
    x = "Rain Year",
    y = "Day of Year",
    color = "Legend"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
  scale_color_manual(values = c("Start Date" = "blue", "End Date" = "red",
                                "Average Start Date" = "darkblue", "Average End Date" = "darkred"))


library(dplyr)
library(lubridate)

# Prepare data for modeling
model_data <- rainfall_results %>%
  mutate(
    Year = as.numeric(sub("-.*", "", Rain_Year_Label)), # Extract the year from Rain_Year_Label
    Start_DayOfYear = yday(Observed_Start_Date),
    End_DayOfYear = yday(Observed_End_Date)
  ) %>%
  select(Year, Start_DayOfYear, End_DayOfYear)

# Check the prepared dataset
print(model_data)


library(forecast)

# Fit time series models
start_ts <- ts(model_data$Start_DayOfYear, start = min(model_data$Year), frequency = 1)
end_ts <- ts(model_data$End_DayOfYear, start = min(model_data$Year), frequency = 1)

# Forecast for next 5 years
start_forecast <- forecast(start_ts, h = 5)
end_forecast <- forecast(end_ts, h = 5)

# View predictions
print(start_forecast)
print(end_forecast)

# Plot the forecasts
autoplot(start_forecast) + ggtitle("Start Date Forecast")
autoplot(end_forecast) + ggtitle("End Date Forecast")



library(ggplot2)

# Define future years
future_years <- data.frame(Year = seq(max(model_data$Year) + 1, max(model_data$Year) + 10))

# Fit linear regression models (if not already done)
start_model <- lm(Start_DayOfYear ~ Year, data = model_data)
end_model <- lm(End_DayOfYear ~ Year, data = model_data)

# Predict with confidence intervals
start_predictions <- predict(start_model, newdata = future_years, interval = "confidence")
end_predictions <- predict(end_model, newdata = future_years, interval = "confidence")

# Prepare data for plotting
prediction_data <- future_years %>%
  mutate(
    Predicted_Start = start_predictions[, "fit"],
    Start_Lower_CI = start_predictions[, "lwr"],
    Start_Upper_CI = start_predictions[, "upr"],
    Predicted_End = end_predictions[, "fit"],
    End_Lower_CI = end_predictions[, "lwr"],
    End_Upper_CI = end_predictions[, "upr"]
  )

# Combine observed and predicted data
visualization_data <- model_data %>%
  rename(Observed_Start = Start_DayOfYear, Observed_End = End_DayOfYear) %>%
  mutate(Type = "Observed") %>%
  bind_rows(
    prediction_data %>%
      rename(Observed_Start = Predicted_Start, Observed_End = Predicted_End) %>%
      mutate(Type = "Predicted")
  )

# Plot Start Dates with Confidence Intervals
ggplot(visualization_data, aes(x = Year)) +
  # Observed points
  geom_point(data = filter(visualization_data, Type == "Observed"),
             aes(y = Observed_Start, color = "Observed Start"), size = 3) +
  # Predicted line
  geom_line(data = filter(visualization_data, Type == "Predicted"),
            aes(y = Observed_Start, color = "Predicted Start"), size = 1) +
  # Confidence intervals
  geom_ribbon(data = prediction_data,
              aes(ymin = Start_Lower_CI, ymax = Start_Upper_CI, fill = "Start CI"), alpha = 0.2) +
  # Labels and themes
  labs(
    title = "Predicted and Observed Start Dates with Confidence Intervals",
    x = "Year",
    y = "Day of Year",
    color = "Legend",
    fill = "Legend"
  ) +
  theme_minimal() +
  scale_color_manual(values = c("Observed Start" = "blue", "Predicted Start" = "darkblue")) +
  scale_fill_manual(values = c("Start CI" = "lightblue"))

# Plot End Dates with Confidence Intervals
ggplot(visualization_data, aes(x = Year)) +
  # Observed points
  geom_point(data = filter(visualization_data, Type == "Observed"),
             aes(y = Observed_End, color = "Observed End"), size = 3) +
  # Predicted line
  geom_line(data = filter(visualization_data, Type == "Predicted"),
            aes(y = Observed_End, color = "Predicted End"), size = 1) +
  # Confidence intervals
  geom_ribbon(data = prediction_data,
              aes(ymin = End_Lower_CI, ymax = End_Upper_CI, fill = "End CI"), alpha = 0.2) +
  # Labels and themes
  labs(
    title = "Predicted and Observed End Dates with Confidence Intervals",
    x = "Year",
    y = "Day of Year",
    color = "Legend",
    fill = "Legend"
  ) +
  theme_minimal() +
  scale_color_manual(values = c("Observed End" = "red", "Predicted End" = "darkred")) +
  scale_fill_manual(values = c("End CI" = "pink"))



library(ggplot2)

# Plot Start Dates with Confidence Intervals and Data Labels
ggplot(visualization_data, aes(x = Year)) +
  # Observed points
  geom_point(data = filter(visualization_data, Type == "Observed"),
             aes(y = Observed_Start, color = "Observed Start"), size = 3) +
  # Predicted line
  geom_line(data = filter(visualization_data, Type == "Predicted"),
            aes(y = Observed_Start, color = "Predicted Start"), size = 1) +
  # Confidence intervals
  geom_ribbon(data = prediction_data,
              aes(ymin = Start_Lower_CI, ymax = Start_Upper_CI, fill = "Start CI"), alpha = 0.2) +
  # Add labels for upper and lower confidence intervals
  geom_text(data = prediction_data,
            aes(x = Year, y = Start_Lower_CI, label = round(Start_Lower_CI, 0)), 
            vjust = 1.5, color = "blue", size = 3.5) +
  geom_text(data = prediction_data,
            aes(x = Year, y = Start_Upper_CI, label = round(Start_Upper_CI, 0)), 
            vjust = -0.5, color = "blue", size = 3.5) +
  # Labels and themes
  labs(
    title = "Predicted and Observed Start Dates with Confidence Intervals",
    x = "Year",
    y = "Day of Year",
    color = "Legend",
    fill = "Legend"
  ) +
  theme_minimal() +
  scale_color_manual(values = c("Observed Start" = "blue", "Predicted Start" = "darkblue")) +
  scale_fill_manual(values = c("Start CI" = "lightblue"))

# Plot End Dates with Confidence Intervals and Data Labels
ggplot(visualization_data, aes(x = Year)) +
  # Observed points
  geom_point(data = filter(visualization_data, Type == "Observed"),
             aes(y = Observed_End, color = "Observed End"), size = 3) +
  # Predicted line
  geom_line(data = filter(visualization_data, Type == "Predicted"),
            aes(y = Observed_End, color = "Predicted End"), size = 1) +
  # Confidence intervals
  geom_ribbon(data = prediction_data,
              aes(ymin = End_Lower_CI, ymax = End_Upper_CI, fill = "End CI"), alpha = 0.2) +
  # Add labels for upper and lower confidence intervals
  geom_text(data = prediction_data,
            aes(x = Year, y = End_Lower_CI, label = round(End_Lower_CI, 0)), 
            vjust = 1.5, color = "red", size = 3.5) +
  geom_text(data = prediction_data,
            aes(x = Year, y = End_Upper_CI, label = round(End_Upper_CI, 0)), 
            vjust = -0.5, color = "red", size = 3.5) +
  # Labels and themes
  labs(
    title = "Predicted and Observed End Dates with Confidence Intervals",
    x = "Year",
    y = "Day of Year",
    color = "Legend",
    fill = "Legend"
  ) +
  theme_minimal() +
  scale_color_manual(values = c("Observed End" = "red", "Predicted End" = "darkred")) +
  scale_fill_manual(values = c("End CI" = "pink"))


library(ggplot2)

# Calculate deviations from the average for start and end dates
plot_data <- plot_data %>%
  mutate(
    Start_Deviation = Start_Date_DayOfYear - avg_start_doy,
    End_Deviation = End_Date_DayOfYear - avg_end_doy
  )

# Plot with deviations labeled
ggplot(plot_data) + 
  # Points for start dates
  geom_point(aes(x = Rain_Year_Label, y = Start_Date_DayOfYear, color = "Start Date"), size = 3) + 
  # Points for end dates
  geom_point(aes(x = Rain_Year_Label, y = End_Date_DayOfYear, color = "End Date"), size = 3) + 
  # Average start date line
  geom_hline(aes(yintercept = avg_start_doy, color = "Average Start Date"), linetype = "dashed", size = 1) + 
  # Average end date line
  geom_hline(aes(yintercept = avg_end_doy, color = "Average End Date"), linetype = "dashed", size = 1) + 
  # Labels for start date deviations
  geom_text(aes(x = Rain_Year_Label, y = Start_Date_DayOfYear, label = round(Start_Deviation, 1)), 
            vjust = -1, color = "blue", size = 3.5, data = plot_data) + 
  # Labels for end date deviations
  geom_text(aes(x = Rain_Year_Label, y = End_Date_DayOfYear, label = round(End_Deviation, 1)), 
            vjust = 1.5, color = "red", size = 3.5, data = plot_data) + 
  # Labels and customization
  labs(
    title = "Observed Start and End Dates of Wet Seasons with Deviations",
    x = "Rain Year",
    y = "Day of Year",
    color = "Legend"
  ) + 
  theme_minimal() + 
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) + 
  scale_color_manual(values = c("Start Date" = "blue", "End Date" = "red", 
                                "Average Start Date" = "darkblue", "Average End Date" = "darkred"))


# Load required libraries
library(dplyr)
library(zoo)

# Calculate 6-day rolling averages for Min_Temp, Max_Temp, and Avg_Temp
overall_data <- overall_data %>%
  arrange(Date) %>% # Ensure data is sorted by date
  mutate(
    Rolling_Min_Temp = rollapply(Min_Temp, width = 6, FUN = mean, fill = NA, align = "right"),
    Rolling_Max_Temp = rollapply(Max_Temp, width = 6, FUN = mean, fill = NA, align = "right"),
    Rolling_Avg_Temp = rollapply(Avg_Temp, width = 6, FUN = mean, fill = NA, align = "right")
  )


# Add a column for the seasonal year
overall_data <- overall_data %>%
  mutate(
    Seasonal_Year = if_else(
      month(Date) > 6 | (month(Date) == 6 & day(Date) >= 21), 
      year(Date), 
      year(Date) - 1
    )
  )


# Calculate the statistics for T_min and T_max by seasonal year
seasonal_stats <- overall_data %>%
  group_by(Seasonal_Year) %>%
  summarise(
    Tmin_Highest = max(Min_Temp, na.rm = TRUE),
    Tmin_Lowest = min(Min_Temp, na.rm = TRUE),
    Tmin_25th = quantile(Min_Temp, 0.25, na.rm = TRUE),
    Tmin_75th = quantile(Min_Temp, 0.75, na.rm = TRUE),
    Tmax_Highest = max(Max_Temp, na.rm = TRUE),
    Tmax_Lowest = min(Max_Temp, na.rm = TRUE),
    Tmax_25th = quantile(Max_Temp, 0.25, na.rm = TRUE),
    Tmax_75th = quantile(Max_Temp, 0.75, na.rm = TRUE),
    .groups = "drop"
  )

# Merge the calculated statistics back into the original dataset
overall_data <- overall_data %>%
  left_join(seasonal_stats, by = "Seasonal_Year")


real_season_logic <- function(data) {
  data <- data %>% arrange(Date)
  
  # Initialize the first season
  data <- data %>%
    mutate(Real_Season = NA_character_) %>%
    mutate(Real_Season = if_else(Date == as.Date("2000-05-01"), "Spring", Real_Season))
  
  # Loop through the data to assign seasons sequentially
  for (i in 2:nrow(data)) {
    # Get the current season
    current_season <- data$Real_Season[i - 1]
    
    # Skip if any required value is missing
    if (is.na(current_season) || is.na(data$Rolling_Max_Temp[i]) || is.na(data$Rolling_Min_Temp[i])) {
      next
    }
    
    # Retrieve the constant thresholds for the current row
    tmax_75th <- data$Tmax_75th[i]
    tmin_25th <- data$Tmin_25th[i]
    
    # Evaluate conditions for the next season
    if (current_season == "Spring" && data$Rolling_Max_Temp[i] >= tmax_75th) {
      data$Real_Season[i] <- "Summer"
    } else if (current_season == "Summer" && data$Rolling_Max_Temp[i] < tmax_75th) {
      data$Real_Season[i] <- "Autumn"
    } else if (current_season == "Autumn" && data$Rolling_Min_Temp[i] <= tmin_25th) {
      data$Real_Season[i] <- "Winter"
    } else if (current_season == "Winter" && data$Rolling_Min_Temp[i] > tmin_25th) {
      data$Real_Season[i] <- "Spring"
    } else {
      data$Real_Season[i] <- current_season
    }
  }
  
  return(data)
}

# Apply the updated function
overall_data <- real_season_logic(overall_data)



real_season_logic <- function(data) {
  data <- data %>% arrange(Date)
  
  # Initialize Real_Season column
  data <- data %>%
    mutate(Real_Season = NA_character_) %>%
    mutate(Real_Season = if_else(Date == as.Date("2000-05-01"), "Spring", Real_Season))
  
  # Initialize the current season
  current_season <- "Spring"
  
  # Loop through the data to assign seasons
  for (i in seq_len(nrow(data))) {
    # Skip if any required value is missing
    if (is.na(data$Rolling_Max_Temp[i]) || is.na(data$Rolling_Min_Temp[i])) {
      next
    }
    
    # Set thresholds for the current row
    tmax_75th <- data$Tmax_75th[i]
    tmin_25th <- data$Tmin_25th[i]
    
    # Determine flags for conditions over a rolling window
    summer_flag <- all(data$Rolling_Max_Temp[i:min(i + 8, nrow(data))] >= tmax_75th, na.rm = TRUE)
    autumn_flag <- all(data$Rolling_Max_Temp[i:min(i + 8, nrow(data))] < tmax_75th, na.rm = TRUE)
    winter_flag <- all(data$Rolling_Min_Temp[i:min(i + 8, nrow(data))] <= tmin_25th, na.rm = TRUE)
    spring_flag <- all(data$Rolling_Min_Temp[i:min(i + 8, nrow(data))] > tmin_25th, na.rm = TRUE)
    
    # Assign seasons based on current season and flags
    if (current_season == "Spring" && summer_flag) {
      current_season <- "Summer"
    } else if (current_season == "Summer" && autumn_flag) {
      current_season <- "Autumn"
    } else if (current_season == "Autumn" && winter_flag) {
      current_season <- "Winter"
    } else if (current_season == "Winter" && spring_flag) {
      current_season <- "Spring"
    }
    
    # Assign the current season to the Real_Season column
    data$Real_Season[i] <- current_season
  }
  
  return(data)
}

# Apply the updated function
overall_data <- real_season_logic(overall_data)

# View the result
head(overall_data)

library(dplyr)

# Step 1: Create Season_Change column, handling NA values
overall_data <- overall_data %>%
  mutate(
    Season_Change = cumsum(
      if_else(
        is.na(Real_Season) | is.na(lag(Real_Season, default = first(Real_Season))),
        FALSE, # Do not increment Season_Change if either value is NA
        Real_Season != lag(Real_Season, default = first(Real_Season))
      )
    )
  )

# Step 2: Verify the results
# Check the first few rows to confirm correct Season_Change assignment
head(overall_data, 20)

# Check unique combinations of Season_Change and Real_Season
unique_season_changes <- overall_data %>%
  select(Season_Change, Real_Season) %>%
  distinct()

# Print unique season changes
print(unique_season_changes)


# Group by Season_Change to calculate statistics for each unique season
season_statistics <- overall_data %>%
  group_by(Season_Change, Real_Season) %>%
  summarise(
    Start_Date = min(Date, na.rm = TRUE),
    End_Date = max(Date, na.rm = TRUE),
    Length_Days = as.numeric(End_Date - Start_Date + 1), # Include both start and end dates
    .groups = "drop" # Ungroup after summarisation
  ) %>%
  filter(!is.na(Real_Season)) # Exclude rows with NA Real_Season

# View the season statistics
print(season_statistics)

# Save as a CSV file if needed
write.csv(season_statistics, "Season_Statistics.csv", row.names = FALSE)

# Convert Start_Date and End_Date to day-of-year (DOY) to calculate averages
season_statistics <- season_statistics %>%
  mutate(
    Start_DOY = yday(Start_Date),
    End_DOY = yday(End_Date)
  )

# Aggregate statistics by season
aggregated_season_statistics <- season_statistics %>%
  group_by(Real_Season) %>%
  summarise(
    Avg_Length_Days = mean(Length_Days, na.rm = TRUE), # Average length
    Avg_Start_DOY = mean(Start_DOY, na.rm = TRUE), # Average start date (in DOY)
    Avg_End_DOY = mean(End_DOY, na.rm = TRUE), # Average end date (in DOY)
    Avg_Start_Date = as.Date(paste0("2000-", round(Avg_Start_DOY)), format = "%Y-%j"), # Convert DOY to Date
    Avg_End_Date = as.Date(paste0("2000-", round(Avg_End_DOY)), format = "%Y-%j"), # Convert DOY to Date
    .groups = "drop"
  )

# View the aggregated statistics
print(aggregated_season_statistics)

# Save as a CSV file if needed
write.csv(aggregated_season_statistics, "Aggregated_Season_Statistics.csv", row.names = FALSE)

# Save as a CSV file if needed
write.csv(season_statistics, "Season_Statistics.csv", row.names = FALSE)

str(season_statistics)

library(ggplot2)

# Assign a pseudo-seasonal year based on sequence of seasons
season_statistics <- season_statistics %>%
  mutate(Seasonal_Year = floor((Season_Change + 3) / 4) + 2000) # Assuming the first year is 2000


ggplot(season_statistics, aes(x = Seasonal_Year, y = Length_Days, color = Real_Season)) +
  geom_line() +
  geom_point(size = 2) +
  labs(
    title = "Season Length Over Time",
    x = "Seasonal Year",
    y = "Season Length (Days)",
    color = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(legend.position = "top")


ggplot(season_statistics, aes(x = Seasonal_Year)) +
  geom_line(aes(y = Start_DOY, color = Real_Season), size = 1.2) +
  geom_line(aes(y = End_DOY, linetype = Real_Season), size = 1.2) +
  labs(
    title = "Season Start and End Dates Over Time",
    x = "Seasonal Year",
    y = "Day of Year",
    color = "Season",
    linetype = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(legend.position = "top")


ggplot(season_statistics, aes(x = Seasonal_Year, y = Start_DOY, color = Real_Season)) +
  geom_point(size = 2) +
  geom_smooth(method = "lm", se = FALSE) +
  labs(
    title = "Trends in Season Start Dates",
    x = "Year",
    y = "Start Day of Year",
    color = "Season"
  ) +
  theme_minimal(base_size = 14) +
  theme(legend.position = "top")

