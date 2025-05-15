# 1) Histogram with Error Bars per Time Window

# Calculate the mean frequency by patient per time window
mean_stats_by_patient <- df_filtered_overall %>%
  group_by(patient_id, time_window) %>%
  summarise(mean_freq = n()) %>%
  ungroup()

# Calculate the standard deviation based on mean frequencies and the total frequency per time window
overall_stats <- mean_stats_by_patient %>%
  group_by(time_window) %>%
  summarise(
    total_freq = mean(mean_freq),  # total frequency across all patients
    sd_freq = sd(mean_freq)  # standard deviation of the mean frequencies
  )

# Bar plot with error bars that do not go below 0
p <- ggplot(overall_stats, aes(x = time_window, y = total_freq)) +
  geom_bar(stat = "identity") +
  geom_errorbar(aes(ymin = pmax(0, total_freq - sd_freq), ymax = total_freq + sd_freq), width = 0.2) +
  labs(title = "Overall Frequency of Meals by Time Window",
       x = "Time Window",
       y = "Total Frequency",
       subtitle = "Error bars represent standard deviation of the mean frequency per patient") +
  theme_minimal()

# Show plot
print(p)

# Save the plot
ggsave("Overall_Frequency_by_TimeWindow_with_ErrorBars.png", plot = p, width = 10, height = 6)


# 1) Histogram with Error Bars per Hour

# Calculate the mean frequency by patient per hour
mean_stats_by_patient_hour <- df_filtered_overall %>%
  group_by(patient_id, meal_time_hour) %>%
  summarise(mean_freq = n()) %>%
  ungroup()

# Calculate the standard deviation based on mean frequencies and the total frequency per hour
overall_stats_hour <- mean_stats_by_patient_hour %>%
  group_by(meal_time_hour) %>%
  summarise(
    total_freq = mean(mean_freq),  # total frequency across all patients
    sd_freq = sd(mean_freq)  # standard deviation of the mean frequencies
  )

# Bar plot with error bars that do not go below 0
p_hour <- ggplot(overall_stats_hour, aes(x = meal_time_hour, y = total_freq)) +
  geom_bar(stat = "identity") +
  geom_errorbar(aes(ymin = pmax(0, total_freq - sd_freq), ymax = total_freq + sd_freq), width = 0.2) +
  labs(title = "Overall Frequency of Meals by Hour",
       x = "Hour of Day",
       y = "Total Frequency",
       subtitle = "Error bars represent standard deviation of the mean frequency per patient") +
  theme_minimal()

# Show plot 
print(p_hour)

# Save the plot
ggsave("Overall_Frequency_by_Hour_with_ErrorBars.png", plot = p_hour, width = 10, height = 6)

## 2) LINE GRAPHS PER PATIENT

# Convert meal_date to Date object
df$meal_date <- as.Date(df$meal_date, format = "%m/%d/%Y")

# Add day_number based on the earliest meal_date for each patient
df <- df %>%
  group_by(patient_id) %>%
  mutate(day_number = as.integer(difftime(meal_date, min(meal_date), units = "days")) + 1) %>%
  ungroup()

# Group by patient_id, meal_id, and day_number, then summarize to get one record per meal_id per day
df_grouped <- df %>%
  group_by(patient_id, meal_id, day_number) %>%
  summarise(meal_time_hours = as.integer(format(min(meal_time), "%H"))) %>%
  ungroup()

# Calculate the frequency of meals per hour per day for each patient
df_grouped_frequency <- df_grouped %>%
  group_by(patient_id, day_number, meal_time_hours) %>%
  summarise(count = n()) %>%
  ungroup()

# Create line graphs per patient
unique_patients <- unique(df_grouped$patient_id)

# Add time_window column based on meal_time_hours
df_grouped_frequency <- df_grouped_frequency %>%
  mutate(time_window = case_when(
    meal_time_hours >= 22 | meal_time_hours < 6  ~ "22-6",
    meal_time_hours >= 6  & meal_time_hours < 11 ~ "6-11",
    meal_time_hours >= 11 & meal_time_hours < 15 ~ "11-15",
    meal_time_hours >= 15 & meal_time_hours < 18 ~ "15-18",
    meal_time_hours >= 18 & meal_time_hours < 22 ~ "18-22",
    TRUE ~ "Unknown"
  ))

# Make the time_window variable a factor with ordered levels for overall data
df_grouped_frequency$time_window <- factor(df_grouped_frequency$time_window, levels = c("22-6", "6-11", "11-15", "15-18", "18-22"))

for(patient in unique_patients) {
  patient_data <- subset(df_grouped_frequency, patient_id == patient)
  
  p <- ggplot(patient_data, aes(x = time_window, y = count, group = factor(day_number), color = factor(day_number))) +
    geom_line() +
    geom_point() +
    labs(title = paste("Frequency of Meals Per Time Window for Patient", patient),
         x = "Hour of Day",
         y = "Frequency of Meals",
         color = "Day Number") +
    theme_minimal() +
    scale_y_continuous(breaks = seq(min(patient_data$count), max(patient_data$count), by = 1))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(paste("Frequency_of_Meals_Time_Window_for_Patient_", patient, ".png"), plot = p, width = 10, height = 7)
}


# OVERALL LINE GRAPH

# Group by patient_id, meal_id, day_number, and meal_name
# and then summarize to get the minimum meal_time_hours
df_grouped <- df %>%
  group_by(patient_id, day_number, meal_name) %>%
  summarise(meal_time_hours = min(as.integer(format(meal_time, "%H")))) %>%
  ungroup()

# Get a list of unique patients
unique_patients <- unique(df_grouped$patient_id)

# Loop through each unique patient
for(patient in unique_patients) {
  
  # Filter the grouped data for the current patient
  patient_data <- subset(df_grouped, patient_id == patient & day_number <= 14)  # Filter data to 14-day study period
  
  # Create the graph
  p <- ggplot(patient_data, aes(x = day_number, y = meal_time_hours, group = meal_name, color = meal_name)) +
    geom_line(size = 2) +  # 3x line thickness
    geom_point() +
    labs(title = paste("Meal Timing for Patient", patient),
         x = "Day Number",
         y = "Time of Day (Hours)",
         color = "Meal Name") +
    theme_minimal() +
    scale_colour_manual(values = c("red", "green", "blue", "purple"),
                        breaks = c("dinner", "snack", "lunch", "breakfast"),
                        labels = c("Dinner", "Snack", "Lunch", "Breakfast")) +
    scale_y_continuous(breaks = seq(0, 23, 1), limits = c(0, 23)) +
    scale_x_continuous(breaks = seq(1, 14, 1), limits = c(1, 14))  # 14-day study period
  
  # Print the graph
  print(p)
  
  # Save the graph to a PNG file
  ggsave(paste("Meal_Timing_for_Patient_", patient, ".png"), plot = p, width = 10, height = 7)
}

# Install and load the openxlsx package if not already done
if (!requireNamespace("openxlsx", quietly = TRUE)) {
  install.packages("openxlsx")
}

library(openxlsx)

# Create a new workbook
wb <- createWorkbook()

# Add worksheets for each data frame
addWorksheet(wb, "df_heatmap_data_individual")
writeData(wb, "df_heatmap_data_individual", df_heatmap_data_individual)

addWorksheet(wb, "df_heatmap_data_overall")
writeData(wb, "df_heatmap_data_overall", df_heatmap_data_overall)

addWorksheet(wb, "mean_stats_by_patient")
writeData(wb, "mean_stats_by_patient", mean_stats_by_patient)

addWorksheet(wb, "overall_food_df")
writeData(wb, "overall_food_df", overall_food_df)

addWorksheet(wb, "overall_stats")
writeData(wb, "overall_stats", overall_stats)

addWorksheet(wb, "top_foods_per_patient")
writeData(wb, "top_foods_per_patient", top_foods_per_patient)

# Save workbook to your specified directory
saveWorkbook(wb, "Tables_Analysis.xlsx", overwrite = TRUE)

