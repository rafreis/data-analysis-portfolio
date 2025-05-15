library(ggplot2)
library(dplyr)


setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/elena_koning/elena_koning-attachments (1)")

df = read.csv("RxFood_Eating_Rhythm_Data.csv")


## HISTOGRAM OF MEAL CONSUMPTION

# Convert 'meal_time' to a Date-Time object
df$meal_time <- as.POSIXct(df$meal_time, format = "%H:%M:%S")

# Group by patient_id and meal_id, then summarize to get one record per meal_id
df_grouped <- df %>%
  group_by(patient_id, meal_id) %>%
  summarise(meal_time_hours = as.integer(format(min(meal_time), "%H")))

# Create a directory to store figures
dir.create("figures", showWarnings = FALSE)

# Loop over each unique patient_id to generate a plot
unique_patients <- unique(df_grouped$patient_id)

for (patient in unique_patients) {
  patient_data <- subset(df_grouped, patient_id == patient)
  
  # Count the frequencies of meals for the y-axis
  max_count <- max(table(patient_data$meal_time_hours))
  
  p <- ggplot(patient_data, aes(x = meal_time_hours)) +
    geom_histogram(binwidth = 1, fill = "orange", alpha = 0.7) + 
    scale_x_continuous(breaks = 0:23, labels = 0:23) +  # Integral hours
    scale_y_continuous(breaks = seq(0, max_count, by = 1)) +  # Integer values for Y-axis
    xlab("Time of Day (Hours)") +
    ylab("Frequency of Meals") +
    ggtitle(paste("Histogram of Meal Timings for Patient", patient))
  
  #Print Plot
  print(p)
  
  # Save plot as PNG
  filename <- paste("figures/Patient_", patient, "_Histogram.png", sep = "")
  ggsave(filename, plot = p, width = 10, height = 6)
}

# Count the frequencies of meals for the y-axis for the overall histogram
max_count_overall <- max(table(df_grouped$meal_time_hours))

# Create overall histogram
p <- ggplot(df_grouped, aes(x = meal_time_hours)) +
  geom_histogram(binwidth = 1, fill = "brown", alpha = 0.7) +
  scale_x_continuous(breaks = 0:23, labels = 0:23) +  # Integral hours
  scale_y_continuous(breaks = seq(0, max_count_overall, by = 1)) +  # Integer values for Y-axis
  xlab("Time of Day (Hours)") +
  ylab("Frequency of Meals") +
  ggtitle("Overall Histogram of Meal Timings")

#Print Plot
print(p)

# Save overall histogram
ggsave("figures/Overall_Histogram.png", plot = p, width = 10, height = 6)

# Zip the folder
zip("figures.zip", "figures/")


## BAR PLOTS OF FOOD TYPES

# Create a directory to store food frequency figures
dir.create("food_frequencies", showWarnings = FALSE)

# Loop over each unique patient_id to generate a food frequency plot
unique_patients <- unique(df$patient_id)

for (patient in unique_patients) {
  patient_data <- subset(df, patient_id == patient)
  
  # Count the frequency of each food item
  food_count <- table(patient_data$food_name)
  
  # Convert to dataframe and sort by frequency
  food_df <- as.data.frame(food_count)
  colnames(food_df) <- c("food_name", "frequency")
  food_df <- food_df[order(-food_df$frequency),]
  
  # Take the top 10 most consumed foods
  top_foods <- head(food_df, 10)
  
  # Create the plot
  p <- ggplot(top_foods, aes(x = reorder(food_name, frequency), y = frequency)) +
    geom_bar(stat = "identity", fill = "blue") +
    coord_flip() +
    xlab("Food Items") +
    ylab("Frequency") +
    ggtitle(paste("Top 10 Most Consumed Foods for Patient", patient))
  
  #Print Plot
  print(p)
  
  # Save the plot
  filename <- paste("food_frequencies/Patient_", patient, "_Food_Frequency.png", sep = "")
  ggsave(filename, plot = p, width = 10, height = 6)
}

# Count the frequency of each food item for all patients
overall_food_count <- table(df$food_name)

# Convert to dataframe and sort by frequency
overall_food_df <- as.data.frame(overall_food_count)
colnames(overall_food_df) <- c("food_name", "frequency")
overall_food_df <- overall_food_df[order(-overall_food_df$frequency),]

# Take the top 10 most consumed foods overall
top_foods_overall <- head(overall_food_df, 10)

# Create the overall plot
p <- ggplot(top_foods_overall, aes(x = reorder(food_name, frequency), y = frequency)) +
  geom_bar(stat = "identity", fill = "purple") +
  coord_flip() +
  xlab("Food Items") +
  ylab("Frequency") +
  ggtitle("Top 10 Most Consumed Foods Overall")

#Print Plot
print(p)

# Save the overall plot
p <- p + theme(plot.background = element_rect(fill = "white"))
ggsave("food_frequencies/Overall_Food_Frequency.png", plot = p, width = 10, height = 6)

# Zip the folder
zip("food_frequencies.zip", "food_frequencies/")


## JOINT VISUALIZATION - PER HOUR

# Get top 5 most consumed foods for each patient
top_foods_per_patient <- df %>%
  group_by(patient_id, food_name) %>%
  summarise(count = n()) %>%
  arrange(patient_id, desc(count)) %>%
  group_by(patient_id) %>%
  slice_head(n = 5)

# Get top 20 most consumed foods overall
top_foods_overall <- df %>%
  group_by(food_name) %>%
  summarise(count = n()) %>%
  arrange(desc(count)) %>%
  slice_head(n = 20)

# Create separate filtered data frames
df_filtered_overall <- df %>%
  filter(food_name %in% top_foods_overall$food_name)

df_filtered_individual <- df %>%
  filter(interaction(patient_id, food_name) %in% interaction(top_foods_per_patient$patient_id, top_foods_per_patient$food_name))

# Convert meal_time to hour for easier plotting
df_filtered_overall$meal_time_hour <- as.integer(format(df_filtered_overall$meal_time, "%H"))
df_filtered_individual$meal_time_hour <- as.integer(format(df_filtered_individual$meal_time, "%H"))


# Calculate the count of each food_name at each meal_time_hour for overall and individual
df_heatmap_data_overall <- df_filtered_overall %>% 
  count(meal_time_hour, food_name)

df_heatmap_data_individual <- df_filtered_individual %>% 
  count(patient_id, meal_time_hour, food_name)

# Overall Heatmap
p <- ggplot(df_heatmap_data_overall, aes(x = meal_time_hour, y = food_name)) +
  geom_tile(aes(fill = n), alpha = 1.0) +
  scale_fill_gradient(low = "lightblue", high = "darkblue") +
  labs(title = "Overall Heatmap of Food Types Consumed Over Time",
       x = "Hour of Day",
       y = "Food Type",
       fill = "Frequency") +
  theme_minimal()

# Save the overall plot
ggsave("food__type_frequencies/Overall_Food_Type_PerHour.png", plot = p, width = 10, height = 6)

# Individual Heatmaps
p <- ggplot(df_heatmap_data_individual, aes(x = meal_time_hour, y = food_name)) +
  geom_tile(aes(fill = n), alpha = 1.0) +
  scale_fill_gradient(low = "lightblue", high = "darkblue") +
  labs(title = "Heatmap of Food Types Consumed Over Time",
       x = "Hour of Day",
       y = "Food Type",
       fill = "Frequency") +
  theme_minimal() +
  facet_wrap(~ patient_id, scales = "free", ncol = 2)

# Save the plot
p <- p + theme(plot.background = element_rect(fill = "white"))
filename <- paste("food_type_frequencies/Patient_", patient, "_Food_Type_PerHour.png", sep = "")
ggsave(filename, plot = p, width = 10, height = 6)

## JOINT VISUALIZATION - PER HOUR WINDOW

library(forcats)

# Bin meal_time_hours into time ranges for overall data
df_filtered_overall <- df_filtered_overall %>%
  mutate(time_window = case_when(
    meal_time_hour >= 22 | meal_time_hour < 6  ~ "22-6",
    meal_time_hour >= 6 & meal_time_hour < 11  ~ "6-11",
    meal_time_hour >= 11 & meal_time_hour < 15 ~ "11-15",
    meal_time_hour >= 15 & meal_time_hour < 18 ~ "15-18",
    meal_time_hour >= 18 & meal_time_hour < 22 ~ "18-22",
    TRUE ~ "Other"
  ))

# Make the time_window variable a factor with ordered levels for overall data
df_filtered_overall$time_window <- factor(df_filtered_overall$time_window, levels = c("22-6", "6-11", "11-15", "15-18", "18-22"))

# Bin meal_time_hours into time ranges for individual data
df_filtered_individual <- df_filtered_individual %>%
  mutate(time_window = case_when(
    meal_time_hour >= 22 | meal_time_hour < 6  ~ "22-6",
    meal_time_hour >= 6 & meal_time_hour < 11  ~ "6-11",
    meal_time_hour >= 11 & meal_time_hour < 15 ~ "11-15",
    meal_time_hour >= 15 & meal_time_hour < 18 ~ "15-18",
    meal_time_hour >= 18 & meal_time_hour < 22 ~ "18-22",
    TRUE ~ "Other"
  ))

# Make the time_window variable a factor with ordered levels for individual data
df_filtered_individual$time_window <- factor(df_filtered_individual$time_window, levels = c("22-6", "6-11", "11-15", "15-18", "18-22"))

# Adjust counting and plotting
df_heatmap_data_overall <- df_filtered_overall %>% 
  count(time_window, food_name)

df_heatmap_data_individual <- df_filtered_individual %>% 
  count(patient_id, time_window, food_name)

# Overall Heatmap
p <- ggplot(df_heatmap_data_overall, aes(x = time_window, y = food_name)) +
  geom_tile(aes(fill = n), alpha = 2.0) +
  scale_fill_gradient(low = "pink", high = "darkred") +
  labs(title = "Overall Heatmap of Food Types Consumed Over Time",
       x = "Time Window",
       y = "Food Type",
       fill = "Frequency") +
  theme_minimal()

# Save the overall plot
p <- p + theme(plot.background = element_rect(fill = "white"))
ggsave("food__type_frequencies/Overall_Food_Type_PerHourWindow.png", plot = p, width = 10, height = 6)

# Individual Heatmaps
p <- ggplot(df_heatmap_data_individual, aes(x = time_window, y = food_name)) +
  geom_tile(aes(fill = n), alpha = 2.0) +
  scale_fill_gradient(low = "pink", high = "darkred") +
  labs(title = "Heatmap of Food Types Consumed Over Time by Individual",
       x = "Time Window",
       y = "Food Type",
       fill = "Frequency") +
  theme_minimal() +
  facet_wrap(~ patient_id, scales = "free", ncol = 2)

# Save the plot
p <- p + theme(plot.background = element_rect(fill = "white"))
filename <- paste("food_type_frequencies/Patient_", patient, "_Food_Type_PerHourWindow.png", sep = "")
ggsave(filename, plot = p, width = 10, height = 6)

## JOINT VISUALIZATION - MEAL NAME

# Calculate the count of each food_name at each meal_name
df_heatmap_data_overall <- df_filtered_overall %>% 
  count(meal_name, food_name)

# Calculate the count of each food_name at each meal_name, grouped by patient_id
df_heatmap_data_individual <- df_filtered_individual %>% 
  count(patient_id, meal_name, food_name)

# Adjust your data frames to set the order of meal_name
df_heatmap_data_overall$meal_name <- factor(df_heatmap_data_overall$meal_name, 
                                    levels = c("breakfast", "lunch", "snack", "dinner"))

df_heatmap_data_individual$meal_name <- factor(df_heatmap_data_individual$meal_name, 
                                               levels = c("breakfast", "lunch", "snack", "dinner"))

# Overall Heatmap
p <- ggplot(df_heatmap_data_overall, aes(x = meal_name, y = food_name)) +
  geom_tile(aes(fill = n), alpha = 1.0) +
  scale_fill_gradient(low = "lightblue", high = "darkblue") +
  labs(title = "Overall Heatmap of Food Types Consumed Over Time",
       x = "Meal Name",
       y = "Food Type",
       fill = "Frequency") +
  theme_minimal()

# Save the overall plot
p <- p + theme(plot.background = element_rect(fill = "white"))
ggsave("meal_name_frequencies/Overall_Food_MealNameFrequencies.png", plot = p, width = 10, height = 6)

# Individual Heatmaps
p <- ggplot(df_heatmap_data_individual, aes(x = meal_name, y = food_name)) +
  geom_tile(aes(fill = n), alpha = 1.0) +
  scale_fill_gradient(low = "lightblue", high = "darkblue") +
  labs(title = "Heatmap of Food Types Consumed Over Time",
       x = "Meal Name",
       y = "Food Type",
       fill = "Frequency") +
  theme_minimal() +
  facet_wrap(~ patient_id, scales = "free", ncol = 2)

# Save the plot
p <- p + theme(plot.background = element_rect(fill = "white"))
filename <- paste("meal_name_frequencies/Patient_", patient, "_Food_MealNameFrequencies.png.png", sep = "")
ggsave(filename, plot = p, width = 10, height = 6)

