library(readxl)             # Load the packages
library(tidyverse)
library(factoextra)
library(cluster)
library(dplyr)
library(tidyr)
library(ggplot2)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/paulhertz")
data <- read_excel("outfile_pdffactor_analysis (1).xlsx")

# Remove lines with blank responses (missing values)
data <- data[complete.cases(data), ]

# Select columns with Choice in their names
choices_columns <- select(data, contains("Choice"))

# Apply table function to each column to get the frequency table
frequency_table <- lapply(choices_columns, table)

# Convert frequency_table to data frames and combine them into a single table
combined_table <- bind_rows(lapply(frequency_table, as.data.frame), .id = "Question")

# Calculate the percentage for each choice in each question
combined_table <- combined_table %>%
  group_by(Question) %>%
  mutate(Percentage = Freq / sum(Freq) * 100)

# Print the combined frequency table with the percentage column
print(combined_table)




# Select columns with Choice and Strength in their names
choice_strength_columns <- select(data, contains(c("Choice", "Strength")))

# Gather the data into long format
gathered_data <- choice_strength_columns %>%
  pivot_longer(everything(), names_to = c("Question", ".value"), names_pattern = "(.*) (Choice|Strength)")

# Calculate the frequency for each question's "Choice" and "Strength" combinations
result_table <- gathered_data %>%
  group_by(Question, Choice, Strength) %>%
  summarise(Frequency = n()) %>%
  ungroup()

# Calculate the total frequency for each question separately
total_frequency <- result_table %>%
  group_by(Question) %>%
  summarise(Total = sum(Frequency)) %>%
  ungroup()

# Merge the result table with the total frequency and calculate the percentage
result_table <- left_join(result_table, total_frequency, by = "Question") %>%
  mutate(Percentage = (Frequency / Total) * 100) %>%
  select(-Total)

# Print the result table with the percentage column
print(result_table)


