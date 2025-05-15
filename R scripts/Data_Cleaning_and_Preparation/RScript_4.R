library(openxlsx)
library(dplyr)
library(tidyr)
library(readr)
library(readxl)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/marcearcel")

df <- read.csv("MainQuestionaireMarcin.csv")


# Pivoting the data to the participant level
df_wide <- df %>%
  pivot_wider(names_from = "Question.Key", values_from = "Response")

# Renaming columns to only contain question numbers
colnames(df_wide) <- gsub("responseQ", "Q", colnames(df_wide))
colnames(df_wide) <- gsub("-quantised", "_quantised", colnames(df_wide))

# Remove columns ending with '_quantised'
df_wide <- df_wide %>%
  select(-ends_with("_quantised"))


## ADD GENDER AND AGE DATA
gender_data <- read_excel("GenderData.xlsx")
# Join the tables
final_data <- left_join(df_wide, gender_data, by = c("Participant.Private.ID" = "Participant"))

## RELIABILITY ANALYSIS

# Load the required libraries
library(dplyr)
library(psych)

# Function to create a list of data for a scale
get_scale_data <- function(scale_vars, scale_name, data) {
  scale_data <- lapply(scale_vars, function(var) {
    return(data.frame(
      Scale = scale_name,
      Variable = var,
      Mean = mean(data[[var]], na.rm = TRUE),
      SD = sd(data[[var]], na.rm = TRUE),
      Alpha = NA
    ))
  })
  
  # Add averaged scale data
  scale_data <- c(scale_data, list(data.frame(
    Scale = scale_name,
    Variable = paste(scale_name, "_Avg", sep = ""),
    Mean = mean(rowMeans(data[scale_vars], na.rm = TRUE), na.rm = TRUE),
    SD = sd(rowMeans(data[scale_vars], na.rm = TRUE), na.rm = TRUE),
    Alpha = psych::alpha(data[scale_vars])$total$raw_alpha
  )))
  
  return(scale_data)
}

# Define subscale variables
subscales_vars <- list(
  Avoidance_of_Femininity = c("Q4", "Q8", "Q10"),
  Negativity_toward_Sexual_Minorities = c("Q1", "Q5", "Q13"),
  Self_Reliance_through_Mechanical_Skills = c("Q6", "Q7", "Q14"),
  Toughness = c("Q17", "Q19", "Q20"),
  Dominance = c("Q2", "Q3", "Q12"),
  Importance_of_Sex = c("Q9", "Q11", "Q18"),
  Restrictive_Emotionality = c("Q15", "Q16", "Q21")
)

# Get data for each subscale
subscales_data <- lapply(names(subscales_vars), function(scale_name) {
  get_scale_data(subscales_vars[[scale_name]], scale_name, final_data)
})

# Create 'Total_Masculinity_Scale' by summing all the subscale questions
all_subscale_questions <- unlist(subscales_vars)
final_data$Total_Masculinity_Scale <- rowMeans(final_data[, all_subscale_questions], na.rm = TRUE)

# Add total scale data
total_scale_data <- list(data.frame(
  Scale = "Total_Masculinity_Scale",
  Variable = "Total_Masculinity_Scale",
  Mean = mean(final_data$Total_Masculinity_Scale, na.rm = TRUE),
  SD = sd(final_data$Total_Masculinity_Scale, na.rm = TRUE),
  Alpha = psych::alpha(final_data[, unlist(subscales_vars)], na.rm = TRUE)$total$raw_alpha
))

# Flatten the list of subscale data
subscales_data_flat <- do.call(rbind, unlist(subscales_data, recursive = FALSE))

# Combine the flattened subscale data with the total scale data
reliability_summary <- rbind(subscales_data_flat, total_scale_data[[1]])

# Compute the subscale scores
final_data$Avoidance_of_Femininity <- rowMeans(final_data[, c("Q4", "Q8", "Q10")], na.rm = TRUE)
final_data$Negativity_toward_Sexual_Minorities <- rowMeans(final_data[, c("Q1", "Q5", "Q13")], na.rm = TRUE)
final_data$Self_Reliance_through_Mechanical_Skills <- rowMeans(final_data[, c("Q6", "Q7", "Q14")], na.rm = TRUE)
final_data$Toughness <- rowMeans(final_data[, c("Q17", "Q19", "Q20")], na.rm = TRUE)
final_data$Dominance <- rowMeans(final_data[, c("Q2", "Q3", "Q12")], na.rm = TRUE)
final_data$Importance_of_Sex <- rowMeans(final_data[, c("Q9", "Q11", "Q18")], na.rm = TRUE)
final_data$Restrictive_Emotionality <- rowMeans(final_data[, c("Q15", "Q16", "Q21")], na.rm = TRUE)


# Compute the total scale
final_data$Total_Scale <- rowMeans(final_data[, c("Q4", "Q8", "Q10", "Q1", "Q5", "Q13", "Q6", "Q7", "Q14", "Q17", "Q19", "Q20", "Q2", "Q3", "Q12", "Q9", "Q11", "Q18", "Q15", "Q16", "Q21")], na.rm = TRUE)



## DATA DISTRIBUTION EXAMINATION

# Load the ggplot2 package for plotting
library(ggplot2)

# Load the e1071 package for skewness and kurtosis
library(e1071)

# Create a data frame for boxplot
avg_data <- data.frame(
  Value = c(final_data$Avoidance_of_Femininity, final_data$Negativity_toward_Sexual_Minorities, final_data$Self_Reliance_through_Mechanical_Skills, final_data$Toughness, final_data$Dominance, final_data$Importance_of_Sex, final_data$Restrictive_Emotionality, final_data$Total_Scale),
  Scale = rep(c("Avoidance_of_Femininity", "Negativity_toward_Sexual_Minorities", "Self_Reliance_through_Mechanical_Skills", "Toughness", "Dominance", "Importance_of_Sex", "Restrictive_Emotionality", "Total_Scale"), each = nrow(final_data))
)

# Define the desired order
ordered_scales <- c("Avoidance_of_Femininity", "Negativity_toward_Sexual_Minorities", 
                    "Self_Reliance_through_Mechanical_Skills", "Toughness", "Dominance", 
                    "Importance_of_Sex", "Restrictive_Emotionality", "Total_Scale")

# Set the order in the data frame
avg_data$Scale <- factor(avg_data$Scale, levels = ordered_scales)

# Create the ggplot again
p <- ggplot(avg_data, aes(x = Scale, y = Value)) +
  geom_boxplot() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
  ylab("Average Scale Value") +
  ggtitle("Boxplot of Average Scale Values")

# Save the ggplot
ggsave("ordered_boxplot.png", plot = p, width = 10, height = 7)



# Create histograms
par(mfrow = c(1, 2)) # 1x2 grid for the plots

hist(final_data$Avoidance_of_Femininity, main = "Avoidance of Femininity", xlab = "Average Value")
hist(final_data$Negativity_toward_Sexual_Minorities, main = "Negativity toward Sexual Minorities", xlab = "Average Value")
hist(final_data$Self_Reliance_through_Mechanical_Skills, main = "Self-Reliance through Mechanical Skills", xlab = "Average Value")
hist(final_data$Toughness, main = "Toughness", xlab = "Average Value")
hist(final_data$Dominance, main = "Dominance", xlab = "Average Value")
hist(final_data$Importance_of_Sex, main = "Importance of Sex", xlab = "Average Value")
hist(final_data$Restrictive_Emotionality, main = "Restrictive Emotionality", xlab = "Average Value")
hist(final_data$Total_Scale, main = "Total Scale", xlab = "Average Value")

#save figures
# List of scales
scales_list <- c("Avoidance_of_Femininity", "Negativity_toward_Sexual_Minorities", 
                 "Self_Reliance_through_Mechanical_Skills", "Toughness", "Dominance", 
                 "Importance_of_Sex", "Restrictive_Emotionality", "Total_Scale")

# Loop through each scale
for (scale in scales_list) {
  # Create file name
  file_name <- paste0(scale, ".png")
  
  # Open the png device
  png(file_name, width = 800, height = 600)
  
  # Create histogram
  hist(final_data[[scale]], main = scale, xlab = "Average Value")
  
  # Close the png device
  dev.off()
}


## NORMALITY CHECK ##

# Function to calculate skewness, kurtosis, and Shapiro-Wilk test
calculate_stats <- function(scale_name, data_vec) {
  skew <- skewness(data_vec, na.rm = TRUE)
  kurt <- kurtosis(data_vec, na.rm = TRUE)
  shapiro_wilk <- shapiro.test(na.omit(data_vec))$p.value
  
  return(data.frame(
    Scale = scale_name,
    Skewness = skew,
    Kurtosis = kurt,
    `Shapiro-Wilk p-value` = shapiro_wilk
  ))
}

# List of subscales and the total scale
scales_list <- c("Avoidance_of_Femininity", "Negativity_toward_Sexual_Minorities", "Self_Reliance_through_Mechanical_Skills", "Toughness", "Dominance", "Importance_of_Sex", "Restrictive_Emotionality", "Total_Masculinity_Scale")

# Get data for each scale
normality_data <- lapply(scales_list, function(scale) {
  calculate_stats(scale, final_data[[scale]])
})

# Combine all data
normality_results <- do.call(rbind, normality_data)

# Print results
print(normality_results)

## SAMPLE CHARACTERIZATION

# Count the number of men and women
women_count <- sum(final_data$Gender == 1)
men_count <- sum(final_data$Gender == 2)

# Calculate the average age
average_age <- mean(final_data$Age)

# Save these values in the environment
women_count
men_count
average_age

library(dplyr)
library(tidyr)

# Filter out Gender = 0
filtered_data <- final_data %>%
  filter(Gender %in% c(1, 2))

# Calculate mean and SD for each gender and scale
gender_stats_mean <- filtered_data %>%
  select(Avoidance_of_Femininity, Negativity_toward_Sexual_Minorities, 
         Self_Reliance_through_Mechanical_Skills, Toughness, Dominance, 
         Importance_of_Sex, Restrictive_Emotionality, Total_Masculinity_Scale, Gender) %>%
  gather(key = "Scale", value = "Value", -Gender) %>%
  group_by(Scale, Gender) %>%
  summarise(Mean = mean(Value, na.rm = TRUE)) %>%
  spread(key = Gender, value = Mean)

# Rename columns for clarity
colnames(gender_stats_mean) <- c("Scale", "Mean_Female", "Mean_Male")

# Calculate SD
gender_stats_sd <- filtered_data %>%
  select(Avoidance_of_Femininity, Negativity_toward_Sexual_Minorities, 
         Self_Reliance_through_Mechanical_Skills, Toughness, Dominance, 
         Importance_of_Sex, Restrictive_Emotionality, Total_Masculinity_Scale, Gender) %>%
  gather(key = "Scale", value = "Value", -Gender) %>%
  group_by(Scale, Gender) %>%
  summarise(SD = sd(Value, na.rm = TRUE)) %>%
  spread(key = Gender, value = SD)

# Rename columns for clarity
colnames(gender_stats_sd) <- c("Scale", "SD_Female", "SD_Male")

# Join the mean and SD dataframes
gender_stats <- left_join(gender_stats_mean, gender_stats_sd, by = "Scale")

## PERFORM F-TEST BETWEEN GENDERS

# Test for equal variances
perform_f_test <- function(scale_name, data) {
  female_data <- data[data$Gender == 1, scale_name, drop = TRUE]
  male_data <- data[data$Gender == 2, scale_name, drop = TRUE]
  
  f_test_result <- var.test(female_data, male_data)
  
  return(data.frame(
    Scale = scale_name,
    F_statistic = f_test_result$statistic,
    p_value = f_test_result$p.value
  ))
}

# List of subscales
scales_list <- c("Avoidance_of_Femininity", "Negativity_toward_Sexual_Minorities", 
                 "Self_Reliance_through_Mechanical_Skills", "Toughness", "Dominance", 
                 "Importance_of_Sex", "Restrictive_Emotionality", "Total_Masculinity_Scale")

# Apply the F-test for each scale
f_test_results <- lapply(scales_list, function(scale) {
  perform_f_test(scale, filtered_data)
})

# Combine all data into a dataframe
equalvar_test_results_df <- do.call(rbind, f_test_results)

# Function to perform Welch's ANOVA for each scale
perform_anova_test <- function(scale_name, data) {
  anova_test_result <- oneway.test(data[[scale_name]] ~ data$Gender, var.equal = FALSE)
  
  return(data.frame(
    Scale = scale_name,
    F_statistic = anova_test_result$statistic,
    p_value = anova_test_result$p.value
  ))
}


# Apply the ANOVA test for each scale
anova_test_results <- lapply(scales_list, function(scale) {
  perform_anova_test(scale, filtered_data)
})

# Combine all data
anova_test_results_df <- do.call(rbind, anova_test_results)

# Join ANOVA test results with the gender_stats dataframe
gender_stats_with_anova_test <- left_join(gender_stats, anova_test_results_df, by = "Scale")


## PERFORM F-TEST BETWEEN AGE BRACKETS

# Create new variables for age brackets
filtered_data <- filtered_data %>%
  mutate(
    Age_Bracket_Women = case_when(
      Gender == 1 & Age >= 18 & Age <= 30 ~ "18-30",
      Gender == 1 & Age >= 31 ~ "31+",
      TRUE ~ NA_character_
    ),
    Age_Bracket_Men = case_when(
      Gender == 2 & Age >= 18 & Age <= 30 ~ "18-30",
      Gender == 2 & Age >= 31 ~ "31+",
      TRUE ~ NA_character_
    )
  )

# Test for equal variances

perform_f_test_agebrackets <- function(scale_name, data, age_bracket_var) {
  data_filtered <- data %>% filter(!is.na(!!sym(scale_name)), !is.na(!!sym(age_bracket_var)))
  
  group1_data <- data_filtered[data_filtered[[age_bracket_var]] == "18-30", scale_name, drop = TRUE]
  group2_data <- data_filtered[data_filtered[[age_bracket_var]] == "31+", scale_name, drop = TRUE]
  
  f_test_result <- var.test(group1_data, group2_data)
  
  return(data.frame(
    Scale = scale_name,
    F_statistic = f_test_result$statistic,
    p_value = f_test_result$p.value
  ))
}

# Define your age bracket variables for men and women
age_bracket_var_men = "Age_Bracket_Men"
age_bracket_var_women = "Age_Bracket_Women"

# Apply the F-test for each scale for men
f_test_results_men <- lapply(scales_list, function(scale) {
  perform_f_test_agebrackets(scale, filtered_data, age_bracket_var_men)
})

# Apply the F-test for each scale for women
f_test_results_women <- lapply(scales_list, function(scale) {
  perform_f_test_agebrackets(scale, filtered_data, age_bracket_var_women)
})

# Combine all data into a dataframe
equalvar_test_results_men_df <- do.call(rbind, f_test_results_men)
equalvar_test_results_women_df <- do.call(rbind, f_test_results_women)


# Modified Function to perform ANOVA for each scale
perform_anova_test <- function(scale_name, data, age_bracket_var) {
  data_filtered <- data %>% 
    filter(!is.na(!!sym(scale_name)), !is.na(!!sym(age_bracket_var)))
  
  if (length(unique(data_filtered[[age_bracket_var]])) < 2) {
    return(data.frame(
      Scale = scale_name,
      F_statistic = NA,
      p_value = NA
    ))
  }
  
  anova_test_result <- oneway.test(data_filtered[[scale_name]] ~ data_filtered[[age_bracket_var]], var.equal = TRUE)
  
  return(data.frame(
    Scale = scale_name,
    F_statistic = anova_test_result$statistic,
    p_value = anova_test_result$p.value
  ))
}


# Function to generate statistics for given age brackets
generate_stats_by_age <- function(age_bracket_var, data, scales_list) {
  # Filter out NA age brackets
  data <- data %>% filter(!is.na(!!sym(age_bracket_var)))
  # Calculate mean
  age_stats_mean <- data %>%
    select(all_of(c(scales_list, age_bracket_var))) %>%
    gather(key = "Scale", value = "Value", -all_of(age_bracket_var)) %>%
    group_by(Scale, !!sym(age_bracket_var)) %>%
    summarise(Mean = mean(Value, na.rm = TRUE)) %>%
    spread(key = !!sym(age_bracket_var), value = Mean)
  
  # Calculate SD
  age_stats_sd <- data %>%
    select(all_of(c(scales_list, age_bracket_var))) %>%
    gather(key = "Scale", value = "Value", -all_of(age_bracket_var)) %>%
    group_by(Scale, !!sym(age_bracket_var)) %>%
    summarise(SD = sd(Value, na.rm = TRUE)) %>%
    spread(key = !!sym(age_bracket_var), value = SD)
  
  # Join mean and SD dataframes
  age_stats <- left_join(age_stats_mean, age_stats_sd, by = "Scale", suffix = c("_Mean", "_SD"))
  
  # Perform Welch's ANOVA
  anova_results <- lapply(scales_list, function(scale) {
    perform_anova_test(scale, data, age_bracket_var)
  })
  
  anova_results_df <- do.call(rbind, anova_results)
  
  # Join mean, SD, and ANOVA results
  stats_with_anova <- left_join(age_stats, anova_results_df, by = "Scale")
  
  return(stats_with_anova)
}

# Generate statistics for women and men
stats_women_by_age <- generate_stats_by_age("Age_Bracket_Women", filtered_data, scales_list)
stats_men_by_age <- generate_stats_by_age("Age_Bracket_Men", filtered_data, scales_list)

## EXPORT RESULTS

# Importing the required libraries for data manipulation and Excel file writing
library(openxlsx)

# Assuming that the dataframes have been correctly generated as:
# reliability_summary, normality_results, equalvar_test_results_df,
# gender_stats_with_anova_test, equalvar_test_results_men_df,
# stats_men_be_age, equalvar_test_results_women_df, stats_women_be_age

# Define the Excel file path
excel_file_path <- "Results.xlsx"

# Create a new workbook
wb <- createWorkbook()

# Add worksheets and write dataframes to them
addWorksheet(wb, "Reliability Summary")
writeData(wb, "Reliability Summary", reliability_summary)

addWorksheet(wb, "Normality Results")
writeData(wb, "Normality Results", normality_results)

addWorksheet(wb, "Equal Var Test - Gender")
writeData(wb, "Equal Var Test - Gender", equalvar_test_results_df)

addWorksheet(wb, "Gender Stats with ANOVA")
writeData(wb, "Gender Stats with ANOVA", gender_stats_with_anova_test)

addWorksheet(wb, "Equal Var Test - Men by Age")
writeData(wb, "Equal Var Test - Men by Age", equalvar_test_results_men_df)

addWorksheet(wb, "Stats - Men by Age")
writeData(wb, "Stats - Men by Age", stats_men_by_age)

addWorksheet(wb, "Equal Var Test - Women by Age")
writeData(wb, "Equal Var Test - Women by Age", equalvar_test_results_women_df)

addWorksheet(wb, "Stats - Women by Age")
writeData(wb, "Stats - Women by Age", stats_women_by_age)

# Save the workbook to file
saveWorkbook(wb, excel_file_path, overwrite = TRUE)

