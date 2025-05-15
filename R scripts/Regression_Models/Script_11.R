# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/bel_yuv")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("data_corrected2.xlsx")

# Get rid of special characters

names(df) <- gsub(" ", "_", trimws(names(df)))
names(df) <- gsub("\\s+", "_", trimws(names(df), whitespace = "[\\h\\v\\s]+"))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))
names(df) <- gsub("\\$", "", names(df))
names(df) <- gsub("\\+", "_", names(df))

# Trim all values

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))



# Trim all values and convert to lowercase

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
    # Convert character strings to lowercase
    x <- tolower(x)
  }
  return(x)
}))


str(df)


# Identify character columns
char_cols <- sapply(df, is.character)

# Print unique values for each character column
for (col in names(df)[char_cols]) {
  cat("\nUnique values in", col, ":\n")
  print(unique(df[[col]]))
}


library(dplyr)

df <- df %>%
  mutate(
    gender = case_when(
      gender == "male" ~ "Male",
      gender == "m" ~ "Male",
      gender == "female" ~ "Female",
      gender == "f" ~ "Female",
      TRUE ~ gender
    ),
    
    moi = case_when(
      moi == "stabss" ~ "stab",
      moi == "stabs" ~ "stab",
      moi == "mva" ~ "motor vehicle accident",
      moi == "ffh" ~ "fall from height",
      moi == "pva" ~ "pedestrian vs auto",
      moi == "fall" ~ "fall",
      moi == "blunt (wall fell on patient)" ~ "blunt trauma",
      moi == "blunt (assault)" ~ "blunt trauma",
      moi == "gsw" ~ "gunshot wound",
      moi == "blunt (kicked by cow)" ~ "blunt trauma",
      moi == "blunt (rugby)" ~ "blunt trauma",
      moi == "blunt" ~ "blunt trauma",
      TRUE ~ moi
    )
  )

create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each category variable
  for (category in categories) {
    # Ensure the category variable is a factor
    data[[category]] <- factor(data[[category]])
    
    # Calculate counts
    counts <- table(data[[category]])
    
    # Create a dataframe for this category
    freq_table <- data.frame(
      "Category" = rep(category, length(counts)),
      "Level" = names(counts),
      "Count" = as.integer(counts),
      stringsAsFactors = FALSE
    )
    
    # Calculate and add percentages
    freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
    
    # Add the result to the list
    all_freq_tables[[category]] <- freq_table
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

vars_cat <- c("gender", "moi"             ,                "level_of_injury"       ,          "asia_score"              ,       
              "recovery_discharge" ,             "management_descriptive" ,         "management_categorical"   ,      
              "date_of_call"        ,            "consent"                 ,        "speaker"                   ,     
              "reason_for_patient_not_speaking", "follow_up"                ,       "rehab"                      ,    
              "rehab_discipline"      ,          "doctor_review"             ,      "level_of_facility_review"    ,   
              "type_of_treatment"      ,         "last_neurosurgery_review"   ,     "last_rehabilitation_review"   ,  
              "last_gp_review"          ,        "barriers_to_care"            ,    "recovery_from_discharge"       , 
              "life_satisfaction"        ,       "health_satisfaction"          ,   "psychological_satisfaction"     ,
              "pain"                      ,      "predominent_area_of_pain"     ,   "rating_of_pain"                 ,
              "autonomic_dysfunction"      ,     "weakness"                     ,   "predominant_area_of_weakness"   ,
              "stiffness"                   ,    "predominant_area_of_stiffness",   "dress"                          ,
              "mobilise"                     ,   "work"                         ,   "physical_activities"            ,
              "holding_things"                ,  "cleaning_self"                ,   "eat_self"                       ,
              "write_use_phone"                , "socialise"                    ,   "toilet")


df_freq <- create_frequency_tables(df, vars_cat)
colnames(df)

vars_num <- c("age","life_satisfaction"        ,       "health_satisfaction"          ,   "psychological_satisfaction"     ,
              "rating_of_pain")

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    Median = numeric(),
    SEM = numeric(),  # Standard Error of the Mean
    SD = numeric(),   # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    median_val <- median(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      Median = median_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

# Example usage with your dataframe and list of variables
df_descriptive_stats <- calculate_descriptive_stats(df, vars_num)


# Recoding variables
df <- df %>%
  mutate(
    # Follow-up: Group rare values, code very low occurrences as NA
    follow_up = case_when(
      follow_up %in% c("2021", "self") ~ "Other",
      follow_up == "yes" ~ "Yes",
      follow_up == "no" ~ "No",
      TRUE ~ NA_character_
    ),
    
    
    # Level of Injury: Grouping rare categories, marking very low as missing
    level_of_injury = case_when(
      level_of_injury %in% c("ct", "tl") ~ NA_character_,
      level_of_injury %in% c("c", "l", "t") ~ as.character(level_of_injury),
      TRUE ~ as.character(level_of_injury)
    ),
    
    # Mechanism of Injury (MOI): Group rare values and code low observations as NA
    moi = case_when(
      moi %in% c("gunshot wound", "pedestrian vs auto", "fall") ~ "Other",
      moi %in% c("blunt trauma", "fall from height", "motor vehicle accident", "stab") ~ moi,
      TRUE ~ NA_character_
    ),
    
    
    
    # Type of Treatment: Grouping minor categories
    type_of_treatment = case_when(
      type_of_treatment == "follow-up" ~ NA_character_,
      type_of_treatment %in% c("home therapy", "none", "on treatment") ~ type_of_treatment,
      TRUE ~ NA_character_
    ),
    
    # Recovery at Discharge: Grouping categories
    recovery_discharge = case_when(
      recovery_discharge == "worsened" ~ NA_character_,
      TRUE ~ recovery_discharge
    ),
    
    # Rehabilitation Discipline: Group rare categories, set low values to NA
    rehab_discipline = case_when(
      rehab_discipline == "occupational" ~ NA_character_,
      rehab_discipline == "physiotherapy" ~ rehab_discipline,
      TRUE ~ NA_character_
    ),
    
    # Speaker: Combine categories where applicable
    speaker = case_when(
      speaker == "caregiver" ~ "caregiver",
      speaker == "patient" ~ "Patient",
      TRUE ~ NA_character_
    ),
    
    # Pain: No grouping required, handling small values
    pain = ifelse(pain %in% c("yes", "no"), pain, NA_character_),
    
    # Socialization: No grouping required, handling small values
    socialise = ifelse(socialise %in% c("yes", "no"), socialise, NA_character_),
    
    # Work: No grouping required, handling small values
    work = ifelse(work %in% c("yes", "no"), work, NA_character_),
    
    # Mobility: No grouping required, handling small values
    mobilise = ifelse(mobilise %in% c("yes", "no"), mobilise, NA_character_),
    
    # Weakness: No grouping required, handling small values
    weakness = ifelse(weakness %in% c("yes", "no"), weakness, NA_character_),
    
    # Last GP Review: Combining years
    last_gp_review = case_when(
      last_gp_review %in% c("2018", "2022", "2023", "2024") ~ last_gp_review,
      last_gp_review == "ongoing" ~ "Ongoing",
      TRUE ~ NA_character_
    ),
    
    # Predominant Area of Weakness: Grouping low values, marking as missing
    predominant_area_of_weakness = case_when(
      predominant_area_of_weakness %in% c("forearm", "hip", "knee", "shoulder") ~ predominant_area_of_weakness,
      TRUE ~ NA_character_
    ),
    
    # Predominant Area of Stiffness: Grouping low values, marking as missing
    predominant_area_of_stiffness = case_when(
      predominant_area_of_stiffness %in% c("ankle", "hip", "knee", "shoulder") ~ predominant_area_of_stiffness,
      TRUE ~ NA_character_
    ),
    
    # Dress: No grouping required, handling small values
    dress = ifelse(dress %in% c("yes", "no"), dress, NA_character_),
    
    # Holding Things: No grouping required, handling small values
    holding_things = ifelse(holding_things %in% c("yes", "no"), holding_things, NA_character_),
    
    # Toilet: No grouping required, handling small values
    toilet = ifelse(toilet %in% c("yes", "no"), toilet, NA_character_)
  )

df_freq2 <- create_frequency_tables(df, vars_cat)

# Hypotheses Tests

# Calculate mean, median, and standard deviation of age
hypothesis_1 <- df %>%
  summarise(
    Mean_Age = mean(age, na.rm = TRUE),
    Median_Age = median(age, na.rm = TRUE),
    SD_Age = sd(age, na.rm = TRUE)
  )

# Display the results data frame
print(hypothesis_1)


# Calculate group-wise mean, median, and SD
age_gender_stats <- df %>%
  group_by(gender) %>%
  summarise(
    Mean_Age = mean(age, na.rm = TRUE),
    Median_Age = median(age, na.rm = TRUE),
    SD_Age = sd(age, na.rm = TRUE)
  )

# Perform t-test on age by gender
t_test_result <- t.test(age ~ gender, data = df)

# Save t-test results into a data frame
hypothesis_2 <- data.frame(
  Gender = age_gender_stats$gender,
  Mean_Age = age_gender_stats$Mean_Age,
  Median_Age = age_gender_stats$Median_Age,
  SD_Age = age_gender_stats$SD_Age,
  T_Statistic = t_test_result$statistic,
  P_Value = t_test_result$p.value
)

# Display the results data frame
print(hypothesis_2)

# Perform Kruskal-Wallis test for age across mechanisms of injury
kruskal_test_moi <- kruskal.test(age ~ moi, data = df)


# Chi-square test for mechanism of injury by gender
moi_gender_table <- table(df$gender, df$moi)
moi_gender_chisq <- chisq.test(moi_gender_table)

# Convert the table to a data frame and calculate proportions
hypothesis_5 <- as.data.frame(moi_gender_table)
names(hypothesis_5) <- c("Gender", "Mechanism_of_Injury", "Frequency")

# Calculate proportions by gender and by mechanism of injury
hypothesis_5$Proportion_of_Gender <- hypothesis_5$Frequency / rowSums(moi_gender_table)[hypothesis_5$Gender] * 100
hypothesis_5$Proportion_of_Mechanism <- hypothesis_5$Frequency / colSums(moi_gender_table)[hypothesis_5$Mechanism_of_Injury] * 100

# Add test statistics
hypothesis_5$Chi_Square_Statistic <- moi_gender_chisq$statistic
hypothesis_5$P_Value <- moi_gender_chisq$p.value

# Display the results data frame
print(hypothesis_5)

# Chi-square test for level of injury by gender
loi_gender_table <- table(df$gender, df$level_of_injury)
loi_gender_chisq <- chisq.test(loi_gender_table)

# Convert the table to a data frame and calculate proportions
hypothesis_6 <- as.data.frame(loi_gender_table)
names(hypothesis_6) <- c("Gender", "Level_of_Injury", "Frequency")

# Calculate proportions by gender and by level of injury
hypothesis_6$Proportion_of_Gender <- hypothesis_6$Frequency / rowSums(loi_gender_table)[hypothesis_6$Gender] * 100
hypothesis_6$Proportion_of_Level <- hypothesis_6$Frequency / colSums(loi_gender_table)[hypothesis_6$Level_of_Injury] * 100

# Add test statistics
hypothesis_6$Chi_Square_Statistic <- loi_gender_chisq$statistic
hypothesis_6$P_Value <- loi_gender_chisq$p.value

# Display the results data frame
print(hypothesis_6)

# Chi-square test for ASIA scale by gender
asia_gender_table <- table(df$gender, df$asia_score)
asia_gender_chisq <- chisq.test(asia_gender_table)

# Convert the table to a data frame and calculate proportions
hypothesis_7 <- as.data.frame(asia_gender_table)
names(hypothesis_7) <- c("Gender", "ASIA_Score", "Frequency")

# Calculate proportions by gender and by ASIA score
hypothesis_7$Proportion_of_Gender <- hypothesis_7$Frequency / rowSums(asia_gender_table)[hypothesis_7$Gender] * 100
hypothesis_7$Proportion_of_ASIA <- hypothesis_7$Frequency / colSums(asia_gender_table)[hypothesis_7$ASIA_Score] * 100

# Add test statistics
hypothesis_7$Chi_Square_Statistic <- asia_gender_chisq$statistic
hypothesis_7$P_Value <- asia_gender_chisq$p.value

# Display the results data frame
print(hypothesis_7)


# Chi-square test for recovery at discharge by gender
recovery_gender_table <- table(df$gender, df$recovery_discharge)
recovery_gender_chisq <- chisq.test(recovery_gender_table)

# Convert the table to a data frame and calculate proportions
hypothesis_8 <- as.data.frame(recovery_gender_table)
names(hypothesis_8) <- c("Gender", "Recovery_At_Discharge", "Frequency")

# Calculate proportions by gender and by recovery outcome
hypothesis_8$Proportion_of_Gender <- hypothesis_8$Frequency / rowSums(recovery_gender_table)[hypothesis_8$Gender] * 100
hypothesis_8$Proportion_of_Recovery <- hypothesis_8$Frequency / colSums(recovery_gender_table)[hypothesis_8$Recovery_At_Discharge] * 100

# Add test statistics
hypothesis_8$Chi_Square_Statistic <- recovery_gender_chisq$statistic
hypothesis_8$P_Value <- recovery_gender_chisq$p.value

# Display the results data frame
print(hypothesis_8)

# Chi-square test for type of surgery by gender
surgery_gender_table <- table(df$gender, df$type_of_treatment)
surgery_gender_chisq <- chisq.test(surgery_gender_table)

# Convert the table to a data frame and calculate proportions
hypothesis_9 <- as.data.frame(surgery_gender_table)
names(hypothesis_9) <- c("Gender", "Type_of_Surgery", "Frequency")

# Calculate proportions by gender and by type of surgery
hypothesis_9$Proportion_of_Gender <- hypothesis_9$Frequency / rowSums(surgery_gender_table)[hypothesis_9$Gender] * 100
hypothesis_9$Proportion_of_Surgery <- hypothesis_9$Frequency / colSums(surgery_gender_table)[hypothesis_9$Type_of_Surgery] * 100

# Add test statistics
hypothesis_9$Chi_Square_Statistic <- surgery_gender_chisq$statistic
hypothesis_9$P_Value <- surgery_gender_chisq$p.value

# Display the results data frame
print(hypothesis_9)

# Ensure the 'doi' column is numeric; coerce non-numeric strings to NA
df$doi <- as.numeric(as.character(df$doi))

# Handle potential NA values that result from non-numeric strings
if (any(is.na(df$doi))) {
  warning("NA introduced by coercion in 'doa'. Some values were not numeric and have been set to NA.")
}

# Now convert the numeric values to dates for 'doa'
df$doi <- as.Date(df$doi, origin = "1899-12-30")


# Ensure the 'doa' column is numeric; coerce non-numeric strings to NA
df$doa <- as.numeric(as.character(df$doa))

# Handle potential NA values that result from non-numeric strings
if (any(is.na(df$doa))) {
  warning("NA introduced by coercion in 'doa'. Some values were not numeric and have been set to NA.")
}

# Now convert the numeric values to dates for 'doa'
df$doa <- as.Date(df$doa, origin = "1899-12-30")

# Display any rows where NA was introduced in 'doa', if you want to inspect them
df[is.na(df$doa), ]

# Ensure the 'doo' column is numeric; coerce non-numeric strings to NA
df$doo <- as.numeric(as.character(df$doo))

# Handle potential NA values that result from non-numeric strings
if (any(is.na(df$doo))) {
  warning("NA introduced by coercion in 'doo'. Some values were not numeric and have been set to NA.")
}

# Now convert the numeric values to dates for 'doo'
df$doo <- as.Date(df$doo, origin = "1899-12-30")

# Display any rows where NA was introduced in 'doo', if you want to inspect them
df[is.na(df$doo), ]


# Calculating time intervals
df$injury_to_admission <- as.numeric(df$doa - df$doi)
df$admission_to_treatment <- as.numeric(df$doo - df$doa)
df$injury_to_treatment <- as.numeric(df$doo - df$doi)

# Calculate summary statistics for time intervals
hypothesis_10 <- df %>%
  summarise(
    Mean_injury_to_admission = mean(injury_to_admission, na.rm = TRUE),
    SD_injury_to_admission = sd(injury_to_admission, na.rm = TRUE),
    Median_injury_to_admission = median(injury_to_admission, na.rm = TRUE),
    
    Mean_admission_to_treatment = mean(admission_to_treatment, na.rm = TRUE),
    SD_admission_to_treatment = sd(admission_to_treatment, na.rm = TRUE),
    Median_admission_to_treatment = median(admission_to_treatment, na.rm = TRUE),
    
    Mean_injury_to_treatment = mean(injury_to_treatment, na.rm = TRUE),
    SD_injury_to_treatment = sd(injury_to_treatment, na.rm = TRUE),
    Median_injury_to_treatment = median(injury_to_treatment, na.rm = TRUE)
  )

# Display the results
print(hypothesis_10)

hypothesis_11 <- df %>%
  group_by(moi) %>%
  summarise(
    Mean_injury_to_admission = mean(injury_to_admission, na.rm = TRUE),
    SD_injury_to_admission = sd(injury_to_admission, na.rm = TRUE),
    Median_injury_to_admission = median(injury_to_admission, na.rm = TRUE),
    
    Mean_admission_to_treatment = mean(admission_to_treatment, na.rm = TRUE),
    SD_admission_to_treatment = sd(admission_to_treatment, na.rm = TRUE),
    Median_admission_to_treatment = median(admission_to_treatment, na.rm = TRUE),
    
    Mean_injury_to_treatment = mean(injury_to_treatment, na.rm = TRUE),
    SD_injury_to_treatment = sd(injury_to_treatment, na.rm = TRUE),
    Median_injury_to_treatment = median(injury_to_treatment, na.rm = TRUE)
  )

# Display the results
print(hypothesis_11)

hypothesis_12 <- df %>%
  group_by(level_of_injury) %>%
  summarise(
    Mean_injury_to_admission = mean(injury_to_admission, na.rm = TRUE),
    SD_injury_to_admission = sd(injury_to_admission, na.rm = TRUE),
    Median_injury_to_admission = median(injury_to_admission, na.rm = TRUE),
    
    Mean_admission_to_treatment = mean(admission_to_treatment, na.rm = TRUE),
    SD_admission_to_treatment = sd(admission_to_treatment, na.rm = TRUE),
    Median_admission_to_treatment = median(admission_to_treatment, na.rm = TRUE),
    
    Mean_injury_to_treatment = mean(injury_to_treatment, na.rm = TRUE),
    SD_injury_to_treatment = sd(injury_to_treatment, na.rm = TRUE),
    Median_injury_to_treatment = median(injury_to_treatment, na.rm = TRUE)
  )

# Display the results
print(hypothesis_12)


hypothesis_13 <- df %>%
  group_by(asia_score) %>%
  summarise(
    Mean_injury_to_admission = mean(injury_to_admission, na.rm = TRUE),
    SD_injury_to_admission = sd(injury_to_admission, na.rm = TRUE),
    Median_injury_to_admission = median(injury_to_admission, na.rm = TRUE),
    
    Mean_admission_to_treatment = mean(admission_to_treatment, na.rm = TRUE),
    SD_admission_to_treatment = sd(admission_to_treatment, na.rm = TRUE),
    Median_admission_to_treatment = median(admission_to_treatment, na.rm = TRUE),
    
    Mean_injury_to_treatment = mean(injury_to_treatment, na.rm = TRUE),
    SD_injury_to_treatment = sd(injury_to_treatment, na.rm = TRUE),
    Median_injury_to_treatment = median(injury_to_treatment, na.rm = TRUE)
  )

# Display the results
print(hypothesis_13)



# Linear regression of injury to admission time on age
model_age <- lm(injury_to_admission ~ age, data = df)

# Create a dataframe to hold summary results
hypothesis_14 <- data.frame(
  R_Squared = summary(model_age)$r.squared,
  F_Statistic = summary(model_age)$fstatistic[1],
  P_Value = pf(summary(model_age)$fstatistic[1], summary(model_age)$fstatistic[2], summary(model_age)$fstatistic[3], lower.tail = FALSE),
  Beta_Coefficient_Age = coef(model_age)["age"]
)

# Display the results
print(hypothesis_14)

anova_or_kruskal <- function(data, target_variable, group_variable) {
  # Ensure target_variable and group_variable are treated as factors if not already
  data[[group_variable]] <- as.factor(data[[group_variable]])
  
  # Normality and homogeneity tests
  normality_test <- shapiro.test(data[[target_variable]])
  homogeneity_test <- bartlett.test(data[[target_variable]] ~ data[[group_variable]])
  
  # Output results of the tests
  print(list("Normality Test" = normality_test, "Homogeneity Test" = homogeneity_test))
  
  # Select test based on normality and homogeneity outcomes
  if (normality_test$p.value > 0.05 && homogeneity_test$p.value > 0.05) {
    # ANOVA
    test_result <- aov(data[[target_variable]] ~ data[[group_variable]])
    summary_test <- summary(test_result)
    result_df <- data.frame(
      Target_Variable = target_variable,
      Group_Variable = group_variable,
      Test = "ANOVA",
      F_Statistic = summary_test[[1]][["F value"]][1],
      P_Value = summary_test[[1]][["Pr(>F)"]][1],
      Group = levels(data[[group_variable]]),
      Mean = tapply(data[[target_variable]], data[[group_variable]], mean, na.rm = TRUE)
    )
  } else {
    # Kruskal-Wallis
    test_result <- kruskal.test(data[[target_variable]] ~ data[[group_variable]])
    medians <- tapply(data[[target_variable]], data[[group_variable]], median, na.rm = TRUE)
    means <- tapply(data[[target_variable]], data[[group_variable]], mean, na.rm = TRUE)
    
    result_df <- data.frame(
      Target_Variable = target_variable,
      Group_Variable = group_variable,
      Test = "Kruskal-Wallis",
      Kruskal_Statistic = test_result$statistic,
      P_Value = test_result$p.value,
      Group = names(medians),
      Median = as.vector(medians),
      Mean = as.vector(means)
    )
  }
  
  # Return the results dataframe
  return(result_df)
}




# Hypothesis 15: Analyzing time to admission grouped by mechanism of injury
hypothesis_15_results <- anova_or_kruskal(df, "injury_to_admission", "moi")
print(hypothesis_15_results)

# Hypothesis 16: Analyzing time to surgery grouped by level of injury
hypothesis_16_results <- anova_or_kruskal(df, "injury_to_treatment", "level_of_injury")
print(hypothesis_16_results)

# Hypothesis 17: Analyzing age grouped by ASIA Impairment Scale
hypothesis_17_results <- anova_or_kruskal(df, "age", "asia_score")
print(hypothesis_17_results)

# Hypothesis 18: Dependency with Surgery Offered
hypothesis_18_results <- anova_or_kruskal(df, "injury_to_admission", "type_of_treatment")
print(hypothesis_18_results)


# Hypothesis 19: Dependency on mechanism of injury for injury to treatment
hypothesis_19_results <- anova_or_kruskal(df, "injury_to_treatment", "moi")
print("Hypothesis 19 Results:")
print(hypothesis_19_results)

# Hypothesis 20: Dependency on level of injury for injury to treatment
hypothesis_20_results <- anova_or_kruskal(df, "injury_to_treatment", "level_of_injury")
print("Hypothesis 20 Results:")
print(hypothesis_20_results)

# Hypothesis 21: Dependency on sex for injury to treatment
hypothesis_21_results <- anova_or_kruskal(df, "injury_to_treatment", "gender")
print("Hypothesis 21 Results:")
print(hypothesis_21_results)


df_freq_age <- create_frequency_tables(df, "age")

df$age_group <- case_when(
  df$age <= 25 ~ "Young (≤ 25)",
  df$age > 25 & df$age <= 50 ~ "Middle-aged (26-50)",
  df$age > 50 ~ "Older (> 50)",
  TRUE ~ NA_character_
)


# Hypothesis 22: Dependency on age for injury to treatment
hypothesis_22_results <- anova_or_kruskal(df, "injury_to_treatment", "age_group")
print("Hypothesis 22 Results:")
print(hypothesis_22_results)

# Hypothesis 23: Dependency on ASIA Impairment Scale for injury to treatment
hypothesis_23_results <- anova_or_kruskal(df, "injury_to_treatment", "asia_score")
print("Hypothesis 23 Results:")
print(hypothesis_23_results)

# Hypothesis 24: Dependency on age for admission to treatment
hypothesis_24_results <- anova_or_kruskal(df, "admission_to_treatment", "age_group")
print("Hypothesis 24 Results:")
print(hypothesis_24_results)

# Hypothesis 25: Dependency on sex for admission to treatment
hypothesis_25_results <- anova_or_kruskal(df, "admission_to_treatment", "gender")
print("Hypothesis 25 Results:")
print(hypothesis_25_results)

# Hypothesis 26: Dependency on level of injury for admission to treatment
hypothesis_26_results <- anova_or_kruskal(df, "admission_to_treatment", "level_of_injury")
print("Hypothesis 26 Results:")
print(hypothesis_26_results)

# Hypothesis 27: Dependency on mechanism of injury for admission to treatment
hypothesis_27_results <- anova_or_kruskal(df, "admission_to_treatment", "moi")
print("Hypothesis 27 Results:")
print(hypothesis_27_results)

# Hypothesis 28: Dependency on ASIA Impairment Scale for admission to treatment
hypothesis_28_results <- anova_or_kruskal(df, "admission_to_treatment", "asia_score")
print("Hypothesis 28 Results:")
print(hypothesis_28_results)

# Hypothesis 29: Relationship between injury to admission and recovery
hypothesis_29_results <- anova_or_kruskal(df, "injury_to_admission", "recovery_discharge")
print("Hypothesis 29 Results:")
print(hypothesis_29_results)

# Hypothesis 30: Relationship between injury to admission and spine functionality
#hypothesis_30_results <- anova_or_kruskal(df, "injury_to_admission", "spine_functionality")
#print("Hypothesis 30 Results:")
#print(hypothesis_30_results)

# Hypothesis 31: Relationship between injury to admission and quality of life
#hypothesis_31_results <- anova_or_kruskal(df, "injury_to_admission", "quality_of_life")
#print("Hypothesis 31 Results:")
#print(hypothesis_31_results)

# Hypothesis 32: Relationship between injury to admission and ADLs
#hypothesis_32_results <- anova_or_kruskal(df, "injury_to_admission", "ADLs")
#print("Hypothesis 32 Results:")
#print(hypothesis_32_results)

# Hypothesis 33: Impact of injury to treatment on spine functionality
#hypothesis_33_results <- anova_or_kruskal(df, "injury_to_treatment", "spine_functionality")
#print("Hypothesis 33 Results:")
#print(hypothesis_33_results)

# Hypothesis 34: Impact of injury to treatment on quality of life
#hypothesis_34_results <- anova_or_kruskal(df, "time_to_treatment", "quality_of_life")
#print("Hypothesis 34 Results:")
#print(hypothesis_34_results)

# Hypothesis 35: Impact of injury to treatment on ADLs
#hypothesis_35_results <- anova_or_kruskal(df, "time_to_treatment", "ADLs")
#print("Hypothesis 35 Results:")
#print(hypothesis_35_results)

# Hypothesis 36: Impact of mechanism of injury on quality of life, spine functionality, and ADLs
#hypothesis_36_results_qol <- anova_or_kruskal(df, "quality_of_life", "moi")
#hypothesis_36_results_spine <- anova_or_kruskal(df, "spine_functionality", "moi")
#hypothesis_36_results_adl <- anova_or_kruskal(df, "ADLs", "moi")
#print("Hypothesis 36 Results - Quality of Life:")
#print(hypothesis_36_results_qol)
#print("Hypothesis 36 Results - Spine Functionality:")
#print(hypothesis_36_results_spine)
#print("Hypothesis 36 Results - ADLs:")
#print(hypothesis_36_results_adl)

# Hypothesis 37: Impact of level of injury on quality of life, spine functionality, and ADLs
#hypothesis_37_results_qol <- anova_or_kruskal(df, "quality_of_life", "level_of_injury")
#hypothesis_37_results_spine <- anova_or_kruskal(df, "spine_functionality", "level_of_injury")
#hypothesis_37_results_adl <- anova_or_kruskal(df, "ADLs", "level_of_injury")
print("Hypothesis 37 Results - Quality of Life:")
#print(hypothesis_37_results_qol)
print("Hypothesis 37 Results - Spine Functionality:")
#print(hypothesis_37_results_spine)
print("Hypothesis 37 Results - ADLs:")
#print(hypothesis_37_results_adl)

# Hypothesis 38: Impact of the ASIA Impairment Scale on quality of life, spine functionality, and ADLs
#hypothesis_38_results_qol <- anova_or_kruskal(df, "quality_of_life", "asia_score")
#hypothesis_38_results_spine <- anova_or_kruskal(df, "spine_functionality", "asia_score")
#hypothesis_38_results_adl <- anova_or_kruskal(df, "ADLs", "asia_score")
print("Hypothesis 38 Results - Quality of Life:")
#print(hypothesis_38_results_qol)
print("Hypothesis 38 Results - Spine Functionality:")
#print(hypothesis_38_results_spine)
print("Hypothesis 38 Results - ADLs:")
#print(hypothesis_38_results_adl)

# Hypothesis 39: Does the ASIA Impairment Scale determine follow-up?
followup_asia_table <- table(df$asia_score, df$follow_up)
followup_asia_chisq <- chisq.test(followup_asia_table)

# Convert the table to a data frame and calculate proportions
hypothesis_39 <- as.data.frame(followup_asia_table)
names(hypothesis_39) <- c("ASIA_Score", "Follow_Up", "Frequency")

# Calculate proportions by ASIA score and follow-up status
hypothesis_39$Proportion_of_ASIA <- hypothesis_39$Frequency / rowSums(followup_asia_table)[hypothesis_39$ASIA_Score] * 100
hypothesis_39$Proportion_of_FollowUp <- hypothesis_39$Frequency / colSums(followup_asia_table)[hypothesis_39$Follow_Up] * 100

# Add test statistics
hypothesis_39$Chi_Square_Statistic <- followup_asia_chisq$statistic
hypothesis_39$P_Value <- followup_asia_chisq$p.value

# Display the results
print("Hypothesis 39 Results - ASIA Impairment Scale and Follow-up:")
print(hypothesis_39)


# Hypothesis 40a: Is level of injury associated with a particular surgery?
surgery_loi_table <- table(df$level_of_injury, df$type_of_treatment)
surgery_loi_chisq <- chisq.test(surgery_loi_table)

# Convert the table to a data frame and calculate proportions
hypothesis_40a <- as.data.frame(surgery_loi_table)
names(hypothesis_40a) <- c("Level_of_Injury", "Type_of_Treatment", "Frequency")

# Calculate proportions by level of injury and type of treatment
hypothesis_40a$Proportion_of_Injury <- hypothesis_40a$Frequency / rowSums(surgery_loi_table)[hypothesis_40a$Level_of_Injury] * 100
hypothesis_40a$Proportion_of_Treatment <- hypothesis_40a$Frequency / colSums(surgery_loi_table)[hypothesis_40a$Type_of_Treatment] * 100

# Add test statistics
hypothesis_40a$Chi_Square_Statistic <- surgery_loi_chisq$statistic
hypothesis_40a$P_Value <- surgery_loi_chisq$p.value

# Display the results
print("Hypothesis 40a Results - Level of Injury and Surgery:")
print(hypothesis_40a)


# Hypothesis 40b: Is ASIA Impairment Scale associated with a particular surgery?
surgery_asia_table <- table(df$asia_score, df$type_of_treatment)
surgery_asia_chisq <- chisq.test(surgery_asia_table)

# Convert the table to a data frame and calculate proportions
hypothesis_40b <- as.data.frame(surgery_asia_table)
names(hypothesis_40b) <- c("ASIA_Score", "Type_of_Treatment", "Frequency")

# Calculate proportions by ASIA score and type of treatment
hypothesis_40b$Proportion_of_ASIA <- hypothesis_40b$Frequency / rowSums(surgery_asia_table)[hypothesis_40b$ASIA_Score] * 100
hypothesis_40b$Proportion_of_Treatment <- hypothesis_40b$Frequency / colSums(surgery_asia_table)[hypothesis_40b$Type_of_Treatment] * 100

# Add test statistics
hypothesis_40b$Chi_Square_Statistic <- surgery_asia_chisq$statistic
hypothesis_40b$P_Value <- surgery_asia_chisq$p.value

# Display the results
print("Hypothesis 40b Results - ASIA Impairment Scale and Surgery:")
print(hypothesis_40b)


# Hypothesis 41a: Level of injury and average duration from injury to admission
hypothesis_41a_results <- anova_or_kruskal(df, "injury_to_admission", "level_of_injury")
print("Hypothesis 41a Results - Level of Injury and Time to Admission:")
print(hypothesis_41a_results)

# Hypothesis 41b: ASIA Impairment Scale and average duration from injury to admission
hypothesis_41b_results <- anova_or_kruskal(df, "injury_to_admission", "asia_score")
print("Hypothesis 41b Results - ASIA Impairment Scale and Time to Admission:")
print(hypothesis_41b_results)

# Hypothesis 41c: Level of injury and average duration from injury to treatment
hypothesis_41c_results <- anova_or_kruskal(df, "admission_to_treatment", "level_of_injury")
print("Hypothesis 41c Results - Level of Injury and Time to Treatment:")
print(hypothesis_41c_results)

# Hypothesis 41d: ASIA Impairment Scale and average duration from injury to treatment
hypothesis_41d_results <- anova_or_kruskal(df, "admission_to_treatment", "asia_score")
print("Hypothesis 41d Results - ASIA Impairment Scale and Time to Treatment:")
print(hypothesis_41d_results)

# Perform the test using the function
hypothesis_4 <- anova_or_kruskal(df, "age", "moi")

# Display the results
print(hypothesis_4)

# Perform the test using the function
hypothesis_3 <- anova_or_kruskal(df, "age", "level_of_injury")

# Display the results
print(hypothesis_3)

# Combine all the hypotheses with 7 columns into a single dataframe
kruskal_combined <- rbind(
  
  hypothesis_15_results,
  hypothesis_16_results,
  hypothesis_17_results,
  hypothesis_18_results,
  hypothesis_19_results,
  hypothesis_20_results,
  hypothesis_21_results,
  hypothesis_22_results,
  hypothesis_23_results,
  hypothesis_24_results,
  hypothesis_25_results,
  hypothesis_26_results,
  hypothesis_27_results,
  hypothesis_28_results,
  hypothesis_29_results,
  
  hypothesis_41a_results,
  hypothesis_41b_results,
  hypothesis_41c_results,
  hypothesis_41d_results
)

# Print the concatenated data frame
print(kruskal_combined)



# Exporting all hypothesis separately excluding Kruskal-Wallis tests
data_list <- list(
  "DF" = df,
  "Freq" = df_freq,
  "Freq2" = df_freq2,
  "Hypothesis_1" = hypothesis_1,
  "Hypothesis_2" = hypothesis_2,
  "Hypothesis_3" = hypothesis_3,
  "Hypothesis_4" = hypothesis_4,
  "Hypothesis_5" = hypothesis_5,
  "Hypothesis_6" = hypothesis_6,
  "Hypothesis_7" = hypothesis_7,
  "Hypothesis_8" = hypothesis_8,
  "Hypothesis_9" = hypothesis_9,
  "Hypothesis_10" = hypothesis_10,
  "Hypothesis_11" = hypothesis_11,
  "Hypothesis_12" = hypothesis_12,
  "Hypothesis_13" = hypothesis_13,
  "Hypothesis_14" = hypothesis_14,
  "Hypothesis_39" = hypothesis_39,
  "Hypothesis_40a" = hypothesis_40a,
  "Hypothesis_40b" = hypothesis_40b,
  "Kruskal_Combined" = kruskal_combined
)

# Function to save the data in APA formatted Excel
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

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")

# Notify that the export is complete
print("Excel file saved as APA_Formatted_Tables2.xlsx")


colnames(df)




# Ensure no missing values in relevant columns
library(dplyr)

# Perform chi-square tests and combine results into a single table
chi_square_results <- list()

# 1. Chi-Square: Gender vs Management
df_filtered <- df %>% filter(!is.na(gender), !is.na(management_categorical))
gender_management_table <- table(df_filtered$gender, df_filtered$management_categorical)
gender_management_chisq <- chisq.test(gender_management_table)

# Convert the table to a data frame and add proportions
gender_management_df <- as.data.frame(gender_management_table)
names(gender_management_df) <- c("Group", "Management", "Frequency")

gender_management_df <- gender_management_df %>%
  mutate(
    Proportion_of_Group = Frequency / rowSums(gender_management_table)[Group] * 100,
    Proportion_of_Management = Frequency / colSums(gender_management_table)[Management] * 100,
    Test = "Chi-Square",
    Variable = "Gender",
    Statistic = gender_management_chisq$statistic,
    P_Value = gender_management_chisq$p.value
  )

chi_square_results[[1]] <- gender_management_df

# 2. Chi-Square: Level of Injury vs Management
df_filtered <- df %>% filter(!is.na(level_of_injury), !is.na(management_categorical))
loi_management_table <- table(df_filtered$level_of_injury, df_filtered$management_categorical)
loi_management_chisq <- chisq.test(loi_management_table)

loi_management_df <- as.data.frame(loi_management_table)
names(loi_management_df) <- c("Group", "Management", "Frequency")

loi_management_df <- loi_management_df %>%
  mutate(
    Proportion_of_Group = Frequency / rowSums(loi_management_table)[Group] * 100,
    Proportion_of_Management = Frequency / colSums(loi_management_table)[Management] * 100,
    Test = "Chi-Square",
    Variable = "Level of Injury",
    Statistic = loi_management_chisq$statistic,
    P_Value = loi_management_chisq$p.value
  )

chi_square_results[[2]] <- loi_management_df

# 3. Chi-Square: ASIA Score vs Management
df_filtered <- df %>% filter(!is.na(asia_score), !is.na(management_categorical))
asia_management_table <- table(df_filtered$asia_score, df_filtered$management_categorical)
asia_management_chisq <- chisq.test(asia_management_table)

asia_management_df <- as.data.frame(asia_management_table)
names(asia_management_df) <- c("Group", "Management", "Frequency")

asia_management_df <- asia_management_df %>%
  mutate(
    Proportion_of_Group = Frequency / rowSums(asia_management_table)[Group] * 100,
    Proportion_of_Management = Frequency / colSums(asia_management_table)[Management] * 100,
    Test = "Chi-Square",
    Variable = "ASIA Score",
    Statistic = asia_management_chisq$statistic,
    P_Value = asia_management_chisq$p.value
  )

chi_square_results[[3]] <- asia_management_df

# 4. Chi-Square: Mechanism of Injury vs Management
df_filtered <- df %>% filter(!is.na(moi), !is.na(management_categorical))
moi_management_table <- table(df_filtered$moi, df_filtered$management_categorical)
moi_management_chisq <- chisq.test(moi_management_table)

moi_management_df <- as.data.frame(moi_management_table)
names(moi_management_df) <- c("Group", "Management", "Frequency")

moi_management_df <- moi_management_df %>%
  mutate(
    Proportion_of_Group = Frequency / rowSums(moi_management_table)[Group] * 100,
    Proportion_of_Management = Frequency / colSums(moi_management_table)[Management] * 100,
    Test = "Chi-Square",
    Variable = "Mechanism of Injury",
    Statistic = moi_management_chisq$statistic,
    P_Value = moi_management_chisq$p.value
  )

chi_square_results[[4]] <- moi_management_df

# Combine chi-square results
chi_square_combined <- do.call(rbind, chi_square_results)

# Perform Kruskal-Wallis tests and combine results into a single table
kruskal_results <- list()

# 1. Kruskal-Wallis: Injury to Admission vs Management
df_filtered <- df %>% filter(!is.na(injury_to_admission), !is.na(management_categorical))
kw_injury_admission <- kruskal.test(injury_to_admission ~ management_categorical, data = df_filtered)

kruskal_results[[1]] <- df_filtered %>%
  group_by(management_categorical) %>%
  summarise(
    Test = "Kruskal-Wallis",
    Target_Variable = "Injury to Admission",
    Median = median(injury_to_admission, na.rm = TRUE),
    Mean = mean(injury_to_admission, na.rm = TRUE),
    Count = n()
  ) %>%
  mutate(
    Statistic = kw_injury_admission$statistic,
    P_Value = kw_injury_admission$p.value
  )

# 2. Kruskal-Wallis: Injury to Treatment vs Management
df_filtered <- df %>% filter(!is.na(injury_to_treatment), !is.na(management_categorical))
kw_injury_treatment <- kruskal.test(injury_to_treatment ~ management_categorical, data = df_filtered)

kruskal_results[[2]] <- df_filtered %>%
  group_by(management_categorical) %>%
  summarise(
    Test = "Kruskal-Wallis",
    Target_Variable = "Injury to Treatment",
    Median = median(injury_to_treatment, na.rm = TRUE),
    Mean = mean(injury_to_treatment, na.rm = TRUE),
    Count = n()
  ) %>%
  mutate(
    Statistic = kw_injury_treatment$statistic,
    P_Value = kw_injury_treatment$p.value
  )

# Combine Kruskal-Wallis results
kruskal_combined <- do.call(rbind, kruskal_results)

# Display the results
print("Chi-Square Combined Results")
print(chi_square_combined)

print("Kruskal-Wallis Combined Results")
print(kruskal_combined)

# Exporting all hypothesis separately excluding Kruskal-Wallis tests
data_list <- list(
  "Management ChiSq" = chi_square_combined,
  "Management Kruskal" = kruskal_combined
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Management_Results.xlsx")
