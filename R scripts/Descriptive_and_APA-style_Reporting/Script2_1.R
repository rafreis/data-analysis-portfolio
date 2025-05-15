setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/chippie77")

library(openxlsx)
df <- read.xlsx("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/chippie77/Data3.xlsx")

#Clean Names
names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))

colnames(df)

#Explore duplicate columns

# Find duplicate column names
duplicated_columns <- names(df)[duplicated(names(df))]

# Print out the duplicate column names
print(duplicated_columns)

# Function to rename duplicate columns
rename_duplicates <- function(df) {
  names(df) <- make.unique(names(df), sep = "_")
  return(df)
}

# Apply the function to your dataframe
df <- rename_duplicates(df)

#Identify Studies
library(dplyr)
library(tidyverse)

df <- df %>%
  mutate(
    DiPa = if_else(str_detect(Sample.Set, "2"), "DiPa", NA_character_),  # Creates a DiPa column
    DiGa = if_else(str_detect(Sample.Set, "1"), "DiGa", NA_character_)   # Creates a DiGa column
  )

# Select releavnt pairs of No/Yes Adele users

library(dplyr)

select_adele_cases <- function(data, patient_id_col, date_col) {
  data %>%
    # Ensure data is sorted by patient and within patient by date of interview
    arrange(!!sym(patient_id_col), !!sym(date_col)) %>%
    # Group by patient
    group_by(!!sym(patient_id_col)) %>%
    # Select the first two rows for each patient
    slice_head(n = 2) %>%
    # Ungroup data
    ungroup()
}

df <- select_adele_cases(df, "Patient", "Date.of.interview")

# Classify patients as control or treatment groups

classify_patients <- function(data, patient_id_col, interview_col, adele_col) {
  data %>%
    # Sort by patient ID and interview
    arrange(!!sym(patient_id_col), !!sym(interview_col)) %>%
    # Group by patient ID
    group_by(!!sym(patient_id_col)) %>%
    # Create a classification for each patient
    mutate(
      Classification = case_when(
        # Check if there's a "no" followed by a "yes" at any point
        any(lag(!!sym(adele_col), 1, default = NA) == "no" & !!sym(adele_col) == "yes") ~ "Treatment",
        # Check if all responses are "no"
        all(!!sym(adele_col) == "no") ~ "Control",
        # Default case
        TRUE ~ NA_character_
      ),
      # Add the Pre_Post column
      Pre_Post = case_when(
        row_number() == 1 ~ "Pre",
        row_number() == 2 ~ "Post",
        TRUE ~ NA_character_
      )
    ) %>%
    # Ungroup data
    ungroup()
}

# Apply the function to your dataframe
df <- classify_patients(df, "Patient", "Interview", "ADELE.user.Yes_No")



# Clean the 6MTW scale

library(dplyr)

df <- df %>%
  mutate(`6.min.walk.test.distance.in.meter` = ifelse(`6.min.walk.test._6_MWT_` == "no", NA, `6.min.walk.test.distance.in.meter`))


# Define scales

scales_diga <- c("6.min.walk.test.distance.in.meter", 'HBP_MEAN', 'PCS_12',	'MCS_12', 'HCS_MEAN', 'DIGA_HELM_Evaluation_Overall', 'HLS_SUM',
                 'development_1_Mean', 'development_2_Mean', 'development_3_Mean', 'development_4_Mean',
                 'development_5_Mean',
                 "Blood.pressure._.at.rest._systolic",                                                                                                                                                                                                                                                        
                 "Blood.pressure._.at.rest._diastolic_")

scales_dipa <- c("6.min.walk.test.distance.in.meter", 'PCS_12',	'MCS_12', 'HCS_MEAN', 'HLS_SUM',
                 'development_1_Mean', 'development_2_Mean', 'development_3_Mean', 'development_4_Mean',
                 'development_5_Mean',
                 "Blood.pressure._.at.rest._systolic",                                                                                                                                                                                                                                                        
                 "Blood.pressure._.at.rest._diastolic_")

# Reliability and Validity tests for caregiving scales

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SE_of_the_Mean = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Added for item-total correlation
    Alpha = numeric(),
    N = numeric(),   # New column to store sample size
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val = alpha_results$total$raw_alpha
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"] # Get ITC for the item
      n <- sum(!is.na(item_data))  # Calculate the non-missing count
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = mean(item_data, na.rm = TRUE),
        SE_of_the_Mean = sd(item_data, na.rm = TRUE) / sqrt(n),
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,
        Alpha = NA,
        N = n  # Include the sample size
      ))
    }
    
    # Calculate the mean score for the scale and add as a new column
    scale_mean <- rowMeans(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_mean
    n_scale <- sum(!is.na(scale_mean))  # Calculate the non-missing count for the scale
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = mean(scale_mean, na.rm = TRUE),
      SE_of_the_Mean = sd(scale_mean, na.rm = TRUE) / sqrt(n_scale),
      StDev = sd(scale_mean, na.rm = TRUE),
      ITC = NA,  # ITC doesn't apply to the whole scale
      Alpha = alpha_val,
      N = n_scale  # Include the sample size for the scale
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}


scales <- list(
  "development_1" = c("1......your.overall.state.of.health",
                      "2.....your.confidence.in.coping.independently.with.everyday.life",
                      "3......your.confidence.in.dealing.with.your.chronic.illness",
                      "4.....taking.medication.regularly",
                      "5.....keeping.a.heart_blood.pressure.diary",
                      "6......your.confidence.in.communicating.with.your.doctor",
                      "7......daily.exercise",
                      "8......regular.drinking",
                      "9......a.regular.daily.routine.and.time.orientation._e.g..eating_.sleeping_",
                      "10......dizziness.and.delirium",
                      "11......risk.of.falling",
                      "12......loss.of.local.orientation",
                      "13......the.communication.between.you.and.your.environment",
                      "14......your.ability.to.keep.appointments",
                      "15......your.subjective.feeling.of.safety.at.home",
                      "16......your.subjective.feeling.of.safety.when.you.are.out.and.about",
                      "17......your.self_confidence",
                      "18......your.fear.of.unforeseen.health.events",
                      "19......your.ability.to.live.independently.at.home"),
  "development_2" = c("20..Quality.of.life.and.health:.Maintaining.your.independence.and.health.situation.by.monitoring.vital.parameters._e.g..heart_blood.pressure_weight.diary_.and.prevention",
                      "21..Mobility.Safety.and.protection:.Maintaining.mobility_.safety.and.assistance.in.the.event.of.incidents.or.abnormalities",
                      "22..Independence.and.self_management:.Maintaining.independence.in.home.life.and.everyday.life",
                      "23..Communication.and.cognition:.Maintaining.cognitive.abilities.and.the.ability.to.communicate.with.their.social.environment",
                      "24..Social.interaction.and.participation:.Maintaining.social.interaction.&.participation.in.public.life.and.in.the.community",
                      "25..Patient.sovereignty.and.information:.More.effective.health.action.through.education_.provision.of.information.or.advice;.maintaining.the.ability.to.self_manage.and.enabling.better.handling.of.ones.own.situation",
                      "26..Adherence:.Maintaining.adherence.to.treatment_.e.g..through.reminders.when.taking.medication",
                      "27..Individualization:.Targeted.care.through.individualization_.e.g..reminders_.motivation.to.exercise_.etc."),
  
  "development_3" = c( "28......the.overall.state.of.health"                                  ,                                                                                                                                                                                                                     
                       "29..….the.confidence.in.coping.independently.with.everyday.life"       ,                                                                                                                                                                                                                    
                       "30......theconfidence.in.dealing.with.your.chronic.illness"             ,                                                                                                                                                                                                                   
                       "31.....taking.medication.regularly"                                      ,                                                                                                                                                                                                                  
                       "32.....keeping.a.heart_blood.pressure.diary"                              ,                                                                                                                                                                                                                
                       "33......the.confidence.in.communicating.with.your.doctor"                   ,                                                                                                                                                                                                               
                       "34......daily.exercise"                                                    ,                                                                                                                                                                                                                
                       "35.....regular.drinking"                                                    ,                                                                                                                                                                                                               
                       "36......a.regular.daily.routine.and.time.orientation._e.g..eating_.sleeping_",                                                                                                                                                                                                              
                       "37......dizziness.and.delirium"                                               ,                                                                                                                                                                                                             
                       "38......risk.of.falling"                                                       ,                                                                                                                                                                                                            
                       "39......loss.of.local.orientation"                                              ,                                                                                                                                                                                                           
                       "40......the.communication.between.you.and.your.environment"                      ,                                                                                                                                                                                                          
                       "41......the.ability.to.keep.appointments"                                         ,                                                                                                                                                                                                         
                       "42......the.subjective.feeling.of.safety.at.home"                                   ,                                                                                                                                                                                                       
                       "43......your.subjective.feeling.of.safety.when.you.are.out.and.about"              ,                                                                                                                                                                                                        
                       "44......your.self_confidence"                                                        ,                                                                                                                                                                                                      
                       "45......your.fear.of.unforeseen.health.events"                                        ,                                                                                                                                                                                                     
                       "46......your.ability.to.live.independently.at.home"),
  "development_4" = c("47..Quality.of.life.and.health:.Maintaining.independence.and.health.situation.by.monitoring.vital.parameters._e.g..heart_blood.pressure_weight.diary_.and.prevention"              ,                                                                                                        
                      "48..Mobility.Safety.and.protection:.Maintaining.mobility_.safety.and.assistance.in.the.event.of.incidents.or.abnormalities"                                                         ,                                                                                                       
                      "49..Independence.and.self_management:.Maintaining.independence.in.home.life.and.everyday.life"                                                                                      ,                                                                                                       
                      "50..Communication.and.cognition:.Maintaining.cognitive.abilities.and.the.ability.to.communicate.with.their.social.environment"                                                       ,                                                                                                      
                      "51..Social.interaction.and.participation:.Maintaining.social.interaction.&.participation.in.public.life.and.in.the.community"                                                         ,                                                                                                     
                      "52..Patient.sovereignty.and.information:.More.effective.health.action.through.education_.provision.of.information.or.advice;.maintaining.the.ability.to.self_manage.and.enabling.better.handling.of.ones.own.situation" ,                                                                   
                      "53..Adherence:.Maintaining.adherence.to.treatment_.e.g..through.reminders.when.taking.medication"      ,                                                                                                                                                                                    
                      "54..Individualization:.Targeted.care.through.individualization_.e.g..reminders_.motivation.to.exercise_.etc." ),
  "development_5" = c("55..Information.about.the.illness.and.the.best.possible.way.to.deal.with.the.care.recipient.Education_.provision.of.information.and.advice.to.caregivers;.enabling.better.handling.of.the.illness.or.care.situation.of.the.person.being.cared.for"  ,                                       
                      "56..Individualization.Targeted.care.for.the.care.recipient.through.individualization.of.care.services_.e.g..reminders_.motivation.to.exercise_.etc." ,                                                                                                                                      
                      "57..Communication.and.participation.Improved.communication.with.the.patient.or.caregiver;.networking.among.caregiving.relatives;.maintaining.relationships"    ,                                                                                                                            
                      "58..Organization.Organization.and.management.of.everyday.care"      ,                                                                                                                                                                                                                       
                      "59..Relief.for.family.caregivers.Reducing.the.physical.and.psychological.stress.caused.by.the.care.situation"  ,                                                                                                                                                                            
                      "60..Quality.of.professional.care.Improving.the.quality.of.care.for.those.in.need.of.care")
)


alpha_results <- reliability_analysis(df, scales)

df_reliability <- alpha_results$statistics


# Validity Analysis

factor_analysis_loadings <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Scale = character(),
    Variable = character(),
    FactorLoading = numeric(),
    N = integer(),  # New column for sample size
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[, scales[[scale]]]
    
    # Remove rows with any missing values in the subset data
    complete_subset_data <- subset_data[complete.cases(subset_data), ]
    
    # Get the number of complete cases
    n_obs <- nrow(complete_subset_data)
    
    # Perform factor analysis; using 1 factor as it's the most common scenario
    fa_results <- factanal(complete_subset_data, factors = 1, rotation = "varimax")
    
    # Extract factor loadings
    loadings <- fa_results$loadings[, 1]
    
    # Append results for each item in the scale
    for (item in scales[[scale]]) {
      results <- rbind(results, data.frame(
        Scale = scale,
        Variable = item,
        FactorLoading = loadings[item],
        N = n_obs  # Add the sample size for the subset
      ))
    }
    
    # Calculate total variance explained
    variance_explained = sum(loadings^2) / length(loadings)
    
    # Print summary of the results - Total Variance Explained
    cat("Scale:", scale, "\n")
    cat("Total Variance Explained by Factor:", variance_explained, "\n")
  }
  
  return(results)
}



df_validity <- factor_analysis_loadings(df, scales)

purified_scales <- list(
  "development_1" = c("1......your.overall.state.of.health",
                      "2.....your.confidence.in.coping.independently.with.everyday.life",
                      "3......your.confidence.in.dealing.with.your.chronic.illness",
                      "4.....taking.medication.regularly",
                      "5.....keeping.a.heart_blood.pressure.diary",
                      "7......daily.exercise",
                      "8......regular.drinking",
                      "9......a.regular.daily.routine.and.time.orientation._e.g..eating_.sleeping_",
                      "10......dizziness.and.delirium",
                      "11......risk.of.falling",
                      "13......the.communication.between.you.and.your.environment",
                      "14......your.ability.to.keep.appointments",
                      "15......your.subjective.feeling.of.safety.at.home",
                      "16......your.subjective.feeling.of.safety.when.you.are.out.and.about",
                      "17......your.self_confidence",
                      "18......your.fear.of.unforeseen.health.events",
                      "19......your.ability.to.live.independently.at.home"),
  "development_2" = c("20..Quality.of.life.and.health:.Maintaining.your.independence.and.health.situation.by.monitoring.vital.parameters._e.g..heart_blood.pressure_weight.diary_.and.prevention",
                      "21..Mobility.Safety.and.protection:.Maintaining.mobility_.safety.and.assistance.in.the.event.of.incidents.or.abnormalities",
                      "22..Independence.and.self_management:.Maintaining.independence.in.home.life.and.everyday.life",
                      "23..Communication.and.cognition:.Maintaining.cognitive.abilities.and.the.ability.to.communicate.with.their.social.environment",
                      "24..Social.interaction.and.participation:.Maintaining.social.interaction.&.participation.in.public.life.and.in.the.community",
                      "25..Patient.sovereignty.and.information:.More.effective.health.action.through.education_.provision.of.information.or.advice;.maintaining.the.ability.to.self_manage.and.enabling.better.handling.of.ones.own.situation",
                      "26..Adherence:.Maintaining.adherence.to.treatment_.e.g..through.reminders.when.taking.medication",
                      "27..Individualization:.Targeted.care.through.individualization_.e.g..reminders_.motivation.to.exercise_.etc."),
  
  "development_3" = c( "28......the.overall.state.of.health"                                  ,                                                                                                                                                                                                                     
                       "29..….the.confidence.in.coping.independently.with.everyday.life"       ,                                                                                                                                                                                                                    
                       "30......theconfidence.in.dealing.with.your.chronic.illness"             ,                                                                                                                                                                                                                   
                       "31.....taking.medication.regularly"                                      ,                                                                                                                                                                                                                  
                       "32.....keeping.a.heart_blood.pressure.diary"                              ,                                                                                                                                                                                                                
                       "33......the.confidence.in.communicating.with.your.doctor"                   ,                                                                                                                                                                                                               
                       "34......daily.exercise"                                                    ,                                                                                                                                                                                                                
                       "35.....regular.drinking"                                                    ,                                                                                                                                                                                                               
                       "36......a.regular.daily.routine.and.time.orientation._e.g..eating_.sleeping_",                                                                                                                                                                                                              
                       "37......dizziness.and.delirium"                                               ,                                                                                                                                                                                                             
                       "38......risk.of.falling"                                                       ,                                                                                                                                                                                                            
                       "39......loss.of.local.orientation"                                              ,                                                                                                                                                                                                           
                       
                       "41......the.ability.to.keep.appointments"                                         ,                                                                                                                                                                                                         
                       "42......the.subjective.feeling.of.safety.at.home"                                   ,                                                                                                                                                                                                       
                       "43......your.subjective.feeling.of.safety.when.you.are.out.and.about"              ,                                                                                                                                                                                                        
                       "44......your.self_confidence"                                                        ,                                                                                                                                                                                                      
                       "45......your.fear.of.unforeseen.health.events"                                        ,                                                                                                                                                                                                     
                       "46......your.ability.to.live.independently.at.home"),
  "development_4" = c("47..Quality.of.life.and.health:.Maintaining.independence.and.health.situation.by.monitoring.vital.parameters._e.g..heart_blood.pressure_weight.diary_.and.prevention"              ,                                                                                                        
                      "48..Mobility.Safety.and.protection:.Maintaining.mobility_.safety.and.assistance.in.the.event.of.incidents.or.abnormalities"                                                         ,                                                                                                       
                      "49..Independence.and.self_management:.Maintaining.independence.in.home.life.and.everyday.life"                                                                                      ,                                                                                                       
                      "50..Communication.and.cognition:.Maintaining.cognitive.abilities.and.the.ability.to.communicate.with.their.social.environment"                                                       ,                                                                                                      
                      
                      "52..Patient.sovereignty.and.information:.More.effective.health.action.through.education_.provision.of.information.or.advice;.maintaining.the.ability.to.self_manage.and.enabling.better.handling.of.ones.own.situation" ,                                                                   
                      "53..Adherence:.Maintaining.adherence.to.treatment_.e.g..through.reminders.when.taking.medication"      ,                                                                                                                                                                                    
                      "54..Individualization:.Targeted.care.through.individualization_.e.g..reminders_.motivation.to.exercise_.etc." ),
  "development_5" = c("55..Information.about.the.illness.and.the.best.possible.way.to.deal.with.the.care.recipient.Education_.provision.of.information.and.advice.to.caregivers;.enabling.better.handling.of.the.illness.or.care.situation.of.the.person.being.cared.for"  ,                                       
                      "56..Individualization.Targeted.care.for.the.care.recipient.through.individualization.of.care.services_.e.g..reminders_.motivation.to.exercise_.etc." ,                                                                                                                                      
                      "57..Communication.and.participation.Improved.communication.with.the.patient.or.caregiver;.networking.among.caregiving.relatives;.maintaining.relationships"    ,                                                                                                                            
                      "58..Organization.Organization.and.management.of.everyday.care"      ,                                                                                                                                                                                                                       
                      "59..Relief.for.family.caregivers.Reducing.the.physical.and.psychological.stress.caused.by.the.care.situation"  ,                                                                                                                                                                            
                      "60..Quality.of.professional.care.Improving.the.quality.of.care.for.those.in.need.of.care")
)

alpha_results <- reliability_analysis(df, purified_scales)

df_reliability2 <- alpha_results$statistics
df_validity2 <- factor_analysis_loadings(df, purified_scales)

#Recalculate means

df_recoded <- alpha_results$data_with_scales

# Splitting the df_recoded dataset by the values in the 'study' column
df_dipa <- filter(df_recoded, !is.na(DiPa))
df_diga <- filter(df_recoded, !is.na(DiGa))

# Update scale names

scales_diga <- c("6.min.walk.test.distance.in.meter", 'HBP_MEAN', 'PCS_12',	'MCS_12', 'HCS_MEAN', 'DIGA_HELM_Evaluation_Overall', 'HLS_SUM',
                 'development_1', 'development_2', 'development_3', 'development_4',
                 'development_5',
                 "Blood.pressure._.at.rest._systolic",                                                                                                                                                                                                                                                        
                 "Blood.pressure._.at.rest._diastolic_")

scales_dipa <- c("6.min.walk.test.distance.in.meter", 'PCS_12',	'MCS_12', 'HCS_MEAN', 'HLS_SUM',
                 'development_1', 'development_2', 'development_3', 'development_4',
                 'development_5',
                 "Blood.pressure._.at.rest._systolic",                                                                                                                                                                                                                                                        
                 "Blood.pressure._.at.rest._diastolic_")

# Sample Characterization

# Inspect Age percentiles

df_diga$Age <- as.numeric(as.character(df_diga$Age))
df_dipa$Age <- as.numeric(as.character(df_dipa$Age))

age_percentiles <- quantile(df_dipa$Age, probs = c(0, 0.20, 0.80, 1), na.rm = TRUE)

# Print the percentiles
print(age_percentiles)

categorize_age <- function(data, age_variable, cutoffs) {
  # Create a temporary copy of the data to handle NAs
  temp_data <- data
  
  # Create a new variable for categorized age
  temp_data$categorized_age <- cut(
    temp_data[[age_variable]],
    breaks = c(-Inf, cutoffs, Inf), 
    labels = c("36-64", "65-79", "80+"),
    include.lowest = TRUE
  )
  
  # Merge the categorized_age column back into the original data
  data$categorized_age <- temp_data$categorized_age
  
  return(data)
}



age_cutoffs <- c(65, 80)
df_diga <- categorize_age(df_diga, "Age", age_cutoffs)

df_dipa <- categorize_age(df_dipa, "Age", age_cutoffs)

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

demographics <- c("categorized_age", "Gender", "Care.Level")
df_freq_diga <- create_frequency_tables(df_diga, demographics)
df_freq_dipa <- create_frequency_tables(df_dipa, demographics)


#  Outlier Evaluation

# Convert specified variables to numeric
for (var in scales_diga) {
  df_diga[[var]] <- as.numeric(as.character(df_diga[[var]]))
}

for (var in scales_dipa) {
  df_dipa[[var]] <- as.numeric(as.character(df_dipa[[var]]))
}

library(dplyr)

calculate_z_scores <- function(data, vars, id_var) {
  # Calculate z-scores for each variable
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~scale(.) %>% as.vector, .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  return(z_score_data)
}

id_var <- "Patient"
z_scores_df_dipa <- calculate_z_scores(df_dipa, scales_dipa, id_var)
z_scores_df_diga <- calculate_z_scores(df_diga, scales_diga, id_var)

# Descriptive Statistics

library(moments)

# Load any required libraries if not already done
# library(e1071)  # For skewness and kurtosis functions

calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(),  # Standard Error of the Mean
    SD = numeric(),   # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    N = numeric(),    # New column to store sample size
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    n_val <- sum(!is.na(variable_data))  # Count of non-missing observations
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(n_val)
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val,
      N = n_val  # Include the sample size
    ))
  }
  
  return(results)
}



df_descriptive_stats_diga <- calculate_descriptive_stats(df_diga, scales_diga)
df_descriptive_stats_dipa <- calculate_descriptive_stats(df_dipa, scales_dipa)

# Normality Assessment

library(e1071)

calculate_stats <- function(data, variables) {
  # Initialize the results dataframe
  results <- data.frame(
    Variable = character(),
    Skewness = numeric(),
    Kurtosis = numeric(),
    Shapiro_Wilk_F = numeric(),
    Shapiro_Wilk_p_value = numeric(),
    N = numeric(),  # New column to store sample size
    stringsAsFactors = FALSE
  )
  
  for (var in variables) {
    if (var %in% names(data)) {
      # Count of non-missing observations
      n_val <- sum(!is.na(data[[var]]))
      
      # Calculate skewness and kurtosis
      skew <- skewness(data[[var]], na.rm = TRUE)
      kurt <- kurtosis(data[[var]], na.rm = TRUE)
      
      # Perform Shapiro-Wilk test, if sufficient data points
      if (n_val >= 3) {  # Shapiro-Wilk requires at least 3 data points
        shapiro_test <- shapiro.test(na.omit(data[[var]]))
        shapiro_stat <- shapiro_test$statistic
        shapiro_p_value <- shapiro_test$p.value
      } else {
        shapiro_stat = NA
        shapiro_p_value = NA
      }
      
      # Add results to the dataframe
      results <- rbind(results, data.frame(
        Variable = var,
        Skewness = skew,
        Kurtosis = kurt,
        Shapiro_Wilk_F = shapiro_stat,
        Shapiro_Wilk_p_value = shapiro_p_value,
        N = n_val  # Include the sample size
      ))
    } else {
      warning(paste("Variable", var, "not found in the data. Skipping."))
    }
  }
  
  return(results)
}


df_normality_results_dipa <- calculate_stats(df_dipa, scales_dipa)
df_normality_results_diga <- calculate_stats(df_diga, scales_diga)

# Descriptive Statistics Disaggregated by Pre/Post

library(dplyr)

calculate_descriptive_stats_bygroups <- function(data, desc_vars, primary_cat_column, secondary_cat_column) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    PrimaryCategory = character(),
    SecondaryCategory = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each combination of categories
  for (var in desc_vars) {
    # Group data by both categorical columns
    grouped_data <- data %>%
      group_by(!!sym(primary_cat_column), !!sym(secondary_cat_column)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE)
      ) %>%
      mutate(Variable = var) %>%
      # Capture both categories for each group
      mutate(PrimaryCategory = !!sym(primary_cat_column), SecondaryCategory = !!sym(secondary_cat_column)) %>%
      # Select and reorder columns for consistency
      select(PrimaryCategory, SecondaryCategory, Variable, Mean, SEM, SD)
    
    # Append the results for each variable and category combination to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}



df_descriptive_stats_bygroup_dipa <- calculate_descriptive_stats_bygroups(df_dipa, scales_dipa,"Classification", "Pre_Post")
df_descriptive_stats_bygroup_diga <- calculate_descriptive_stats_bygroups(df_diga, scales_diga,"Classification", "Pre_Post")


# Dot-and-whiskers to visualize pre/post differences

library(ggplot2)

# Function to create dot-and-whisker plots
create_mean_sd_plot <- function(data, variables, factor_column1, factor_column2, levels_factor1) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(!!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1)) %>%
    group_by(!!sym(factor_column1), !!sym(factor_column2), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      .groups = 'drop'
    )
  
  # Create the plot
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Mean, color = !!sym(factor_column2))) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mean and Standard Deviation Plot", 
      x = "Primary Factor",  
      y = "Value",
      color = "Secondary Factor"
    ) +
    scale_color_discrete(name = "Secondary Factor")
  
  return(p)
}


factor1_levels <- c("Pre", "Post")
factor_column2 <- "Classification"
factor_column1 <- "Pre_Post"
plot <- create_mean_sd_plot(df_diga, scales_diga, factor_column1, factor_column2,factor1_levels)
print(plot)
ggsave("mean_sd_plot_diga.png", plot = plot, width = 12, height = 8)

plot <- create_mean_sd_plot(df_dipa, scales_dipa, factor_column1, factor_column2,factor1_levels)
print(plot)
ggsave("mean_sd_plot_dipa.png", plot = plot, width = 12, height = 8)


# Whiskers as 95% CI

library(ggplot2)
library(dplyr)
library(tidyr)
library(broom)  # For tidy summary statistics

create_mean_ci_plot <- function(data, variables, factor_column1, factor_column2, levels_factor1, legend_title) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(!!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1)) %>%
    group_by(!!sym(factor_column1), !!sym(factor_column2), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      N = n(),
      SD = sd(Value, na.rm = TRUE),
      SEM = SD / sqrt(N),
      .groups = 'drop'
    ) %>%
    mutate(
      CI = qt(0.975, df = N-1) * SEM,
      Lower = Mean - CI,
      Upper = Mean + CI
    )
  
  # Create the plot
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Mean, color = !!sym(factor_column2))) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = Lower, ymax = Upper), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mean and 95% Confidence Interval Plot", 
      x = "Primary Factor",
      y = "Value",
      color = legend_title
    ) +
    scale_color_discrete(name = legend_title)
  
  return(p)
}



plot <- create_mean_ci_plot(df_diga, scales_diga, factor_column1, factor_column2,factor1_levels, "Adele Use")
print(plot)
ggsave("mean_ci_plot_diga.png", plot = plot, width = 12, height = 8)

plot <- create_mean_ci_plot(df_dipa, scales_dipa, factor_column1, factor_column2,factor1_levels, "Adele Use")
print(plot)
ggsave("mean_ci_plot_dipa.png", plot = plot, width = 12, height = 8)

# Linear Mixed Model with Random Effects

library(lme4)
library(broom.mixed)
library(afex)

#Clean Names
names(df_dipa) <- gsub("\\.", "_", names(df_dipa))
names(df_diga) <- gsub("\\.", "_", names(df_diga))
names(df_dipa) <- gsub("6", "six", names(df_dipa))
names(df_diga) <- gsub("6", "six", names(df_diga))

# Declare scales again

colnames(df_diga)

scales_diga <- c("six_min_walk_test_distance_in_meter", 'HBP_MEAN', 'PCS_12',	'MCS_12', 'HCS_MEAN', 'DIGA_HELM_Evaluation_Overall', 'HLS_SUM',
                 'development_1', 'development_2', 'development_3', 'development_4',
                 'development_5',
                 "Blood_pressure___at_rest__systolic"  ,                                                                                                                                                                                                                                                          
                 "Blood_pressure___at_rest__diastolic_")



scales_dipa <- c("six_min_walk_test_distance_in_meter", 'PCS_12',	'MCS_12', 'HCS_MEAN', 'HLS_SUM',
                 'development_1', 'development_2', 'development_3', 'development_4',
                 'development_5',
                 "Blood_pressure___at_rest__systolic"  ,                                                                                                                                                                                                                                                          
                 "Blood_pressure___at_rest__diastolic_")

# Remove Outliers before modelling - Z-score > +-4

df_diga <- df_diga %>%
  mutate(
    HBP_MEAN = if_else(Patient == 62, NA_real_, HBP_MEAN)
  )

# Function to fit LMM for multiple response variables without a between-subjects factor
library(lme4)
library(broom.mixed)  # For tidy output of mixed models

fit_lmm_and_format <- function(data, within_subject_var, between_subject_var, random_effects, response_vars, save_plots = FALSE) {
  lmm_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically with the interaction term
    formula_str <- paste(response_var, "~", within_subject_var, "*", between_subject_var, "+", random_effects)
    formula <- as.formula(formula_str)
    
    # Fit the linear mixed-effects model
    lmer_model <- lmer(formula, data = data)
    
    # Print the summary of the model for fit statistics
    summary_info <- summary(lmer_model)
    print(summary_info)
    
    # Get the sample size for the response variable
    n_obs <- length(fitted(lmer_model))  # Total number of observations used
    
    # Extract the tidy output and add the N column
    lmm_results <- tidy(lmer_model, effects = "fixed") %>%
      mutate(ResponseVariable = response_var,
             N = n_obs)  # Add sample size information
    
    # Optionally print the tidy output
    print(lmm_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lmer_model), main = "Q-Q Plot")
      qqline(residuals(lmer_model))
      dev.off()
    }
    
    # Store the results in a list
    lmm_results_list[[response_var]] <- lmm_results
  }
  
  return(lmm_results_list)
}



# Call the function with the relevant parameters
lmm_results_list_dipa <- fit_lmm_and_format(
  data = df_dipa,
  within_subject_var = "Pre_Post",
  between_subject_var = "Classification",
  random_effects = "(1|Patient)",
  response_vars = scales_dipa,
  save_plots = TRUE
)

# Call the function with the relevant parameters
lmm_results_list_diga <- fit_lmm_and_format(
  data = df_diga,
  within_subject_var = "Pre_Post",
  between_subject_var = "Classification",
  random_effects = "(1|Patient)",
  response_vars = scales_diga,
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults_dipa <- bind_rows(lmm_results_list_dipa)
df_modelresults_diga <- bind_rows(lmm_results_list_diga)


# Controlling by Age and Gender

library(lme4)
library(broom.mixed)  # for tidy()

fit_lmm_and_format_control <- function(data, within_subject_var, between_subject_var, random_effects, response_vars, additional_fixed_effects = NULL, save_plots = FALSE) {
  lmm_results_list <- list()
  
  for (response_var in response_vars) {
    # Start constructing the formula
    formula_string <- paste(response_var, "~", within_subject_var, "*", between_subject_var)
    
    # Add additional fixed effects if provided
    if (!is.null(additional_fixed_effects)) {
      formula_string <- paste(formula_string, "+", additional_fixed_effects)
    }
    
    # Finish constructing the formula by adding random effects
    formula_string <- paste(formula_string, "+", random_effects)
    formula <- as.formula(formula_string)
    
    # Fit the linear mixed-effects model
    lmer_model <- lmer(formula, data = data)
    
    # Calculate the sample size for the model
    n_obs <- length(fitted(lmer_model))  # Total number of observations used
    
    # Extract the tidy output and add the N column
    lmm_results <- tidy(lmer_model, effects = "fixed") %>%
      mutate(ResponseVariable = response_var,
             N = n_obs)  # Include the sample size
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub("[[:punct:] ]", "_", response_var)
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lmer_model))
      qqline(residuals(lmer_model))
      dev.off()
    }
    
    # Store the results in a list
    lmm_results_list[[response_var]] <- lmm_results
  }
  
  return(lmm_results_list)
}


additional_effects <- "categorized_age + Gender"
lmm_results_list_dipa_control <- fit_lmm_and_format_control(
  data = df_dipa,
  within_subject_var = "Pre_Post",
  between_subject_var = "Classification",
  random_effects = "(1|Patient)",
  response_vars = scales_dipa,
  additional_fixed_effects = additional_effects, # Include additional fixed effects
  save_plots = TRUE
)

lmm_results_list_diga_control <- fit_lmm_and_format_control(
  data = df_diga,
  within_subject_var = "Pre_Post",
  between_subject_var = "Classification",
  random_effects = "(1|Patient)",
  response_vars = scales_diga,
  additional_fixed_effects = additional_effects, # Include additional fixed effects
  save_plots = TRUE
)


df_modelresults_dipa_controlled <- bind_rows(lmm_results_list_dipa_control)
df_modelresults_diga_controlled <- bind_rows(lmm_results_list_diga_control)




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
data_list <- list(
  "Reliability - Care" = df_reliability,
  "Reliability - Care2" = df_reliability2,
  "Validity - Care" = df_validity,
  "Validity - Care2" = df_validity,
  "Frequency Table - DiPa" = df_freq_dipa, 
  "Frequency Table - DiGa" = df_freq_diga, 
  "Descriptive Stats - DiPa" = df_descriptive_stats_dipa,
  "Descriptive Stats - DiGa" = df_descriptive_stats_diga,
  "Descriptive PrePost - DiPa" = df_descriptive_stats_bygroup_dipa,
  "Descriptive PrePost - DiGa" = df_descriptive_stats_bygroup_diga,
  "Normality Assessment - DiPa" = df_normality_results_dipa,
  "Normality Assessment - DiGa" = df_normality_results_diga,
  "LMM results - DiPa" = df_modelresults_dipa,
  "LMM results - DiGa" = df_modelresults_diga,
  "LMM results (controlled) - DiPa" = df_modelresults_dipa_controlled,
  "LMM results (controlled) - DiGa" = df_modelresults_diga_controlled,
  "DF" = df_dipa
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_withBloodPressure.xlsx")