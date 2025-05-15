# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/floria")

# Read data from a CSV file
df <- read.csv("results-survey758656(2).csv")

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


#Simplify Column Names

colnames(df) <- c(
  "ResponseID", "DateSubmitted", "LastPage", "StartLanguage", "Seed",
  "DateStarted", "LastActionDate", "JobType", "Citizenship",
  "CitizenshipOther", "EthnicityGermany", "EthnicityEU", "EthnicityNonEU",
  "EthnicityEastAsia", "EthnicitySouthAsia", "EthnicityWestAsia", 
  "EthnicityArabic", "EthnicityAfrica", "EthnicityUSCanada", "EthnicityLatinAmerica",
  "EthnicityOceania", "EthnicityOther", "EthnicityNA", "AgeGroup", "ParentsUni",
  "GenderIdentity", "GenderOther", "GenderDiffBirth", "GenderDiffOther", "LGBTQIA",
  "LGBTQIAOther", "FirstLanguage", "Disability", "DisabilityOther", 
  "Neurodivergent", "NeurodivergentOther", "DisabilityHearing", "DisabilityVision",
  "DisabilityCognitive", "DisabilityCommunicative", "DisabilityMobility",
  "DisabilityMental", "DisabilityUnseen", "DisabilityNA", "DisabilityOtherChallenge",
  "Caretaker", "CaretakerOther", "ComfortSocialClass", "ComfortNationality", 
  "ComfortEthnicity", "ComfortAge", "ComfortSexualOrientation", "ComfortGender", 
  "ComfortPolitics", "ComfortReligion", "ComfortDisability", "ComfortParenthood", 
  "ComfortMentalHealth", "ComfortPregnancy", "FeltDiscriminated", "Belonging", 
  "DiversityPersonal", "DiversityFuture", "ExclusionLanguage", "ComfortOpinions", 
  "OpinionValued", "WorryCommonality", "SelfIdentity", "TeamRespect", "ManagerRespect", 
  "IdeaQuality", "FairTaskDivision", "LeaderDiversityCommitment", 
  "InstituteDiversitySupport", "DiversityPreparedness", "NoDiscrimination", 
  "DiscriminationEducation", "DiscriminationGermanLanguage", "DiscriminationEnglishLanguage",
  "DiscriminationNationality", "DiscriminationEthnicity", "DiscriminationAge",
  "DiscriminationSexualOrientation", "DiscriminationGender", "DiscriminationReligion",
  "DiscriminationPhysical", "DiscriminationParenthood", "DiscriminationMentalHealth",
  "DiscriminationPregnancy", "DiscriminationUnknown", "DiscriminationReport",
  "SatisfactionBiasMeasures", "StartDate", "ApplicationDiscrimination",
  "LearnedJobContact", "LearnedJobWebsite", "LearnedJobSearchSites", 
  "LearnedJobSocialMedia", "LearnedJobAgency", "LearnedJobReferral", 
  "LearnedJobCareerFair", "LearnedJobCareerServices", "LearnedJobResearchPapers",
  "LearnedJobOther", "Onboarding_CommunicationClarity", "Onboarding_HRHelpComfort", 
  "Onboarding_Support", "ResourcesAvailability", "RecommendMPI", "AwarenessSessions",
  "TotalTime", "GroupTimeDiversity", "GroupTimeInclusion", "GroupTimeCandidateExperience"
)

# Separate Variables

vars_all <- c(
  "JobType", "Citizenship",
  "CitizenshipOther", "EthnicityGermany", "EthnicityEU", "EthnicityNonEU",
  "EthnicityEastAsia", "EthnicitySouthAsia", "EthnicityWestAsia", 
  "EthnicityArabic", "EthnicityAfrica", "EthnicityUSCanada", "EthnicityLatinAmerica",
  "EthnicityOceania", "EthnicityOther", "EthnicityNA", "AgeGroup", "ParentsUni",
  "GenderIdentity", "GenderOther", "GenderDiffBirth", "GenderDiffOther", "LGBTQIA",
  "LGBTQIAOther", "FirstLanguage", "Disability", "DisabilityOther", 
  "Neurodivergent", "NeurodivergentOther", "DisabilityHearing", "DisabilityVision",
  "DisabilityCognitive", "DisabilityCommunicative", "DisabilityMobility",
  "DisabilityMental", "DisabilityUnseen", "DisabilityNA", "DisabilityOtherChallenge",
  "Caretaker", "CaretakerOther", "ComfortSocialClass", "ComfortNationality", 
  "ComfortEthnicity", "ComfortAge", "ComfortSexualOrientation", "ComfortGender", 
  "ComfortPolitics", "ComfortReligion", "ComfortDisability", "ComfortParenthood", 
  "ComfortMentalHealth", "ComfortPregnancy", "FeltDiscriminated", "Belonging", 
  "DiversityPersonal", "DiversityFuture", "ExclusionLanguage", "ComfortOpinions", 
  "OpinionValued", "WorryCommonality", "SelfIdentity", "TeamRespect", "ManagerRespect", 
  "IdeaQuality", "FairTaskDivision", "LeaderDiversityCommitment", 
  "InstituteDiversitySupport", "DiversityPreparedness", "NoDiscrimination", 
  "DiscriminationEducation", "DiscriminationGermanLanguage", "DiscriminationEnglishLanguage",
  "DiscriminationNationality", "DiscriminationEthnicity", "DiscriminationAge",
  "DiscriminationSexualOrientation", "DiscriminationGender", "DiscriminationReligion",
  "DiscriminationPhysical", "DiscriminationParenthood", "DiscriminationMentalHealth",
  "DiscriminationPregnancy", "DiscriminationUnknown", "DiscriminationReport",
  "SatisfactionBiasMeasures", "ApplicationDiscrimination",
  "LearnedJobContact", "LearnedJobWebsite", "LearnedJobSearchSites", 
  "LearnedJobSocialMedia", "LearnedJobAgency", "LearnedJobReferral", 
  "LearnedJobCareerFair", "LearnedJobCareerServices", "LearnedJobResearchPapers",
  "LearnedJobOther", "Onboarding_HiringCommunicationClarity", "Onboarding_HRHelpComfort", 
  "Onboarding_Support", "ResourcesAvailability", "RecommendMPI", "AwarenessSessions"
  
  )

demographics <- c(
  "AgeGroup", "JobType", "Citizenship", "CitizenshipOther", 
  "EthnicityGermany", "EthnicityEU", "EthnicityNonEU", "EthnicityEastAsia",
  "EthnicitySouthAsia", "EthnicityWestAsia", "EthnicityArabic", "EthnicityAfrica",
  "EthnicityUSCanada", "EthnicityLatinAmerica", "EthnicityOceania",
  "EthnicityOther", "EthnicityNA", "GenderIdentity", "GenderOther",
  "GenderDiffBirth", "GenderDiffOther", "LGBTQIA", "LGBTQIAOther"
)

disability_neuro <- c(
  "Disability", "DisabilityOther", "Neurodivergent",
  "DisabilityHearing", "DisabilityVision", "DisabilityCognitive",
  "DisabilityCommunicative", "DisabilityMobility", "DisabilityMental",
  "DisabilityUnseen", "DisabilityNA", "DisabilityOtherChallenge"
)

caretaking <- c("Caretaker", "CaretakerOther")

comfort_inclusion <- c(
  "ComfortSocialClass", "ComfortNationality", "ComfortEthnicity", 
  "ComfortAge", "ComfortSexualOrientation", "ComfortGender",
  "ComfortPolitics", "ComfortReligion", "ComfortDisability", 
  "ComfortParenthood", "ComfortMentalHealth", "ComfortPregnancy"
  
)

attitudes <- c("FeltDiscriminated", "Belonging", 
               "DiversityPersonal", "DiversityFuture", "ExclusionLanguage", "ComfortOpinions", 
               "OpinionValued", "WorryCommonality", "SelfIdentity", "TeamRespect", "ManagerRespect", 
               "IdeaQuality", "FairTaskDivision", "LeaderDiversityCommitment", 
               "InstituteDiversitySupport", "DiversityPreparedness","SatisfactionBiasMeasures" , "ApplicationDiscrimination", "AwarenessSessions")

discrimination <- c(
  "FeltDiscriminated", "NoDiscrimination", "DiscriminationEducation", 
  "DiscriminationGermanLanguage", "DiscriminationEnglishLanguage",
  "DiscriminationNationality", "DiscriminationEthnicity", 
  "DiscriminationAge", "DiscriminationSexualOrientation", 
  "DiscriminationGender", "DiscriminationReligion", "DiscriminationPhysical", 
  "DiscriminationParenthood", "DiscriminationMentalHealth", 
  "DiscriminationPregnancy", "DiscriminationUnknown", "DiscriminationReport"
)

recruitment <- c(
  "LearnedJobContact", "LearnedJobWebsite", "LearnedJobSearchSites",
  "LearnedJobSocialMedia", "LearnedJobAgency", "LearnedJobReferral",
  "LearnedJobCareerFair", "LearnedJobCareerServices", 
  "LearnedJobResearchPapers", "LearnedJobOther",
  "Onboarding_CommunicationClarity", "Onboarding_HRHelpComfort", "Onboarding_Support",
  "ResourcesAvailability", "RecommendMPI"
)

library(tidyr)
library(dplyr)

# Recode agreement scales
agreement_levels <- c("fully disagree", "partially disagree", "neither agree nor disagree", 
                      "partially agree", "fully agree")
agreement_map <- setNames(1:5, agreement_levels)

recode_agreement <- function(column) {
  factor(column, levels = agreement_levels) |> as.numeric()
}

# Recode binary scales ("yes"/"no")
recode_binary <- function(column) {
  ifelse(column == "yes", 1, ifelse(column == "no", 0, NA))
}

# Recode satisfaction scale
satisfaction_levels <- c("very dissatisfied", "dissatisfied", "neither / nor", 
                         "satisfied", "very satisfied", "does not apply", 
                         "i don't want to answer this question")
satisfaction_map <- setNames(1:6, satisfaction_levels[-length(satisfaction_levels)])

recode_satisfaction <- function(column) {
  factor(column, levels = satisfaction_levels) |> as.numeric()
}

# Recode utility scale
utility_levels <- c("not useful", "rather not useful", "rather useful", "very useful", 
                    "i don't want to answer this question")
utility_map <- setNames(1:4, utility_levels[-length(utility_levels)])

recode_utility <- function(column) {
  factor(column, levels = utility_levels) |> as.numeric()
}

# Apply recoding
agreement_columns <- c(
  "Belonging", "DiversityPersonal", "DiversityFuture", "ExclusionLanguage",
  "ComfortOpinions", "OpinionValued", "WorryCommonality", "SelfIdentity",
  "TeamRespect", "ManagerRespect", "IdeaQuality", "FairTaskDivision",
  "LeaderDiversityCommitment", "InstituteDiversitySupport", "DiversityPreparedness",
  "Onboarding_CommunicationClarity", "Onboarding_HRHelpComfort", "Onboarding_Support", "ResourcesAvailability"
)

colnames(df)
satisfaction_columns <- c("SatisfactionBiasMeasures")
utility_columns <- c("AwarenessSessions")

# Recode each set of columns
df[agreement_columns] <- lapply(df[agreement_columns], recode_agreement)
df[satisfaction_columns] <- lapply(df[satisfaction_columns], recode_satisfaction)
df[utility_columns] <- lapply(df[utility_columns], recode_utility)


# Derive Ethnicity column
df$Ethnicity <- case_when(
  df$EthnicityEU == "yes" ~ "EU",
  df$EthnicityNonEU == "yes" ~ "Non-EU Europe",
  df$EthnicityEastAsia == "yes" ~ "Asia",
  df$EthnicitySouthAsia == "yes" ~ "Asia",
  df$EthnicityWestAsia == "yes" ~ "Other",
  df$EthnicityArabic == "yes" ~ "Other",
  df$EthnicityAfrica == "yes" ~ "Other",
  df$EthnicityUSCanada == "yes" ~ "Other",
  df$EthnicityLatinAmerica == "yes" ~ "Latin America",
  df$EthnicityOceania == "yes" ~ "Oceania",
  df$EthnicityOther == "yes" ~ "Other",
  df$EthnicityNA == "yes" ~ "Other",
  TRUE ~ "No Response" # Catch-all for those without a "yes" response
)

# Check the distribution of the new Ethnicity column
table(df$Ethnicity)

vars_all <- c(
  "JobType", "Citizenship",
  "Ethnicity", "AgeGroup", "ParentsUni",
  "GenderIdentity", "GenderOther", "GenderDiffBirth",  "LGBTQIA",
   "FirstLanguage", "Disability",  
  "Neurodivergent",  "DisabilityHearing", "DisabilityVision",
  "DisabilityCognitive", "DisabilityCommunicative", "DisabilityMobility",
  "DisabilityMental", "DisabilityUnseen", "DisabilityNA",
  "Caretaker",  "ComfortSocialClass", "ComfortNationality", 
  "ComfortEthnicity", "ComfortAge", "ComfortSexualOrientation", "ComfortGender", 
  "ComfortPolitics", "ComfortReligion", "ComfortDisability", "ComfortParenthood", 
  "ComfortMentalHealth", "ComfortPregnancy", "FeltDiscriminated", "Belonging", 
  "DiversityPersonal", "DiversityFuture", "ExclusionLanguage", "ComfortOpinions", 
  "OpinionValued", "WorryCommonality", "SelfIdentity", "TeamRespect", "ManagerRespect", 
  "IdeaQuality", "FairTaskDivision", "LeaderDiversityCommitment", 
  "InstituteDiversitySupport", "DiversityPreparedness", "NoDiscrimination", 
  "DiscriminationEducation", "DiscriminationGermanLanguage", "DiscriminationEnglishLanguage",
  "DiscriminationNationality", "DiscriminationEthnicity", "DiscriminationAge",
  "DiscriminationSexualOrientation", "DiscriminationGender", "DiscriminationReligion",
  "DiscriminationPhysical", "DiscriminationParenthood", "DiscriminationMentalHealth",
  "DiscriminationPregnancy", "DiscriminationUnknown", "DiscriminationReport",
  "SatisfactionBiasMeasures", "ApplicationDiscrimination",
  "LearnedJobContact", "LearnedJobWebsite", "LearnedJobSearchSites", 
  "LearnedJobSocialMedia", "LearnedJobAgency", "LearnedJobReferral", 
  "LearnedJobCareerFair", "LearnedJobCareerServices", "LearnedJobResearchPapers",
   "Onboarding_CommunicationClarity", "Onboarding_HRHelpComfort", 
  "Onboarding_Support", "ResourcesAvailability", "RecommendMPI", "AwarenessSessions"
  
)

colnames(df)

demographics <- c(
  "AgeGroup", "JobType", "Citizenship", 
  "Ethnicity", "GenderIdentity", 
  "GenderDiffBirth",  "LGBTQIA", "FirstLanguage"
)

# Combine all numerical columns
numerical_columns <- c(agreement_columns, satisfaction_columns, utility_columns)

# Function to create _num lists and modify original lists
generate_num_and_modify <- function(list_name, list_vars, numerical_columns) {
  overlap <- intersect(list_vars, numerical_columns)
  if (length(overlap) > 0) {
    # Create the _num list
    assign(paste0(list_name, "_num"), overlap, envir = .GlobalEnv)
    # Update the original list to exclude numerical columns
    assign(list_name, setdiff(list_vars, overlap), envir = .GlobalEnv)
  }
}

# Apply the function to each list
generate_num_and_modify("disability_neuro", disability_neuro, numerical_columns)
generate_num_and_modify("caretaking", caretaking, numerical_columns)
generate_num_and_modify("comfort_inclusion", comfort_inclusion, numerical_columns)
generate_num_and_modify("attitudes", attitudes, numerical_columns)
generate_num_and_modify("discrimination", discrimination, numerical_columns)
generate_num_and_modify("recruitment", recruitment, numerical_columns)

create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each category variable
  for (category in categories) {
    if (!category %in% colnames(data)) {
      warning(paste("Category", category, "is not in the dataframe. Skipping."))
      next
    }
    
    # Check if the column has data
    if (all(is.na(data[[category]]))) {
      warning(paste("Category", category, "contains only NA values. Skipping."))
      next
    }
    
    # Ensure the category variable is a factor
    data[[category]] <- factor(data[[category]], exclude = NULL)
    
    # Calculate counts
    counts <- table(data[[category]])
    
    if (length(counts) == 0) {
      warning(paste("Category", category, "has no levels with data. Skipping."))
      next
    }
    
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
  
  if (length(all_freq_tables) == 0) {
    stop("No valid categories were processed. Please check the input data and categories.")
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}



df_freq <- create_frequency_tables(df, vars_all)


# Recode Citizenship Column
df$Citizenship <- case_when(
  df$Citizenship %in% c("germany") ~ "Germany",
  df$Citizenship %in% c("elsewhere in europe outside the eu/eea", 
                        "elsewhere in europe within the eu/eea") ~ "Elsewhere in Europe",
  df$Citizenship %in% c("asia") ~ "Asia",
  TRUE ~ "Others" # Combine all other categories
)


# Recode AgeGroup Column
df$AgeGroup <- case_when(
  df$AgeGroup %in% c("18 - 24") ~ "18 - 24",
  df$AgeGroup %in% c("25 - 34") ~ "25 - 34",
  df$AgeGroup %in% c("35 - 50") ~ "35 - 50",
  df$AgeGroup %in% c("51 - 60", "over 60") ~ "Over 50",
  TRUE ~ "I don't want to answer this question" # Handle "I don't want to answer this question"
)


# Recode ParentsUni Column
df$ParentsUni <- case_when(
  df$ParentsUni %in% c("do not know / not sure", "i don't want to answer this question") ~ "Do not know / Prefer not to answer",
  df$ParentsUni == "no, neither of my parents / guardian attended university" ~ "No",
  df$ParentsUni == "yes, one or both of my parents / guardian attended university" ~ "Yes (University)",
  df$ParentsUni == "yes, one or both of my parents / guardian have a doctorate" ~ "Yes (Doctorate)",
  TRUE ~ NA_character_ # Handle unexpected values as NA
)

# Recode GenderIdentity Column
df$GenderIdentity <- case_when(
  df$GenderIdentity == "man" ~ "Man",
  df$GenderIdentity == "woman" ~ "Woman",
  df$GenderIdentity %in% c("non-binary", "other", "prefer not to say") ~ "Other/Prefer not to say",
  TRUE ~ NA_character_ # Handle unexpected values as NA
)


# Recode LGBTQIA Column
df$LGBTQIA <- case_when(
  df$LGBTQIA == "yes" ~ "Yes",
  df$LGBTQIA == "no" ~ "No",
  df$LGBTQIA %in% c("i don't want to answer this question", "other") ~ "Prefer not to say",
  TRUE ~ NA_character_ # Handle unexpected values as NA
)


# Recode Disability Column
df$Disability <- case_when(
  df$Disability == "yes" ~ "Yes",
  df$Disability == "no" ~ "No",
  df$Disability %in% c("i don't want to answer this question", "other") ~ "Prefer not to say",
  TRUE ~ NA_character_ # Handle unexpected values as NA
)

# Recode Neurodivergent Column
df$Neurodivergent <- case_when(
  df$Neurodivergent == "yes" ~ "Yes",
  df$Neurodivergent == "no" ~ "No",
  df$Neurodivergent %in% c("i don't want to answer this question", "other") ~ "Prefer not to say",
  TRUE ~ NA_character_ # Handle unexpected values as NA
)


# Recode Caretaker Column
df$Caretaker <- case_when(
  df$Caretaker == "no" ~ "No",
  df$Caretaker %in% c("yes, children", "yes, other family member") ~ "Yes",
  df$Caretaker %in% c("other", "prefer not to say") ~ "Others",
  TRUE ~ NA_character_ # Handle unexpected values as NA
)

# Recode FirstLanguage
df$FirstLanguage <- case_when(
  df$FirstLanguage == "english" ~ "English",
  df$FirstLanguage == "german" ~ "German",
  df$FirstLanguage %in% c("other", "i don't want to answer this question") ~ "Other",
  TRUE ~ NA_character_ # For any unexpected values
)


df_freq2 <- create_frequency_tables(df, vars_all)

chi_square_analysis_multiple <- function(data, row_vars, col_vars) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over column variables
  for (col_var in col_vars) {
    data[[col_var]] <- factor(data[[col_var]]) # Ensure the column variable is a factor
    
    for (row_var in row_vars) {
      data[[row_var]] <- factor(data[[row_var]]) # Ensure the row variable is a factor
      
      # Crosstab with percentages relative to columns
      crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 2)
      
      # Check if the crosstab is 2x2
      is_2x2_table <- all(dim(table(data[[row_var]], data[[col_var]])) == 2)
      
      # Perform chi-square test
      chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]), correct = !is_2x2_table)
      
      # Convert crosstab to a dataframe
      crosstab_df <- as.data.frame.matrix(crosstab)
      
      for (level in levels(data[[row_var]])) {
        level_df <- data.frame(
          "Row_Variable" = row_var,
          "Row_Level" = level,
          "Column_Variable" = col_var,
          "N" = sum(data[[row_var]] == level, na.rm = TRUE), # Add sample size (N)
          check.names = FALSE
        )
        
        if (nrow(crosstab_df) > 0 && !is.null(crosstab_df[level, , drop = FALSE])) {
          level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
        }
        
        level_df$Chi_Square <- chi_square_test$statistic
        level_df$P_Value <- chi_square_test$p.value
        
        results[[paste0(row_var, "_", col_var, "_", level)]] <- level_df
      }
    }
  }
  
  # Combine all results
  do.call(rbind, results)
}


# Percentages relative to the row
chi_square_analysis_multiple_row <- function(data, row_vars, col_vars) {
  results <- list() # Initialize an empty list to store results
  
  for (col_var in col_vars) {
    data[[col_var]] <- factor(data[[col_var]]) # Ensure the column variable is a factor
    
    for (row_var in row_vars) {
      data[[row_var]] <- factor(data[[row_var]]) # Ensure the row variable is a factor
      
      # Crosstab with percentages relative to rows
      crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 1)
      
      # Check if the crosstab is 2x2
      is_2x2_table <- all(dim(table(data[[row_var]], data[[col_var]])) == 2)
      
      # Perform chi-square test
      chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]), correct = !is_2x2_table)
      
      # Convert crosstab to a dataframe
      crosstab_df <- as.data.frame.matrix(crosstab)
      
      for (level in levels(data[[row_var]])) {
        level_df <- data.frame(
          "Row_Variable" = row_var,
          "Row_Level" = level,
          "Column_Variable" = col_var,
          "N" = sum(data[[row_var]] == level, na.rm = TRUE), # Add sample size (N)
          check.names = FALSE
        )
        
        if (nrow(crosstab_df) > 0 && !is.null(crosstab_df[level, , drop = FALSE])) {
          level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
        }
        
        level_df$Chi_Square <- chi_square_test$statistic
        level_df$P_Value <- chi_square_test$p.value
        
        results[[paste0(row_var, "_", col_var, "_", level)]] <- level_df
      }
    }
  }
  
  # Combine all results
  do.call(rbind, results)
}



### ONE WAY

# Function to perform one-way ANOVA
perform_one_way_anova <- function(data, response_vars, factors) {
  anova_results <- data.frame() # Initialize an empty dataframe to store results
  
  for (var in response_vars) {
    for (factor in factors) {
      if (!is.factor(data[[factor]])) {
        data[[factor]] <- factor(data[[factor]])
      }
      
      # One-way ANOVA
      anova_model <- aov(reformulate(factor, response = var), data = data)
      model_summary <- summary(anova_model)
      
      # Group means, SDs, and Ns
      group_stats <- aggregate(data[[var]], by = list(data[[factor]]), 
                               FUN = function(x) c(mean = mean(x, na.rm = TRUE), 
                                                   sd = sd(x, na.rm = TRUE), 
                                                   n = sum(!is.na(x))))
      group_stats <- do.call(data.frame, group_stats) # Convert to dataframe
      colnames(group_stats) <- c(factor, "Mean", "SD", "N")
      
      # Extract relevant statistics
      model_results <- data.frame(
        Variable = var,
        Effect = factor,
        Sum_Sq = model_summary[[1]]$"Sum Sq"[1],
        Mean_Sq = model_summary[[1]]$"Mean Sq"[1],
        Df = model_summary[[1]]$Df[1],
        FValue = model_summary[[1]]$"F value"[1],
        pValue = model_summary[[1]]$"Pr(>F)"[1]
      )
      
      # Expand group statistics into separate rows
      group_results <- data.frame(
        Group = group_stats[[factor]],
        GroupMean = group_stats$Mean,
        GroupSD = group_stats$SD,
        GroupN = group_stats$N
      )
      
      combined_results <- cbind(
        model_results[rep(1, nrow(group_results)), ],
        group_results
      )
      
      anova_results <- rbind(anova_results, combined_results)
    }
  }
  
  return(anova_results)
}



library(dplyr)

# Corrected function for calculating Cohen's d for independent samples
run_independent_t_tests <- function(data, group_col, group1, group2, measurement_cols) {
  results <- data.frame(Variable = character(),
                        T_Value = numeric(),
                        P_Value = numeric(),
                        Effect_Size = numeric(),
                        Group1_Size = integer(),
                        Group2_Size = integer(),
                        stringsAsFactors = FALSE)
  
  # Iterate over each measurement column
  for (var in measurement_cols) {
    # Extract group data with proper filtering and subsetting
    group1_data <- data %>%
      filter(.[[group_col]] == group1) %>%
      select(var) %>%
      pull() %>%
      na.omit()  # Ensure NA values are removed
    
    group2_data <- data %>%
      filter(.[[group_col]] == group2) %>%
      select(var) %>%
      pull() %>%
      na.omit()  # Ensure NA values are removed
    
    group1_size <- length(group1_data)
    group2_size <- length(group2_data)
    
    if (group1_size > 0 && group2_size > 0) {
      t_test_result <- t.test(group1_data, group2_data, paired = FALSE)
      
      # Calculate Cohen's d for independent samples
      pooled_sd <- sqrt(((group1_size - 1) * sd(group1_data)^2 + (group2_size - 1) * sd(group2_data)^2) / (group1_size + group2_size - 2))
      effect_size <- (mean(group1_data) - mean(group2_data)) / pooled_sd
      
      results <- rbind(results, data.frame(
        Variable = var,
        T_Value = t_test_result$statistic,
        P_Value = t_test_result$p.value,
        Effect_Size = effect_size,
        Group1_Size = group1_size,
        Group2_Size = group2_size
      ))
    }
  }
  
  return(results)
}



#DIVERSITY HYPOTHESES

# German employees are overrepresented in leadership and administrative positions compared to other nationalities

row_variables <- c("Citizenship", 
                   "Ethnicity")  # Replace with your row variable names
column_variable <- "JobType"  # Replace with your column variable name

df_11 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <- "GenderIdentity"  # Replace with your row variable names
column_variable <- "JobType"  # Replace with your column variable name

df_12 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <-  discrimination
column_variable <- "Citizenship"  # Replace with your column variable name

df_14 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <-  discrimination
column_variable <- "Ethnicity"  # Replace with your column variable name

df_142 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <-  c(recruitment)
column_variable <- "FirstLanguage"  # Replace with your column variable name

df_15 <- chi_square_analysis_multiple(df, row_variables, column_variable)

response_vars <- c(recruitment_num, "ExclusionLanguage")  # Add more variables as needed
factor <- "FirstLanguage"  # Specify the factor (independent variable)

# Perform one-way ANOVA
df_151 <- perform_one_way_anova(df, response_vars, factor)



# INCLUSION HYPOTHESES

response_vars <- c(attitudes_num)  # Add more variables as needed
factor <- demographics  # Specify the factor (independent variable)

# Perform one-way ANOVA
df_21 <- perform_one_way_anova(df, response_vars, factor)

row_variables <- demographics 
column_variable <- "FeltDiscriminated" # Replace with your column variable name

df_22 <- chi_square_analysis_multiple_row(df, row_variables, column_variable)

row_variables <- comfort_inclusion 
column_variable <- "JobType" # Replace with your column variable name

df_221 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <- comfort_inclusion 
column_variable <- "GenderIdentity" # Replace with your column variable name

df_222 <- chi_square_analysis_multiple(df, row_variables, column_variable)


row_variables <- comfort_inclusion 
column_variable <- "Citizenship" # Replace with your column variable name

df_223 <- chi_square_analysis_multiple(df, row_variables, column_variable)


row_variables <- c("InstituteDiversitySupport", "Onboarding_Support")
column_variable <- "FirstLanguage" # Replace with your column variable name

df_23 <- chi_square_analysis_multiple(df, row_variables, column_variable)


row_variables <- discrimination
column_variable <- "Citizenship" # Replace with your column variable name

df_24 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <- discrimination
column_variable <- "Ethnicity" # Replace with your column variable name

df_241 <- chi_square_analysis_multiple(df, row_variables, column_variable)


# CANDIDATE EXPERIENCE HYPOTHESES
row_variables <- recruitment
column_variable <- "JobType" # Replace with your column variable name

df_31 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <- recruitment
column_variable <- "Ethnicity" # Replace with your column variable name

df_32 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <- recruitment
column_variable <- "GenderIdentity" # Replace with your column variable name

df_322 <- chi_square_analysis_multiple(df, row_variables, column_variable)

row_variables <- recruitment
column_variable <- "Disability" # Replace with your column variable name

df_323 <- chi_square_analysis_multiple(df, row_variables, column_variable)



## CORRELATION ANALYSIS

calculate_correlation_matrix <- function(data, variables, method = "pearson") {
  # Ensure the method is valid
  if (!method %in% c("pearson", "spearman")) {
    stop("Method must be either 'pearson' or 'spearman'")
  }
  
  # Subset the data for the specified variables
  data_subset <- data[variables]
  
  # Calculate the correlation matrix
  corr_matrix <- cor(data_subset, method = method, use = "complete.obs")
  
  # Initialize a matrix to hold formatted correlation values
  corr_matrix_formatted <- matrix(nrow = nrow(corr_matrix), ncol = ncol(corr_matrix))
  rownames(corr_matrix_formatted) <- colnames(corr_matrix)
  colnames(corr_matrix_formatted) <- colnames(corr_matrix)
  
  # Calculate p-values and apply formatting with stars
  for (i in 1:nrow(corr_matrix)) {
    for (j in 1:ncol(corr_matrix)) {
      if (i == j) {  # Diagonal elements (correlation with itself)
        corr_matrix_formatted[i, j] <- format(round(corr_matrix[i, j], 3), nsmall = 3)
      } else {
        p_value <- cor.test(data_subset[[i]], data_subset[[j]], method = method)$p.value
        stars <- ifelse(p_value < 0.01, "***", 
                        ifelse(p_value < 0.05, "**", 
                               ifelse(p_value < 0.1, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}

vars <- c("RecommendMPI", "Belonging", "SatisfactionBiasMeasures")
df_33 <- calculate_correlation_matrix(df, vars, method = "spearman")

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
  "Cleaned Data" = df,   
  "Frequency Table" = df_freq2,
  "Gender Analysis" = gender_table,
  "Leadership Analysis" = leadership_table,
  "Hypotheses 1.1" = df_11,              
  "Hypotheses 1.2" = df_12,              
  "Hypotheses 1.4" = df_14,
  "Hypotheses 1.42" = df_142, 
  "Hypotheses 1.5" = df_15,              
  "Hypotheses 1.51" = df_151,             
  "Hypotheses 2.1" = df_21,             
  "Hypotheses 2.2" = df_22, 
  "Hypotheses 2.21" = df_221,
  "Hypotheses 2.22" = df_222,
  "Hypotheses 2.23" = df_223,
  "Hypotheses 2.3" = df_23,             
  "Hypotheses 2.4" = df_24,             
  "Hypotheses 2.41" = df_241,            
  "Hypotheses 3.1" = df_31,             
  "Hypotheses 3.2" = df_32,             
  "Hypotheses 3.3" = df_33,             
  "Hypotheses 3.22" = df_322,            
  "Hypotheses 3.23" = df_323            
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")



#Visuals

# Load required libraries
library(ggplot2)
library(dplyr)
library(reshape2)


# Helper function to save plots
save_plot <- function(plot, filename) {
  ggsave(filename, plot = plot, width = 8, height = 6, dpi = 300)
}

# 1. Identity Distributions
plot_identity_distribution <- function(data, column, title) {
  grouped_data <- data %>%
    group_by_at(column) %>%
    summarise(Count = n()) %>%
    mutate(Percentage = (Count / sum(Count)) * 100)
  
  p <- ggplot(grouped_data, aes_string(x = column, y = "Percentage", fill = column)) +
    geom_bar(stat = "identity") +
    geom_text(aes(label = sprintf("%.1f%%", Percentage)), vjust = -0.3, size = 4) +
    labs(title = title, x = column, y = "Percentage") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 1))
  
  save_plot(p, paste0(column, "_identity_distribution.png"))
  return(p)
}

# Example: Gender Identity Distribution
plot_identity_distribution(df, "GenderIdentity", "Gender Identity Distribution")

# 2. Comfort in Expression (Loop)
comfort_columns <- c("ComfortSocialClass", "ComfortNationality", "ComfortEthnicity", 
                     "ComfortAge", "ComfortSexualOrientation", "ComfortGender", 
                     "ComfortPolitics", "ComfortReligion", "ComfortDisability", 
                     "ComfortParenthood", "ComfortMentalHealth", "ComfortPregnancy")

for (column in comfort_columns) {
  grouped_data <- df %>%
    group_by_at(column) %>%
    summarise(Count = n()) %>%
    mutate(Percentage = (Count / sum(Count)) * 100)
  
  p <- ggplot(grouped_data, aes_string(x = column, y = "Percentage", fill = column)) +
    geom_bar(stat = "identity") +
    geom_text(aes(label = sprintf("%.1f%%", Percentage)), vjust = -0.3, size = 4) +
    labs(title = paste("Comfort Levels:", column), x = column, y = "Percentage") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 1))
  
  save_plot(p, paste0(column, "_comfort_distribution.png"))
}

# 3. Experiences of Discrimination
plot_discrimination <- function(data, group_col, discrimination_col, title) {
  grouped_data <- data %>%
    group_by_at(c(group_col, discrimination_col)) %>%
    summarise(Count = n()) %>%
    mutate(Percentage = (Count / sum(Count)) * 100)
  
  p <- ggplot(grouped_data, aes_string(x = discrimination_col, y = "Percentage", fill = group_col)) +
    geom_bar(stat = "identity", position = "dodge") +
    geom_text(aes(label = sprintf("%.1f%%", Percentage)), position = position_dodge(width = 0.9), vjust = -0.3, size = 3) +
    labs(title = title, x = "Discrimination Type", y = "Percentage") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 1))
  
  save_plot(p, paste0(group_col, "_", discrimination_col, "_discrimination.png"))
  return(p)
}

# Example: Discrimination Experiences by Gender Identity
plot_discrimination(df, "GenderIdentity", "FeltDiscriminated", "Experiences of Discrimination by Gender Identity")

# 4. Likert Scale Distributions
plot_likert_distribution <- function(data, likert_cols, title) {
  melted_data <- melt(data[likert_cols])
  
  grouped_data <- melted_data %>%
    group_by(variable, value) %>%
    summarise(Count = n()) %>%
    mutate(Percentage = (Count / sum(Count)) * 100)
  
  p <- ggplot(grouped_data, aes(x = variable, y = Percentage, fill = factor(value))) +
    geom_bar(stat = "identity", position = "fill") +
    geom_text(aes(label = sprintf("%.1f%%", Percentage)), position = position_fill(vjust = 0.5), size = 3) +
    labs(title = title, x = "Likert Items", y = "Proportion", fill = "Response") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 1))
  
  save_plot(p, "Likert_Distributions.png")
  return(p)
}

# Example: Distribution of Likert Statements
likert_columns <- c("Belonging", "DiversityPersonal", "DiversityFuture")
plot_likert_distribution(df, likert_columns, "Distribution of Likert Scale Responses")

# 5. Recommendation to Work at MPI-TM
plot_recommendation <- function(data, column, title) {
  grouped_data <- data %>%
    group_by_at(column) %>%
    summarise(Count = n()) %>%
    mutate(Percentage = (Count / sum(Count)) * 100)
  
  p <- ggplot(grouped_data, aes_string(x = column, y = "Percentage", fill = column)) +
    geom_bar(stat = "identity") +
    geom_text(aes(label = sprintf("%.1f%%", Percentage)), vjust = -0.3, size = 4) +
    labs(title = title, x = "Recommendation Level", y = "Percentage") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 1))
  
  save_plot(p, paste0(column, "_recommendation.png"))
  return(p)
}

# Example: Recommendation to Work at MPI-TM
plot_recommendation(df, "RecommendMPI", "Would You Recommend Working at MPI-TM?")




# New requests

df_demogdistribution <- create_frequency_tables(df, c(demographics, disability_neuro))
df_satisfactionbias <- create_frequency_tables(df, satisfaction_columns)
df_awarenesssessions <- create_frequency_tables(df, "AwarenessSessions")




# Example usage
data_list <- list(
  "Demographic Dist" = df_demogdistribution,
  "Satisfaction" = df_satisfactionbias,
  "Awareness" = df_awarenesssessions
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "NewDistributions.xlsx")


