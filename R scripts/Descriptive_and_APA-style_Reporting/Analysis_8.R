setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/jenspatrick")

library(openxlsx)
df <- read.xlsx("results-survey448487 (17).xlsx")
attributes_df <- read.xlsx("Coding Variables.xlsx", sheet = 'Conjoint')


library(dplyr)
library(tidyr)


choice_cols <- grep("Sign_\\d+_(Selected|Rejected)", names(df), value = TRUE)

# Extract additional respondent-specific information columns
info_cols <- setdiff(names(df), choice_cols)

# Reshape the choice-related columns to long format while keeping respondent info
df_long <- df %>%
  pivot_longer(cols = choice_cols, names_to = "variable", values_to = "Card_ID") %>%
  mutate(Sign = sub("_(Selected|Rejected)$", "", variable),
         Choice_Type = ifelse(grepl("Selected", variable), "Selected", "Rejected")) %>%
  select(-variable)

# Join with Attributes

df_long <- df_long %>%
  left_join(attributes_df, by = c("Card_ID" = "Card_ID"))


library(mlogit)
library(broom)
library(tibble)

run_conjoint_analysis <- function(data, choice_col, formula, id_var, chid_var, alt_var) {
  
  # Convert data to mlogit format
  mlogit_data <- mlogit.data(data, 
                             choice = choice_col, 
                             shape = "long", 
                             id.var = id_var,
                             chid.var = chid_var,
                             alt.var = alt_var)
  
  # Fit the model with the specified formula
  model <- mlogit(formula, data = mlogit_data)
  
  # Generate summary of the model
  summary_model <- summary(model)
  
  # Extracting coefficients and formatting
  coef_table <- summary_model$CoefTable
  coefficients_df <- as.data.frame(coef_table)
  coefficients_df$Variable <- rownames(coef_table)
  coefficients_df <- coefficients_df[, c("Variable", "Estimate", "Std. Error", "z-value", "Pr(>|z|)")]
  
  # Model fit statistics in a dataframe
  model_fit_df <- data.frame(
    LogLik = as.numeric(logLik(model)), 
    AIC = AIC(model), 
    BIC = BIC(model)
  )
  
  # Return the results
  return(list(Coefficients = coefficients_df, Model_Fit_Statistics = model_fit_df))
}


# Assuming 'df_long' is your data frame and 'Choice_Type' is the column
df_long$Choice_Type <- ifelse(df_long$Choice_Type == "Selected", 1, 0)
df_long$unique_scenario_id <- with(df_long, paste(id..Antwort.ID, Sign, sep = "_"))
df_long <- df_long[!is.na(df_long$Card_ID), ]
df_long <- df_long[!is.na(df_long$Choice_Type), ]
df_long <- df_long[!is.na(df_long$Sign), ]
df_long$Choice_Type <- factor(df_long$Choice_Type, levels = c(0, 1))
df_long <- df_long %>%
  filter(id..Antwort.ID != 2280) #Parcipipant without a single choice for a choice set

# Replace NA with a label indicating absence of the level
df_long$Doctorate[is.na(df_long$Doctorate)] <- "NoDoctorate"
df_long$Masters[is.na(df_long$Masters)] <- "NoMasters"

# Convert to factor with "NoDoctorate" and "NoMasters" as the first level
df_long$Doctorate <- factor(df_long$Doctorate, levels = c("NoDoctorate", unique(df_long$Doctorate[df_long$Doctorate != "NoDoctorate"])))
df_long$Masters <- factor(df_long$Masters, levels = c("NoMasters", unique(df_long$Masters[df_long$Masters != "NoMasters"])))


results <- run_conjoint_analysis(df_long, "Choice_Type", 
                                 Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                 "id..Antwort.ID", "unique_scenario_id", "Card_ID")

# Access the coefficients and model fit statistics dataframes
coefficients_df <- results$Coefficients
model_fit_df <- results$Model_Fit_Statistics

# For separate Genders
split_by_gender <- split(df_long, df_long$Gender)

df_men <- split_by_gender[[3]]

results_men <- run_conjoint_analysis(df_men, "Choice_Type", 
                                 Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                 "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_men <- results_men$Coefficients
df_fit_men <- results_men$Model_Fit_Statistics

df_women <- split_by_gender[[4]]
results_women <- run_conjoint_analysis(df_women, "Choice_Type", 
                                     Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                     "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_women <- results_women$Coefficients
df_fit_women <- results_women$Model_Fit_Statistics

# For separate cities
split_by_city_size <- split(df_long, df_long$City_size)

df_grobe_grobstat <- split_by_city_size[[1]]
results_grobe_grobstat <- run_conjoint_analysis(df_grobe_grobstat, "Choice_Type", 
                                       Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                       "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_grobe_grobstat <- results_grobe_grobstat$Coefficients
df_fit_grobe_grobstat <- results_grobe_grobstat$Model_Fit_Statistics

df_grobstat <- split_by_city_size[[2]]
results_grobstat <- run_conjoint_analysis(df_grobstat, "Choice_Type", 
                                                Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                                "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_grobstat <- results_grobstat$Coefficients
df_fit_grobstat <- results_grobstat$Model_Fit_Statistics

df_kleinstadt <- split_by_city_size[[3]]
results_kleinstadt <- run_conjoint_analysis(df_kleinstadt, "Choice_Type", 
                                          Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                          "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_kleinstadt <- results_kleinstadt$Coefficients
df_fit_kleinstadt <- results_kleinstadt$Model_Fit_Statistics


df_landgemeinde <- split_by_city_size[[4]]
results_landgemeinde <- run_conjoint_analysis(df_landgemeinde, "Choice_Type", 
                                          Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                          "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_landgemeinde <- results_landgemeinde$Coefficients
df_fit_landgemeinde <- results_landgemeinde$Model_Fit_Statistics


df_mittelstadt <- split_by_city_size[[5]]
results_mittelstadt <- run_conjoint_analysis(df_mittelstadt, "Choice_Type", 
                                          Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                          "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_mittelstadt <- results_mittelstadt$Coefficients
df_fit_mittelstadt <- results_mittelstadt$Model_Fit_Statistics


# For separate 
split_by_psych_experience <- split(df_long, df_long$Psych_experience)

df_Ja <- split_by_psych_experience[[1]]
results_Ja <- run_conjoint_analysis(df_Ja, "Choice_Type", 
                                             Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                             "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_Ja <- results_Ja$Coefficients
df_fit_Ja <- results_Ja$Model_Fit_Statistics

df_Nein <- split_by_psych_experience[[2]]
results_Nein <- run_conjoint_analysis(df_Nein, "Choice_Type", 
                                    Choice_Type ~ 0 + Profession + Doctorate + Masters + Specialization.1 + Specialization.2, 
                                    "id..Antwort.ID", "unique_scenario_id", "Card_ID")

df_model_Nein <- results_Nein$Coefficients
df_fit_Nein <- results_Nein$Model_Fit_Statistics

# # Summarize choices per scenario
# choice_summary <- df_cleaned %>%
#   group_by(id..Antwort.ID, unique_scenario_id) %>%
#   summarize(total_choices = sum(Choice_Type)) %>%
#   ungroup()
# 
# # Identify scenarios to keep (with exactly one choice made)
# scenarios_to_keep <- choice_summary %>%
#   filter(total_choices == 1) %>%
#   select(id..Antwort.ID, unique_scenario_id)


create_segmented_frequency_tables <- function(data, vars, segmenting_categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each segmenting category
  for (segment in segmenting_categories) {
    # Iterate over each variable for which frequencies are to be calculated
    for (var in vars) {
      # Ensure both the segmenting category and variable are treated as factors
      segment_factor <- factor(data[[segment]])
      var_factor <- factor(data[[var]])
      
      # Calculate counts segmented by the segmenting category
      counts <- table(segment_factor, var_factor)
      
      # Melt the table into a long format for easier handling
      freq_table <- as.data.frame(counts)
      names(freq_table) <- c("Segment", "Level", "Count")
      
      # Add the variable name and calculate percentages
      freq_table$Variable <- var
      freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


df_freq_choice <- create_segmented_frequency_tables(df_long, c("Profession" ,                                          
                                                                  "Doctorate" ,                                            
                                                                  "Masters"   ,                                            
                                                                  "Specialization.1",                                     
                                                                  "Specialization.2" ), "Choice_Type")

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



vars_sociod <- c("Gender"    ,                                            
                  "Age_group" ,                                            
                 "Marital_status" ,                                       
                 "City_size"       ,                                      
                 "Employment_status",                                     
                 "Education"         ,                                    
                 "Net_income"         ,                                   
                 "Psych_experience"    ,                                  
                 "Psych_consideration" )
df_freq_sociod <- create_frequency_tables(df, vars_sociod)

vars_frequency <- c("Experience_HP"                  ,                       
                      "Experience_HPP"                  ,                      
                      "Experience_Psychologe_HPP"        ,                     
                      "Experience_Psychologischer_PT"     ,                    
                      "Experience_Aerztlicher_PT"          ,                   
                      "Experience_None"                     ,                  
                      "Experience_Other"                     ,                 
                      "Consideration_HP"                      ,                
                      "Consideration_HPP"                      ,               
                      "Consideration_Psychologe_HPP"            ,              
                      "Consideration_Psychologischer_PT"         ,             
                      "Consideration_Aerztlicher_PT"              ,            
                      "Consideration_None"                         ,           
                              
                      "Reason_Experience_HP_Access"                  ,         
                      "Reason_Experience_HP_Waiting_Time"             ,        
                      "Reason_Experience_HP_Competence"                ,       
                      "Reason_Experience_HP_Qualifications"             ,      
                      "Reason_Experience_HP_Empathy"                     ,     
                      "Reason_Experience_HP_Cost"                         ,    
                      "Reason_Experience_HP_Recommendation"                ,   
                      "Reason_Experience_HP_Reviews"                        ,  
                      "Reason_Experience_HP_Proximity"                       , 
                      
                      "Reason_Experience_HPP_Access"                          ,
                      "Reason_Experience_HPP_Waiting_Time"                    ,
                      "Reason_Experience_HPP_Competence"                      ,
                      "Reason_Experience_HPP_Qualifications"                  ,
                      "Reason_Experience_HPP_Empathy"                         ,
                      "Reason_Experience_HPP_Cost"                            ,
                      "Reason_Experience_HPP_Recommendation"                  ,
                      "Reason_Experience_HPP_Reviews"                         ,
                      "Reason_Experience_HPP_Proximity"                       ,
                     
                      "Reason_Experience_Psychologe_HPP_Access"               ,
                      "Reason_Experience_Psychologe_HPP_Waiting_Time"         ,
                      "Reason_Experience_Psychologe_HPP_Competence"           ,
                      "Reason_Experience_Psychologe_HPP_Qualifications"       ,
                      "Reason_Experience_Psychologe_HPP_Empathy"              ,
                      "Reason_Experience_Psychologe_HPP_Cost"                 ,
                      "Reason_Experience_Psychologe_HPP_Recommendation"       ,
                      "Reason_Experience_Psychologe_HPP_Reviews"              ,
                      "Reason_Experience_Psychologe_HPP_Proximity"            ,
                   
                      "Reason_Experience_Psychologischer_PT_Access"           ,
                      "Reason_Experience_Psychologischer_PT_Waiting_Time"     ,
                      "Reason_Experience_Psychologischer_PT_Competence"       ,
                      "Reason_Experience_Psychologischer_PT_Qualifications"   ,
                      "Reason_Experience_Psychologischer_PT_Empathy"          ,
                      "Reason_Experience_Psychologischer_PT_Cost"             ,
                      "Reason_Experience_Psychologischer_PT_Recommendation"   ,
                      "Reason_Experience_Psychologischer_PT_Reviews"          ,
                      "Reason_Experience_Psychologischer_PT_Proximity"        ,
                      
                      "Reason_Experience_Aerztlicher_PT_Access"               ,
                      "Reason_Experience_Aerztlicher_PT_Waiting_Time"         ,
                      "Reason_Experience_Aerztlicher_PT_Competence"           ,
                      "Reason_Experience_Aerztlicher_PT_Qualifications"       ,
                      "Reason_Experience_Aerztlicher_PT_Empathy"              ,
                      "Reason_Experience_Aerztlicher_PT_Cost"                 ,
                      "Reason_Experience_Aerztlicher_PT_Recommendation"       ,
                      "Reason_Experience_Aerztlicher_PT_Reviews"              ,
                      "Reason_Experience_Aerztlicher_PT_Proximity"            ,
                      
                      "Reason_Consideration_HP_Access"                        ,
                      "Reason_Consideration_HP_Waiting_Time"                  ,
                      "Reason_Consideration_HP_Competence"                    ,
                      "Reason_Consideration_HP_Qualifications"                ,
                      "Reason_Consideration_HP_Empathy"                       ,
                      "Reason_Consideration_HP_Cost"                          ,
                      "Reason_Consideration_HP_Recommendation"                ,
                      "Reason_Consideration_HP_Reviews"                       ,
                      "Reason_Consideration_HP_Proximity"                     ,
                      
                      "Reason_Consideration_HPP_Access"                       ,
                      "Reason_Consideration_HPP_Waiting_Time"                 ,
                      "Reason_Consideration_HPP_Competence"                   ,
                      "Reason_Consideration_HPP_Qualifications"               ,
                      "Reason_Consideration_HPP_Empathy"                      ,
                      "Reason_Consideration_HPP_Cost"                         ,
                      "Reason_Consideration_HPP_Recommendation"               ,
                      "Reason_Consideration_HPP_Reviews"                      ,
                      "Reason_Consideration_HPP_Proximity"                    ,
                      
                      "Reason_Consideration_Psychologe_HPP_Access"            ,
                      "Reason_Consideration_Psychologe_HPP_Waiting_Time"      ,
                      "Reason_Consideration_Psychologe_HPP_Competence"        ,
                      "Reason_Consideration_Psychologe_HPP_Qualifications"    ,
                      "Reason_Consideration_Psychologe_HPP_Empathy"           ,
                      "Reason_Consideration_Psychologe_HPP_Cost"              ,
                      "Reason_Consideration_Psychologe_HPP_Recommendation"    ,
                      "Reason_Consideration_Psychologe_HPP_Reviews"           ,
                      "Reason_Consideration_Psychologe_HPP_Proximity"         ,
                      
                      "Reason_Consideration_Psychologischer_PT_Access"        ,
                      "Reason_Consideration_Psychologischer_PT_Waiting_Time"  ,
                      "Reason_Consideration_Psychologischer_PT_Competence"    ,
                      "Reason_Consideration_Psychologischer_PT_Qualifications",
                      "Reason_Consideration_Psychologischer_PT_Empathy"       ,
                      "Reason_Consideration_Psychologischer_PT_Cost"          ,
                      "Reason_Consideration_Psychologischer_PT_Recommendation",
                      "Reason_Consideration_Psychologischer_PT_Reviews"       ,
                      "Reason_Consideration_Psychologischer_PT_Proximity"     ,
                      
                      "Reason_Consideration_Aerztlicher_PT_Access"            ,
                      "Reason_Consideration_Aerztlicher_PT_Waiting_Time"      ,
                      "Reason_Consideration_Aerztlicher_PT_Competence"        ,
                      "Reason_Consideration_Aerztlicher_PT_Qualifications"    ,
                      "Reason_Consideration_Aerztlicher_PT_Empathy"           ,
                      "Reason_Consideration_Aerztlicher_PT_Cost"              ,
                      "Reason_Consideration_Aerztlicher_PT_Recommendation"    ,
                      "Reason_Consideration_Aerztlicher_PT_Reviews"           ,
                      "Reason_Consideration_Aerztlicher_PT_Proximity"         ,
                      
                      "Disorder_HP_Cognitive"                                 ,
                      "Disorder_HP_Substance"                                 ,
                      "Disorder_HP_Psychotic"                                 ,
                      "Disorder_HP_Affect"                                    ,
                      "Disorder_HP_Anxiety"                                   ,
                      "Disorder_HP_Somatoform"                                ,
                      "Disorder_HP_Eat_Sleep_Sex"                             ,
                      "Disorder_HP_Personality"                               ,
                      "Disorder_HP_Couple"                                    ,
                      
                      "Disorder_HPP_Cognitive"                                ,
                      "Disorder_HPP_Substance"                                ,
                      "Disorder_HPP_Psychotic"                                ,
                      "Disorder_HPP_Affect"                                   ,
                      "Disorder_HPP_Anxiety"                                  ,
                      "Disorder_HPP_Somatoform"                               ,
                      "Disorder_HPP_Eat_Sleep_Sex"                            ,
                      "Disorder_HPP_Personality"                              ,
                      "Disorder_HPP_Couple"                                   ,
                      
                      "Disorder_Psychologe_HPP_Cognitive"                     ,
                      "Disorder_Psychologe_HPP_Substance"                     ,
                      "Disorder_Psychologe_HPP_Psychotic"                     ,
                      "Disorder_Psychologe_HPP_Affect"                        ,
                      "Disorder_Psychologe_HPP_Anxiety"                       ,
                      "Disorder_Psychologe_HPP_Somatoform"                    ,
                      "Disorder_Psychologe_HPP_Eat_Sleep_Sex"                 ,
                      "Disorder_Psychologe_HPP_Personality"                   ,
                      "Disorder_Psychologe_HPP_Couple"                        ,
                      
                      "Disorder_Psychologischer_PT_Cognitive"                 ,
                      "Disorder_Psychologischer_PT_Substance"                 ,
                      "Disorder_Psychologischer_PT_Psychotic"                 ,
                      "Disorder_Psychologischer_PT_Affect"                    ,
                      "Disorder_Psychologischer_PT_Anxiety"                   ,
                      "Disorder_Psychologischer_PT_Somatoform"                ,
                      "Disorder_Psychologischer_PT_Eat_Sleep_Sex"             ,
                      "Disorder_Psychologischer_PT_Personality"               ,
                      "Disorder_Psychologischer_PT_Couple"                    ,
                      
                      "Disorder_Aerztlicher_PT_Cognitive"                     ,
                      "Disorder_Aerztlicher_PT_Substance"                     ,
                      "Disorder_Aerztlicher_PT_Psychotic"                     ,
                      "Disorder_Aerztlicher_PT_Affect"                        ,
                      "Disorder_Aerztlicher_PT_Anxiety"                       ,
                      "Disorder_Aerztlicher_PT_Somatoform"                    ,
                      "Disorder_Aerztlicher_PT_Eat_Sleep_Sex"                 ,
                      "Disorder_Aerztlicher_PT_Personality"                   ,
                      "Disorder_Aerztlicher_PT_Couple"                        ,
                      
                      "Therapy_HP_Psychodynamic"                              ,
                      "Therapy_HP_Cognitve_Behavioral"                        ,
                      "Therapy_HP_Humanist"                                   ,
                      "Therapy_HP_Systemic"                                   ,
                      "Therapy_HP_Experience"                                 ,
                      "Therapy_HP_Bodyoriented"                               ,
                      "Therapy_HP_Mindfulness"                                ,
                      "Therapy_HP_Dontknow"                                   ,
                      
                      "Therapy_HPP_Psychodynamic"                             ,
                      "Therapy_HPP_Cognitve_Behavioral"                       ,
                      "Therapy_HPP_Humanist"                                  ,
                      "Therapy_HPP_Systemic"                                  ,
                      "Therapy_HPP_Experience"                                ,
                      "Therapy_HPP_Bodyoriented"                              ,
                      "Therapy_HPP_Mindfulness"                               ,
                      "Therapy_HPP_Dontknow"                                  ,
                      
                      "Therapy_Psychologe_HPP_Psychodynamic"                  ,
                      "Therapy_Psychologe_HPP_Cognitve_Behavioral"            ,
                      "Therapy_Psychologe_HPP_Humanist"                       ,
                      "Therapy_Psychologe_HPP_Systemic"                       ,
                      "Therapy_Psychologe_HPP_Experience"                     ,
                      "Therapy_Psychologe_HPP_Bodyoriented"                   ,
                      "Therapy_Psychologe_HPP_Mindfulness"                    ,
                      "Therapy_Psychologe_HPP_Dontknow"                       ,
                      
                      "Therapy_Psychologischer_PT_Psychodynamic"              ,
                      "Therapy_Psychologischer_PT_Cognitve_Behavioral"        ,
                      "Therapy_Psychologischer_PT_Humanist"                   ,
                      "Therapy_Psychologischer_PT_Systemic"                   ,
                      "Therapy_Psychologischer_PT_Experience"                 ,
                      "Therapy_Psychologischer_PT_Bodyoriented"               ,
                      "Therapy_Psychologischer_PT_Mindfulness"                ,
                      "Therapy_Psychologischer_PT_Dontknow"                   ,
                      "Therapy_Psychologischer_PT_Other"                      ,
                      "Therapy_Aerztlicher_PT_Psychodynamic"                  ,
                      "Therapy_Aerztlicher_PT_Cognitve_Behavioral"            ,
                      "Therapy_Aerztlicher_PT_Humanist"                       ,
                      "Therapy_Aerztlicher_PT_Systemic"                       ,
                      "Therapy_Aerztlicher_PT_Experience"                     ,
                      "Therapy_Aerztlicher_PT_Bodyoriented"                   ,
                      "Therapy_Aerztlicher_PT_Mindfulness"                    ,
                      "Therapy_Aerztlicher_PT_Dontknow"                       )

df_freq_others <- create_frequency_tables(df, vars_frequency)

vars_descriptive <- c(
  "Number_Sessions_HP"                                    ,
  "Number_Sessions_HPP"                                   ,
  "Number_Sessions_Psychologe_HPP"                        ,
  "Number_Sessions_Psychologischer_PT"                    ,
  "Number_Sessions_Aerztlicher_PT"                        ,
  "Satisfaction_HP_Access"                                ,
  "Satisfaction_HP_Waiting_Time"                          ,
  "Satisfaction_HP_Proximity"                             ,
  "Satisfaction_HP_Competence"                            ,
  "Satisfaction_HP_Empathy"                               ,
  "Satisfaction_HP_Relief"                                ,
  "Satisfaction_HP_Cost"                                  ,
  "Satisfaction_HPP_Access"                               ,
  "Satisfaction_HPP_Waiting_Time"                         ,
  "Satisfaction_HPP_Proximity"                            ,
  "Satisfaction_HPP_Competence"                           ,
  "Satisfaction_HPP_Empathy"                              ,
  "Satisfaction_HPP_Relief"                               ,
  "Satisfaction_HPP_Cost"                                 ,
  "Satisfaction_Psychologe_HPP_Access"                    ,
  "Satisfaction_Psychologe_HPP_Waiting_Time"              ,
  "Satisfaction_Psychologe_HPP_Proximity"                 ,
  "Satisfaction_Psychologe_HPP_Competence"                ,
  "Satisfaction_Psychologe_HPP_Empathy"                   ,
  "Satisfaction_Psychologe_HPP_Relief"                    ,
  "Satisfaction_Psychologe_HPP_Cost"                      ,
  "Satisfaction_Psychologischer_PT_Access"                ,
  "Satisfaction_Psychologischer_PT_Waiting_Time"          ,
  "Satisfaction_Psychologischer_PT_Proximity"             ,
  "Satisfaction_Psychologischer_PT_Competence"            ,
  "Satisfaction_Psychologischer_PT_Empathy"               ,
  "Satisfaction_Psychologischer_PT_Relief"                ,
  "Satisfaction_Psychologischer_PT_Cost"                  ,
  "Satisfaction_Aerztlicher_PT_Access"                    ,
  "Satisfaction_Aerztlicher_PT_Waiting_Time"              ,
  "Satisfaction_Aerztlicher_PT_Proximity"                 ,
  "Satisfaction_Aerztlicher_PT_Competence"                ,
  "Satisfaction_Aerztlicher_PT_Empathy"                   ,
  "Satisfaction_Aerztlicher_PT_Relief"                    ,
  "Satisfaction_Aerztlicher_PT_Cost"                      ,
  "Satisfaction_HP_Overall"                               ,
  "Satisfaction_HPP_Overall"                              ,
  "Satisfaction_Psychologe_HPP_Overall"                   ,
  "Satisfaction_Psychologischer_PT_Overall"               ,
  "Satisfaction_Aerztlicher_PT_Overall"                   ,
  "Professional_Group_HP_Access"                          ,
  "Professional_Group_Psychologischer_PT_Access"          ,
  "Professional_Group_Aerztlicher_PT_Access"              ,
  "Professional_Group_HP_Waiting_Time"                    ,
  "Professional_Group_Psychologischer_PT_Waiting_Time"    ,
  "Professional_Group_Aerztlicher_PT_Waiting_Time"        ,
  "Professional_Group_HP_Proximity"                       ,
  "Professional_Group_Psychologischer_PT_Proximity"       ,
  "Professional_Group_Aerztlicher_PT_Proximity"           ,
  "Professional_Group_HP_Qualified"                       ,
  "Professional_Group_Psychologischer_PT_Qualified"       ,
  "Professional_Group_Aerztlicher_PT_Qualified"           ,
  "Professional_Group_HP_Competence"                      ,
  "Professional_Group_Psychologischer_PT_Competence"      ,
  "Professional_Group_Aerztlicher_PT_Competence"          ,
  "Professional_Group_HP_Empathy"                         ,
  "Professional_Group_Psychologischer_PT_Empathy"         ,
  "Professional_Group_Aerztlicher_PT_Empathy"             ,
  "Professional_Group_HP_Relief"                          ,
  "Professional_Group_Psychologischer_PT_Relief"          ,
  "Professional_Group_Aerztlicher_PT_Relief"              ,
  "Professional_Group_HP_Cost"                            ,
  "Professional_Group_Psychologischer_PT_Cost"            ,
  "Professional_Group_Aerztlicher_PT_Cost")


## DESCRIPTIVE TABLES

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
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
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

df_descriptive_stats <- calculate_descriptive_stats(df, vars_descriptive)

df_grandtotal <- read.xlsx("frequency.xlsx", sheet = "data (2)")


vars_freqgrandtotal <- c("Experience_Any"      ,                  
                         "Experience_Any"       ,                  "ReasonExperience_Any_Access"       ,    
                         "ReasonExperience_Any_Waiting_Time"  ,    "ReasonExperience_Any_Competence"    ,   
                         "ReasonExperience_Any_Qualifications" ,   "ReasonExperience_Any_Empathy"        ,  
                         "ReasonExperience_Any_Cost"            ,  "ReasonExperience_Any_Recommendation"  , 
                          "ReasonExperience_Any_Reviews"         ,  "ReasonExperience_Any_Proximity"       , 
                          "ReasonConsideration_Any_Access"    ,     "ReasonConsideration_Any_Waiting_Time"  ,
                          "ReasonConsideration_Any_Competence" ,    "ReasonConsideration_Any_Qualifications",
                          "ReasonConsideration_Any_Empathy"     ,   "ReasonConsideration_Any_Cost"          ,
                          "ReasonConsideration_Any_Recommendation", "ReasonConsideration_Any_Reviews"       ,
                         "ReasonConsideration_Any_Proximity",      "Disorder_Any_Cognitive"                ,
                          "Disorder_Any_Substance"           ,      "Disorder_Any_Psychotic"                ,
                          "Disorder_Any_Affect"               ,     "Disorder_Any_Anxiety"                  ,
                          "Disorder_Any_Somatoform"            ,    "Disorder_Any_Eat_Sleep_Sex"            ,
                         "Disorder_Any_Personality"             ,  "Disorder_Any_Couple"                   ,
                         "Therapy_Any_Psychodynamic"            ,  "Therapy_Any_Cognitve_Behavioral"       ,
                          "Therapy_Any_Humanist"                ,   "Therapy_Any_Systemic"                  ,
                         "Therapy_Any_Experience"               ,  "Therapy_Any_Bodyoriented"              ,
                         "Therapy_Any_Mindfulness"              ,  "Therapy_Any_Dontknow")

df_freq_grandtotal <- create_frequency_tables(df_grandtotal, vars_freqgrandtotal)



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
  "Long-Shaped Data" = df,
  "Descriptive Statistics" = df_descriptive_stats,
  "Frequency Table - Choices" = df_freq_choice,
  "Frequency Table - Demographic" = df_freq_sociod,
  "Frequency Table - Others" = df_freq_others,
  "Frequency Table - GrandTotal" = df_freq_grandtotal,
  "Conjoint - Overall" = model_overall_coef,
  "Conjoint - Men" = df_model_men,
  "Conjoint - Women" = df_model_women,
  "Conjoint - Grobstradt" = df_model_grobstat,
  "Conjoint - Grobe Grobstradt" = df_model_grobe_grobstat,
  "Conjoint - kleinstadt" = df_model_kleinstadt,
  "Conjoint - Landgemeide" = df_model_landgemeinde,
  "Conjoint - Mittelstadt" = df_model_mittelstadt,
  "Conjoint - Psych Experience" = df_model_Ja,
  "Conjoint - No Psych Experience" = df_model_Nein
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

# Pairwise Comparison

# Function to perform pairwise z-tests between categories of each attribute
perform_pairwise_comparisons <- function(df, attribute_prefix, model_name) {
  # Extract variables that start with the attribute prefix
  relevant_vars <- grep(attribute_prefix, df$Variable, value = TRUE)
  
  # Prepare a matrix to store results
  results <- matrix(nrow = choose(length(relevant_vars), 2), ncol = 5,
                    dimnames = list(NULL, c("Model", "Pair", "Z Value", "P Value", "Significant")))
  
  # Counter for results matrix
  counter <- 1
  
  # Loop over all pairs of relevant variables
  for (i in 1:(length(relevant_vars) - 1)) {
    for (j in (i + 1):length(relevant_vars)) {
      var1 <- relevant_vars[i]
      var2 <- relevant_vars[j]
      
      # Extract coefficients and standard errors
      beta1 <- df[df$Variable == var1, "Estimate"]
      beta2 <- df[df$Variable == var2, "Estimate"]
      SE1 <- df[df$Variable == var1, "Std. Error"]
      SE2 <- df[df$Variable == var2, "Std. Error"]
      
      # Calculate Z statistic
      Z <- (beta1 - beta2) / sqrt(SE1^2 + SE2^2)
      
      # Calculate p-value
      p_value <- 2 * pnorm(-abs(Z))  # Two-tailed test
      
      # Store results
      results[counter, ] <- c(model_name, paste(var1, "vs", var2), Z, p_value, p_value < 0.05)
      counter <- counter + 1
    }
  }
  
  # Convert results to a data frame for easier handling
  results_df <- as.data.frame(results, stringsAsFactors = FALSE)
  results_df$`Z Value` <- as.numeric(results_df$`Z Value`)
  results_df$`P Value` <- as.numeric(results_df$`P Value`)
  results_df$Significant <- as.logical(results_df$Significant)
  
  return(results_df)
}

# Example of how to use this function on a list of models
list_of_df_models <- list(
  Overall = coefficients_df,
  Men = df_model_men,
  Women = df_model_women,
  GrobeGrobstat = df_model_grobe_grobstat,
  Grobstat = df_model_grobstat,
  Kleinstadt = df_model_kleinstadt,
  Landgemeinde = df_model_landgemeinde,
  Mittelstadt = df_model_mittelstadt,
  PsychExperienceJa = df_model_Ja,
  PsychExperienceNein = df_model_Nein
)


# Running comparisons for the Doctorate attribute across all models and integrating results
all_results_Doc <- do.call(rbind, lapply(names(list_of_df_models), function(model_name) {
  perform_pairwise_comparisons(list_of_df_models[[model_name]], "Doctorate", model_name)
}))

all_results_Profession <- do.call(rbind, lapply(names(list_of_df_models), function(model_name) {
  perform_pairwise_comparisons(list_of_df_models[[model_name]], "Profession", model_name)
}))

all_results_Specialization1 <- do.call(rbind, lapply(names(list_of_df_models), function(model_name) {
  perform_pairwise_comparisons(list_of_df_models[[model_name]], "Specialization.1", model_name)
}))

all_results_Specialization2 <- do.call(rbind, lapply(names(list_of_df_models), function(model_name) {
  perform_pairwise_comparisons(list_of_df_models[[model_name]], "Specialization.2", model_name)
}))

data_list <- list(
  "Pairwise Comp - Doc" = all_results_Doc,
  "Pairwise Comp - Profession" = all_results_Profession,
  "Pairwise Comp - Specialization1" = all_results_Specialization1,
  "Pairwise Comp - Specialization2" = all_results_Specialization2
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Tables_PairwiseComparison.xlsx")
