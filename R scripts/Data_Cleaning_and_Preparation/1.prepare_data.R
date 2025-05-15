### RCT Analysis ###
#Prepare data for analysis


rm(list=ls()) # clear dataset

library(tidyverse)

### Read in control and intervention datasets ###
raw_intervention <- read_csv("intervention_clean.csv")
raw_control <- read_csv("control_clean.csv")

### Give datasets identical structures ###
identical(colnames(raw_intervention), colnames(raw_control))

#False

control_vars  <-  colnames(raw_control)

intervention_vars  <-  colnames(raw_intervention)

intervention_vars_minus_last_one <- intervention_vars[1:75]

identical(control_vars, intervention_vars_minus_last_one)
#[1] FALSE

setdiff(control_vars, intervention_vars_minus_last_one)
#[1] "Service Line"                       "Goal_2_rating_3 (end of tratement)

position <- which((control_vars == "Goal_2_rating_3 (end of tratement)"))
#65

control_vars["Goal_2_rating_3 (end of tratement)"] <- "Goal_2_rating_3 (end of treatment)"

position <- which((control_vars == "Goal_2_rating_3 (end of tratement)"))   

control_vars[control_vars == "Goal_2_rating_3 (end of tratement)"] <- "Goal_2_rating_3 (end of treatment)"
          
setdiff(intervention_vars_minus_last_one, control_vars)
#[1] "...1"
setdiff(control_vars, intervention_vars_minus_last_one)        
#[1] "Service Line"

control_for_merge <- raw_control %>%
  slice(1:86) %>% #there's 86 controls
  mutate(`...1` = seq_len(n())) %>%
  mutate(`...1` = paste("Client ", ...1)) %>%
  rename(`Goal_2_rating_3 (end of treatment)` = `Goal_2_rating_3 (end of tratement)`) %>%
  mutate(Intervention_Control = "Control") %>%
  select(`...1`, Intervention_Control, project_name, `Service Line`, everything()) %>%
  mutate(`Data Collection (V2,3)` = "Control") 

intervention_for_merge <- raw_intervention %>%
  slice(1:76) %>% # there's 76 interventions
  mutate(`Service Line` = "Intervention") %>%
  mutate(Intervention_Control = "Intervention") %>%
  select(`...1`, Intervention_Control, project_name, `Service Line`, everything()) # variable names now align

identical(colnames(intervention_for_merge), colnames(control_for_merge))
#TRUE

setdiff(colnames(intervention_for_merge), colnames(control_for_merge))

### Rename RCDAS to RCADS ###
intervention_for_merge <- intervention_for_merge %>%
  rename_with(~ gsub("^RCDAS", "RCADS", .x), starts_with("RCDAS")) # some of the RCDAS were labelled RCDAS

control_for_merge <- control_for_merge %>%
  rename_with(~ gsub("^RCDAS", "RCADS", .x), starts_with("RCDAS")) # some of the RCDAS were labelled RCDAS

#### Convert all character variables to numeric ####

control_for_merge <- control_for_merge %>%
  mutate(across(starts_with("RCADS"), as.numeric)) %>%
  mutate(across(starts_with("SDQ"), as.numeric)) %>%
  mutate(across(starts_with("YPCORE10"), as.numeric))
  
  
  intervention_for_merge <- intervention_for_merge %>%
  mutate(across(starts_with("RCADS"), as.numeric)) %>%
  mutate(across(starts_with("SDQ"), as.numeric)) %>%
  mutate(across(starts_with("YPCORE10"), as.numeric)) 
  
### Test if names identical ###
  identical(colnames(intervention_for_merge), colnames(control_for_merge))
  #TRUE
  
#### Merge intervention and control ####
  
    merged <- intervention_for_merge %>%
    bind_rows(control_for_merge)

#### Explore missings ####
 
control_summary_stats <- data.frame(
  Variable = names(control_for_merge),
  Valid_Cases = sapply(control_for_merge, function(x) sum(!is.na(x))),
  Missing_Values = sapply(control_for_merge, function(x) sum(is.na(x)))
)

control_summary_stats

intervention_summary_stats <- data.frame(
  Variable = names(intervention_for_merge),
  Valid_Cases = sapply(intervention_for_merge, function(x) sum(!is.na(x))),
  Missing_Values = sapply(intervention_for_merge, function(x) sum(is.na(x)))
)

intervention_summary_stats
  
all_summary_stats <- data.frame(
  Variable = names(merged),
  Valid_Cases = sapply(merged, function(x) sum(!is.na(x))),
  Missing_Values = sapply(merged, function(x) sum(is.na(x)))
)

all_summary_stats

#### More cleaning and tidying ####

cleaned <- merged %>%
  rename(ID = `...1`) %>%
                mutate(across(starts_with("Dates"), \(x) as.Date(x, format = "%d/%m/%Y"))) %>% #make dates into date format
  mutate(time_between_baseline_follow_up = `Dates ROMs_End of treatment` - `Dates ROMs_Baseline`)

date_analysis_DS <- cleaned %>%
  select(starts_with("Dates"), `time_between_baseline_follow_up`) 

date_analysis_DS

#### Loads of the dates are 0022 rather than 2022. Write a function to change that ####

fix_dates <- function(date_var) {
  text <- format(date_var, "%d/%m/%Y")
  # Define a regular expression pattern to match dates in the format DD/MM/0022
  pattern <- "(\\d{2}/\\d{2}/)0022(\\b)"
  
  # Replace '0022' with '2022' using regular expression substitution
  corrected_text <- gsub(pattern, "\\12022\\2", text)
  return(corrected_text)
}
#### Apply function ####
cleaned_dates <- cleaned %>%
  mutate(across(starts_with("Dates"), fix_dates)) %>%
  mutate(across(starts_with("Dates"), as.Date, format = "%d/%m/%Y")) %>%
  mutate(time_between_baseline_follow_up = `Dates ROMs_End of treatment` - `Dates ROMs_Baseline`)

cleaned_dates <- cleaned %>%
  mutate(across(starts_with("Dates"), fix_dates))

date_analysis_DS_2 <- cleaned_dates %>%
  select(starts_with("Dates"), `time_between_baseline_follow_up`) 

print(n= 200,date_analysis_DS_2)

##### Save wide dataset ####

wide_analysis_dataset <- cleaned_dates %>%
  select(ID, Intervention_Control, project_name, client_sex_at_birth, client_age_when_first_seen_years, `Presenting issue`, starts_with("Dates"), `RCADS_1_TOT (baseline)`, `RCADS 2_TOT (midpoint)`, `RCADS 3_TOT (end of treatment)`, `YPCORE10 (baseline)`, `YPCORE10 (midpoint)`, `YPCORE10 (end of treatment)`, time_between_baseline_follow_up, `N. of sessions`) 

write_csv(wide_analysis_dataset, "wide_analysis_dataset.csv")

##### Pivot longer #####


analysis_dataset_RCAD <- cleaned_dates %>%
  select(ID, Intervention_Control, `Presenting issue`, `N. of sessions`,
         project_name, client_sex_at_birth, client_age_when_first_seen_years, starts_with("Dates"), `RCADS_1_TOT (baseline)`, `RCADS 2_TOT (midpoint)`, `RCADS 3_TOT (end of treatment)`, `YPCORE10 (baseline)`, `YPCORE10 (midpoint)`, `YPCORE10 (end of treatment)`, time_between_baseline_follow_up, `N. of sessions`) %>%
  pivot_longer(cols = starts_with("RCADS"),
               names_to = "period",
               names_pattern =  ".*\\((.*?)\\).*",
               values_to = "RCADS") 

analysis_dataset_YPCORE <- cleaned_dates %>%
  select(ID, Intervention_Control, `Presenting issue`, `N. of sessions`, project_name, client_sex_at_birth, client_age_when_first_seen_years, starts_with("Dates"), `YPCORE10 (baseline)`, `YPCORE10 (midpoint)`, `YPCORE10 (end of treatment)`, time_between_baseline_follow_up, `N. of sessions`) %>%
  pivot_longer(cols = starts_with("YPCORE"),
               names_to = "period",
               names_pattern =  ".*\\((.*?)\\).*",
               values_to = "YPCORE")

analysis_dataset <- analysis_dataset_RCAD %>%
  left_join(analysis_dataset_YPCORE)

write_csv(analysis_dataset, "analysis_dataset.csv")

