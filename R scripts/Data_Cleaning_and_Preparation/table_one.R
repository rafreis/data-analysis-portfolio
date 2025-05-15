### RCT Analysis ###
#Data analysis - table one #

#### Load and install packages:

# install.packages("gtsummary")

library(tidyverse) 
library(gtsummary)

rm(list=ls()) # clear environment
 
###Read in datasets from 1.prepare_data.R ####

analysis_dataset <- read.csv("analysis_dataset.csv") 
wide_analysis_dataset <- read_csv("wide_analysis_dataset.csv")

#### Table one comparison of intervention to control ####

wide_analysis_dataset$Intervention_Control <- as.factor(wide_analysis_dataset$Intervention_Control)
wide_analysis_dataset$client_sex_at_birth <- as.factor(wide_analysis_dataset$client_sex_at_birth)


# Define participant characteristics variables
participant_characteristics <- c("client_sex_at_birth", 
                                 "client_age_when_first_seen_years", 
                                 "Presenting issue", 
                                 "N. of sessions")

  # Create a summary table using gtsummary
tbl <- wide_analysis_dataset %>% 
  select(Intervention_Control, all_of(participant_characteristics)) %>%
  tbl_summary(
    by = "Intervention_Control",
    label = list(client_age_when_first_seen_years ~ "Patient Age"),
    statistic = list(all_continuous() ~ "{mean} ({sd})"),
    digits = list(client_age_when_first_seen_years ~ c(0, 1))
  ) %>%
  add_p(test = list(all_continuous() ~ "t.test", all_categorical() ~ "chisq.test"), 
           pvalue_fun = style_pvalue)


# Print the table
tbl

# Convert tbl to a data frame
tbl_df <- as.data.frame(tbl)

# Save the data frame to a CSV file
write.csv(tbl_df, file = "table_one_intervention_control.csv", row.names = FALSE)


####create table 1 but where comparison is baseline to follow-up. Separately for RCADS and YPCORE ####

#Create datasets with each subsample (RCADS baseline, RCADS follow=up, YPCORE baseline, YPCORE follow-up)


baseline_only_RCADS <- wide_analysis_dataset  %>%
  select(ID, Intervention_Control, project_name, client_sex_at_birth, client_age_when_first_seen_years, `Presenting issue`, starts_with("Dates"), `RCADS_1_TOT (baseline)`, `RCADS 2_TOT (midpoint)`, `RCADS 3_TOT (end of treatment)`, `YPCORE10 (baseline)`, `YPCORE10 (midpoint)`, `YPCORE10 (end of treatment)`, time_between_baseline_follow_up, `N. of sessions`) %>%
  filter(!is.na(`RCADS_1_TOT (baseline)`))

follow_up_only_RCADS <- wide_analysis_dataset  %>%
  select(ID, Intervention_Control, project_name, client_sex_at_birth, client_age_when_first_seen_years, `Presenting issue`, starts_with("Dates"), `RCADS_1_TOT (baseline)`, `RCADS 2_TOT (midpoint)`, `RCADS 3_TOT (end of treatment)`, `YPCORE10 (baseline)`, `YPCORE10 (midpoint)`, `YPCORE10 (end of treatment)`, time_between_baseline_follow_up, `N. of sessions`) %>%
  filter(!is.na(`RCADS 3_TOT (end of treatment)`))

baseline_only_YPCORE <- wide_analysis_dataset  %>%
  select(ID, Intervention_Control, project_name, client_sex_at_birth, client_age_when_first_seen_years, `Presenting issue`, starts_with("Dates"), `RCADS_1_TOT (baseline)`, `RCADS 2_TOT (midpoint)`, `RCADS 3_TOT (end of treatment)`, `YPCORE10 (baseline)`, `YPCORE10 (midpoint)`, `YPCORE10 (end of treatment)`, time_between_baseline_follow_up, `N. of sessions`) %>%
  filter(!is.na(`YPCORE10 (baseline)`))

follow_up_only_YPCORE <- wide_analysis_dataset  %>%
  select(ID, Intervention_Control, project_name, client_sex_at_birth, client_age_when_first_seen_years, `Presenting issue`, starts_with("Dates"), `RCADS_1_TOT (baseline)`, `RCADS 2_TOT (midpoint)`, `RCADS 3_TOT (end of treatment)`, `YPCORE10 (baseline)`, `YPCORE10 (midpoint)`, `YPCORE10 (end of treatment)`, time_between_baseline_follow_up, `N. of sessions`) %>%
  filter(!is.na(`YPCORE10 (end of treatment)`))



tbl_baseline_RCADS <- baseline_only_RCADS %>% 
  select(Intervention_Control, all_of(participant_characteristics)) %>%
  tbl_summary(
    label = list(client_age_when_first_seen_years ~ "Patient Age"),
    statistic = list(all_continuous() ~ "{mean} ({sd})"),
    digits = list(client_age_when_first_seen_years ~ c(0, 1))
  ) 

tbl_follow_up_RCADS <- follow_up_only_RCADS %>% 
  select(Intervention_Control, all_of(participant_characteristics)) %>%
  tbl_summary(
    label = list(client_age_when_first_seen_years ~ "Patient Age"),
    statistic = list(all_continuous() ~ "{mean} ({sd})"),
    digits = list(client_age_when_first_seen_years ~ c(0, 1))
  ) 

tbl_baseline_YPCORE <- baseline_only_YPCORE %>% 
  select(Intervention_Control, all_of(participant_characteristics)) %>%
  tbl_summary(
    label = list(client_age_when_first_seen_years ~ "Patient Age"),
    statistic = list(all_continuous() ~ "{mean} ({sd})"),
    digits = list(client_age_when_first_seen_years ~ c(0, 1))
  ) 

tbl_follow_up_YPCORE <- follow_up_only_YPCORE %>% 
  select(Intervention_Control, all_of(participant_characteristics)) %>%
  tbl_summary(
    label = list(client_age_when_first_seen_years ~ "Patient Age"),
    statistic = list(all_continuous() ~ "{mean} ({sd})"),
    digits = list(client_age_when_first_seen_years ~ c(0, 1))
  ) 


# Convert tbls to a data frame
tbl_baseline_RCADS_df <- as.data.frame(tbl_baseline_RCADS)
tbl_follow_up_RCADS_df <- as.data.frame(tbl_follow_up_RCADS)
tbl_baseline_YPCORE <- as.data.frame(tbl_baseline_YPCORE)
tbl_follow_up_YPCORE <- as.data.frame(tbl_follow_up_YPCORE)

tbl_baseline_follow_up_RCADS_df <- tbl_baseline_RCADS_df %>%
  bind_cols(tbl_follow_up_RCADS_df)


tbl_baseline_follow_up_YPCORE <- tbl_baseline_YPCORE %>%
  bind_cols(tbl_follow_up_YPCORE)


tbl_baseline_follow_up_RCADS_df

tbl_baseline_follow_up_YPCORE
# Save the data frame to a CSV file
write.csv(tbl_baseline_follow_up_RCADS_df, file = "table_one_baseline_post_RCADS.csv", row.names = FALSE)

write.csv(tbl_baseline_follow_up_YPCORE, file = "table_one_baseline_post_YPCORE.csv", row.names = FALSE)

