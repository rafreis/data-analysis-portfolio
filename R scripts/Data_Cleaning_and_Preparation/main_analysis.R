### RCT Analysis ###
#Data analysis - mixed model #
#Analysis conducted by Paul Chappell in freelance capacity#

rm(list=ls()) # clear environment

##### Install packages:
#install.packages("Matrix")
#install.packages("lme4")
#install.packages("emmeans")

##### load packages:
library(tidyverse)
library(lme4)
library(emmeans)

#### Mixed model - baseline / follow-up only.Loop for both outcome variables ####

# Read in data
analysis_dataset <- read.csv("analysis_dataset.csv") 

# Define a function to extract fixed effects
getFixedEffects <- function(model) {
  fixef(model)
}

# Define outcome variables
outcome_variables <- c("YPCORE", "RCADS")

# Lists to store different parts of output comparisons:
mixed_model_summary_list <- list()
p_value_list <- list()
pairwise_comparison_list <- list()
EMMs_list <- list()

# Loop over outcome variables
for (outcome_var in outcome_variables) {
  cat("Outcome Variable:", outcome_var, "\n")
  
  # Filter data for baseline and end of treatment
  baseline_follow_up_only <- analysis_dataset %>%
    filter(period %in% c("baseline", "end of treatment"))
  
  # Fit mixed model
  mixed_model <- lmer(as.formula(paste(outcome_var, "~ Intervention_Control * period + (1 | ID)")), data = baseline_follow_up_only)
  
  # Save pairwise comparisons in list
  mixed_model_summary_list[[outcome_var]] <- mixed_model
  
  # Perform bootstrapping
  boot_results <- bootMer(mixed_model, FUN = getFixedEffects, nsim = 1000)
  
  # Extract bootstrapped estimates of fixed effects
  boot_fixed_effects <- boot_results$t
  
  # Compute p-values from bootstrapped estimates
  p_values <- sapply(1:ncol(boot_fixed_effects), function(i) {
    p <- sum(abs(boot_fixed_effects[, i]) >= abs(fixef(mixed_model)[i])) / length(boot_fixed_effects[, i])
    p
  })
  
  # Print p-values
  cat("P-values:\n")
  p_value_list[[outcome_var]] <- p_values

  # Obtain estimated marginal means (EMMs)
  emm <- emmeans(mixed_model, ~ Intervention_Control * period)
  
  #Save EMMs in list:
  EMMs_list[[outcome_var]] <- emm
  
  # Conduct pairwise comparisons
  pairwise_comparisons <- pairs(emm)
  
  # Save pairwise comparisons in list
  pairwise_comparison_list[[outcome_var]] <- pairwise_comparisons
}

# Display summary of mixed model for each outcome variable
print(mixed_model_summary_list[["YPCORE"]])
print(mixed_model_summary_list[["RCADS"]])

# Display p value list for each outcome variable
print(p_value_list[["YPCORE"]])
print(p_value_list[["RCADS"]])

# Display EMMs for each outcome variable
print(EMMs_list[["YPCORE"]])
print(EMMs_list[["RCADS"]])

# Access pairwise comparisons for each outcome variable
print(pairwise_comparison_list[["YPCORE"]])
print(pairwise_comparison_list[["RCADS"]])

### Plot EMMs with CIs 

EMM_dataset_YPCORE <- as.data.frame(EMMs_list[["YPCORE"]])
EMM_dataset_RCADS <- as.data.frame(EMMs_list[["RCADS"]])

ggplot(EMM_dataset_YPCORE, aes(x=period, y=emmean, colour =Intervention_Control, group = Intervention_Control)) +
  geom_point(size = 3) +
  geom_line(aes(group = Intervention_Control), size = 1.5) +
  geom_errorbar(aes(ymin = lower.CL, ymax = upper.CL), width = 0.2, size = 1.5) +
  labs(
    x = "Time",
    y = "YPCORE",
    colour = "Condition"
  ) +
  theme_minimal() +
  theme(legend.position = "bottom")
  

ggplot(EMM_dataset_RCADS, aes(x=period, y=emmean, colour =Intervention_Control, group = Intervention_Control)) +
  geom_point(size = 3) +
  geom_line(aes(group = Intervention_Control), size = 1.5) +
  geom_errorbar(aes(ymin = lower.CL, ymax = upper.CL), width = 0.2, size = 1.5) +
  labs(
    x = "Time",
    y = "RCADS",
    colour = "Condition"
  ) +
  theme_minimal() +
  theme(legend.position = "bottom")


