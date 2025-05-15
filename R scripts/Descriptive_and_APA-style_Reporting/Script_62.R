# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/yolobaer")

# IF CSV

# Read data from a CSV file
df <- read.csv("RGSRRPfinal (1).csv", sep = ";")

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

colnames(df)

RGSRRPfinal <- df

library(MatchIt)
library(survival)


# Radsumall to categorial variable
RGSRRPfinal$Rad_sum_allLN_Steinhelfer_recode <- cut(RGSRRPfinal$Rad_sum_allLN_Steinhelfer,
                                                    breaks = c(-Inf, 2, 4, Inf),
                                                    labels = c("1", "2", "3"),
                                                    right = FALSE)

RGSRRPfinal$Rad_sum_allLN_Steinhelfer_recode <- as.factor(RGSRRPfinal$Rad_sum_allLN_Steinhelfer_recode)

colnames(RGSRRPfinal)

# Exact Matching
matchit_model <- matchit(
  RGSjanein ~ RStat + RadoutsidePLND + Rad_sum_allLN_Steinhelfer_recode,
  method = "exact",
  data = RGSRRPfinal,
  x = TRUE,
  y = TRUE
)


# Matching summary
summary(matchit_model)

#Rename Matched data 
matched_data <- match.data(matchit_model)

#Censoring
matched_data$Time_to_Event_or_Censor <- pmin(matched_data$MonatebisTherapie1, matched_data$MonatebisFU)

matched_data$Event_or_Censor <- with(matched_data, ifelse(MonatebisTherapie1 <= MonatebisFU, Folgetherapiejanein, 0))

# Recreate Surv object safely
surv_obj <- with(matched_data, Surv(Time_to_Event_or_Censor, Event_or_Censor))

# Fit Cox model
cox_model <- coxph(surv_obj ~ RGSjanein + Rad_sum_allLN_Steinhelfer + factor(RadoutsidePLND) + factor(RStat),
                   data = matched_data)

# Summary
summary(cox_model)


# New part

library(survminer)

# Adjusted survival curves (Kaplan-Meier from Cox model)
fit <- survfit(cox_model, newdata = data.frame(
  RGSjanein = c(0, 1),
  Rad_sum_allLN_Steinhelfer = median(matched_data$Rad_sum_allLN_Steinhelfer, na.rm = TRUE),
  RadoutsidePLND = 0,  # reference value
  RStat = 1            # reference value
))

# Legend labels
group_labels <- c("Conventional", "PSMA-RGS")

# Plot the curves
ggsurvplot(
  fit,
  data = matched_data,
  legend.title = "Group",
  legend.labs = group_labels,
  risk.table = TRUE,
  xlab = "Months until start of next treatment",
  ylab = "Probability of remaining treatment-free",
  conf.int = TRUE,
  palette = "Dark2",
  ggtheme = theme_minimal(),
  xlim = c(0, 30)          # limits to 30 months as per request
)

# Calculate adjusted median survival time for each group
# Repeat for each group individually
newdata_conv <- data.frame(
  RGSjanein = 0,
  Rad_sum_allLN_Steinhelfer = median(matched_data$Rad_sum_allLN_Steinhelfer, na.rm = TRUE),
  RadoutsidePLND = 0,
  RStat = 1
)

newdata_rgs <- data.frame(
  RGSjanein = 1,
  Rad_sum_allLN_Steinhelfer = median(matched_data$Rad_sum_allLN_Steinhelfer, na.rm = TRUE),
  RadoutsidePLND = 0,
  RStat = 1
)

# Median survival time extraction
fit_conv <- survfit(cox_model, newdata = newdata_conv)
fit_rgs  <- survfit(cox_model, newdata = newdata_rgs)

# Extract median (time until 50% events)
summary(fit_conv)$table["median"]
summary(fit_rgs)$table["median"]



library(broom)
library(ggplot2)

# Tidy model
cox_tidy <- tidy(cox_model, exponentiate = TRUE, conf.int = TRUE)

# Plot
ggplot(cox_tidy, aes(x = term, y = estimate)) +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = conf.low, ymax = conf.high), width = 0.2) +
  geom_hline(yintercept = 1, linetype = "dashed", color = "gray") +
  labs(y = "Hazard Ratio (HR)", x = "Variable",
       title = "Hazard Ratios and 95% Confidence Intervals") +
  theme_minimal()              # limits to 30 months as per request


fit_na <- survfit(surv_obj ~ RGSjanein, data = matched_data)

ggsurvplot(
  fit_na,
  fun = "cumhaz",
  palette = "Dark2",
  legend.labs = c("Conventional", "PSMA-RGS"),
  legend.title = "Group",
  xlab = "Months",
  ylab = "Cumulative hazard",
  ggtheme = theme_minimal()
)


library(survminer)

timepoints <- c(6, 12, 24)

summary(fit_conv, times = timepoints)$surv
summary(fit_rgs, times = timepoints)$surv


fit <- survfit(surv_obj ~ RGSjanein, data = matched_data)

ggsurvplot(
  fit,
  data = matched_data,
  legend.title = "Group",
  legend.labs = group_labels,
  risk.table = "absolute",        # ensures the table is not combined
  risk.table.col = "strata",      # color-code by strata
  risk.table.height = 0.25,
  xlab = "Months until start of next treatment",
  ylab = "Probability of remaining treatment-free",
  conf.int = TRUE,
  xlim = c(0, 30),                # limits to 30 months as per request
  palette = "Dark2",
  ggtheme = theme_minimal()
)





colnames(matched_data)

# Create Surv object for bRFS
bRFS_obj <- with(matched_data, Surv(MonatebisBCR, BCRjanein))

# Fit Cox model for bRFS
bRFS_model <- coxph(bRFS_obj ~ RGSjanein + Rad_sum_allLN_Steinhelfer + factor(RadoutsidePLND) + factor(RStat),
                    data = matched_data)

# Summary
summary(bRFS_model)

# Kaplan-Meier (adjusted) for bRFS
fit_bRFS <- survfit(bRFS_model, newdata = data.frame(
  RGSjanein = c(0, 1),
  Rad_sum_allLN_Steinhelfer = median(matched_data$Rad_sum_allLN_Steinhelfer, na.rm = TRUE),
  RadoutsidePLND = 0,
  RStat = 1
))

# Plot adjusted bRFS survival curves
ggsurvplot(
  fit_bRFS,
  data = matched_data,
  legend.title = "Group",
  legend.labs = group_labels,
  risk.table = TRUE,
  xlab = "Months",
  ylab = "Probability of remaining bRFS",
  conf.int = TRUE,
  xlim = c(0, 30),
  palette = "Dark2",
  ggtheme = theme_minimal()
)


# Unadjusted Kaplan-Meier for bRFS stratified by group
fit_bRFS_strat <- survfit(Surv(MonatebisBCR, BCRjanein) ~ RGSjanein, data = matched_data)

# Plot with stratified number at risk
ggsurvplot(
  fit_bRFS_strat,
  data = matched_data,
  legend.title = "Group",
  legend.labs = group_labels,
  risk.table = "absolute",         # show separate rows
  risk.table.col = "strata",       # color by group
  risk.table.height = 0.25,
  xlab = "Months",
  ylab = "Probability of remaining bRFS",
  conf.int = TRUE,
  xlim = c(0, 30),
  palette = "Dark2",
  ggtheme = theme_minimal()
)

