# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/njlbaarcelona")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("2025.03.19 - Nuno Lopes - DATA.SET.xlsx", sheet = "Planilha1")

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


# Load necessary packages
library(lme4)
library(mediation)
library(dplyr)
library(lmerTest)
library(boot)
library(broom.mixed)

# Filter dyads
df_dyads <- df %>% filter(dm == 0)

# Center variables
df_dyads <- df_dyads %>%
  mutate(
    value_c = scale(valueofselected, scale = FALSE),
    likeboth_c = scale(likeboth, scale = FALSE),
    intvl = value_c * likeboth_c
  )


# Switching to Bayesian Model for Bootstrapping coefficients

library(brms)
library(tidybayes)
library(dplyr)
library(bayestestR)


# Define the mediator model (a-path with moderation)
med_model <- bf(satisfaction ~ value_c * likeboth_c + (1 | iddyad))

# Define the outcome model (b-path and direct effect)
out_model <- bf(purchaseintention ~ satisfaction + value_c + (1 | iddyad))


# Run the joint model (3250 - 750) × 4 = 2500 × 4 = 10,000 posterior samples
fit_model7 <- brm(
  med_model + out_model + set_rescor(FALSE),
  data = df_dyads,
  chains = 4,
  iter = 3250,
  warmup = 750,
  cores = 4,
  seed = 123
)

summary(fit_model7)


library(dplyr)
library(tidybayes)
library(posterior)

# Step 1: Extract posterior samples as a data frame
posterior_samples <- as_draws_df(fit_model7)

# Step 2: Rename relevant columns for clarity
post <- posterior_samples %>%
  as.data.frame() %>%
  transmute(
    a1 = `b_satisfaction_value_c`,
    a3 = `b_satisfaction_value_c:likeboth_c`,
    b  = `b_purchaseintention_satisfaction`
  )

# Step 3: Get SD of likeboth_c from original data
likeboth_sd <- sd(df_dyads$likeboth_c, na.rm = TRUE)
likeboth_vals <- c(-1, 0, 1) * likeboth_sd  # moderator values: -1SD, Mean, +1SD

# Step 4: Compute conditional indirect effects for each moderator level
indirect_effects <- sapply(likeboth_vals, function(w) {
  (post$a1 + post$a3 * w) * post$b
})

# Step 5: Convert to data frame
indirect_df <- as.data.frame(indirect_effects)
names(indirect_df) <- c("-1SD", "Mean", "+1SD")

# Step 6: Compute Index of Moderated Mediation (IMM)
indirect_df$IMM <- post$a3 * post$b



# Summarize results (posterior mean and 95% credible intervals)
summarize_effects <- function(vec) {
  c(
    Estimate = mean(vec),
    CI_lower = quantile(vec, 0.025),
    CI_upper = quantile(vec, 0.975)
  )
}

summary_table <- as.data.frame(t(apply(indirect_df, 2, summarize_effects)))
summary_table$Effect <- rownames(summary_table)
rownames(summary_table) <- NULL

# Order output nicely
# Convert to tibble before using select
# Rename CI columns for clean selection
colnames(summary_table) <- c("Estimate", "CI_lower", "CI_upper", "Effect")

# Reorder and display
summary_table <- summary_table %>%
  dplyr::select(Effect, Estimate, CI_lower, CI_upper)

# Optional: round for readability
summary_table <- summary_table %>%
  mutate(across(where(is.numeric), ~ round(.x, 3)))

summary_table

summary(fit_model7)



# Hypothesis 2

df_individuals <- df %>% 
  filter(sb == 0) %>%
  mutate(
    value_c = scale(valueofselected, scale = FALSE),
    satisfaction_c = scale(satisfaction, scale = FALSE)
  )


# Mediator model (a-path)
med_model2 <- bf(satisfaction_c ~ value_c + (1 | iddyad))

# Outcome model (b-path + direct effect)
out_model2 <- bf(purchaseintention ~ satisfaction_c + value_c + (1 | iddyad))


fit_model4 <- brm(
  med_model2 + out_model2 + set_rescor(FALSE),
  data = df_individuals,
  chains = 4,
  iter = 3250,
  warmup = 750,
  cores = 4,
  seed = 456  # different seed just to separate from model 7
)


summary(fit_model4)


# Extract posterior draws
post2 <- as_draws_df(fit_model4) %>%
  as.data.frame() %>%
  transmute(
    a = `b_satisfactionc_value_c`,
    b = `b_purchaseintention_satisfaction_c`,
    direct = `b_purchaseintention_value_c`
  ) %>%
  mutate(
    indirect = a * b,
    total = indirect + direct
  )

# Summarize
summarize_effects <- function(vec) {
  c(
    Estimate = mean(vec),
    CI_lower = quantile(vec, 0.025),
    CI_upper = quantile(vec, 0.975)
  )
}

summary_table2 <- as.data.frame(t(apply(post2[, c("a", "b", "indirect", "direct", "total")], 2, summarize_effects)))
summary_table2$Effect <- rownames(summary_table2)

colnames(summary_table2) <- c("Estimate", "CI_lower", "CI_upper", "Effect")

rownames(summary_table2) <- NULL

summary_table2 <- summary_table2 %>%
  dplyr::select(Effect, Estimate, CI_lower, CI_upper) %>%
  dplyr::mutate(across(where(is.numeric), ~ round(.x, 3)))

summary_table2

# New Hypothesis 1'

library(brms)
library(dplyr)
library(tidybayes)
library(posterior)

library(brms)


# Center variables
df <- df %>%
  mutate(
    value_c = scale(valueofselected, scale = FALSE),
    likeboth_c = scale(likeboth, scale = FALSE),
    intvl = value_c * likeboth_c
  )


df_a_path <- df %>% filter(sb == 1, dm %in% c(1, 2))

model_a <- brm(
  formula = satisfaction ~ value_c * likeboth_c + consumemovies + extroverted + female + age + (1 | iddyad / idsubject),
  data = df_a_path,
  chains = 4,
  iter = 3250,
  warmup = 750,
  cores = 4,
  seed = 123
)

summary(model_a)

df_b_path <- df %>% filter(dm == 0)

model_b <- brm(
  formula = purchaseintention ~ satisfaction + value_c + consumemovies + extroverted + female + age + (1 | iddyad),
  data = df_b_path,
  chains = 4,
  iter = 3250,
  warmup = 750,
  cores = 4,
  seed = 123
)

summary(model_b)


model_total <- brm(
  formula = purchaseintention ~ value_c + consumemovies + extroverted + female + age + (1 | iddyad),
  data = df_b_path,
  chains = 4,
  iter = 3250,
  warmup = 750,
  cores = 4,
  seed = 123
)

summary(model_total)



