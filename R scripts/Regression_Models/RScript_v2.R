library(openxlsx)
library(dplyr)
library(tidyr)
library(readr)
library(readxl)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/steve_0139")

df <- read.xlsx("00_Dataset_Regression_Final.xlsx")
df$CMO_Tenure <- as.numeric(df$CMO_Tenure)

## Group by Name

library(dplyr)

df$NewProductIntroduction <- as.numeric(df$NewProductIntroduction)

df_aggregated <- df %>%
  group_by(Name_Compustat) %>%
  summarise(
    mean_Liberalism_Index = mean(Liberalism_Index, na.rm = TRUE),
    mean_AD_Intensity = mean(AD_Intensity, na.rm = TRUE),
    mean_NewProductIntroduction = mean(NewProductIntroduction, na.rm = TRUE),
    mean_Total_Assets = mean(Total_Assets, na.rm = TRUE),
    mean_ROA = mean(ROA, na.rm = TRUE)
  )


summary_stats <- df_aggregated %>%
  summarise(
    mean_mean_Liberalism_Index = mean(mean_Liberalism_Index, na.rm = TRUE),
    sd_mean_Liberalism_Index = sd(mean_Liberalism_Index, na.rm = TRUE),
    mean_mean_AD_Intensity = mean(mean_AD_Intensity, na.rm = TRUE),
    sd_mean_AD_Intensity = sd(mean_AD_Intensity, na.rm = TRUE),
    mean_mean_NewProductIntroduction = mean(mean_NewProductIntroduction, na.rm = TRUE),
    sd_mean_NewProductIntroduction = sd(mean_NewProductIntroduction, na.rm = TRUE),
    mean_mean_Total_Assets = mean(mean_Total_Assets, na.rm = TRUE),
    sd_mean_Total_Assets = sd(mean_Total_Assets, na.rm = TRUE),
    mean_mean_ROA = mean(mean_ROA, na.rm = TRUE),
    sd_mean_ROA = sd(mean_ROA, na.rm = TRUE)
  )


# Convert the summary_stats to a long format
long_format <- tidyr::pivot_longer(summary_stats, everything(), names_to = "Variable", values_to = "Value")

# Separate the statistics (mean, sd) from the Variable names
long_format <- long_format %>%
  tidyr::separate(Variable, into = c("Statistic", "Variable"), sep = "_", extra = "merge") %>%
  dplyr::mutate(Variable = paste(Variable, sep = "_"))

# Pivot back to a wide format to get a column each for mean and standard deviation
reshaped_stats <- long_format %>%
  tidyr::pivot_wider(names_from = Statistic, values_from = Value)

library(dplyr)
library(tidyr)

# Select one value for each Name_Compustat
reduced_df <- df %>%
  group_by(Name_Compustat) %>%
  slice_head(n = 1) %>%
  ungroup()

# Create a crosstab for Liberalism_Label with counts
crosstab_liberalism <- table(reduced_df$Liberalism_Label)
crosstab_liberalism_df <- as.data.frame(crosstab_liberalism)
names(crosstab_liberalism_df) <- c("Liberalism_Label", "Freq")
total_counts_liberalism <- sum(crosstab_liberalism)
crosstab_liberalism_df$Percentage <- (crosstab_liberalism_df$Freq / total_counts_liberalism) * 100

# Create a crosstab for SIC with counts
crosstab_sic <- table(df$SIC)
crosstab_sic_df <- as.data.frame(crosstab_sic)
names(crosstab_sic_df) <- c("SIC", "Freq")
total_counts_sic <- sum(crosstab_sic)
crosstab_sic_df$Percentage <- (crosstab_sic_df$Freq / total_counts_sic) * 100

# Create a crosstab for Gender with counts
crosstab_Gender <- table(reduced_df$Is_Male)
crosstab_Gender_df <- as.data.frame(crosstab_Gender)
names(crosstab_Gender_df) <- c("Gender", "Freq")
total_counts_Gender <- sum(crosstab_Gender)
crosstab_Gender_df$Percentage <- (crosstab_Gender_df$Freq / total_counts_Gender) * 100


## OUTLIER INSPECTION

# Define the variables
vs <- c("Liberalism_Index", "AD_Intensity", "NewProductIntroduction", "Total_Assets", "ROA")

# Function to calculate Z-scores and identify outliers
find_outliers <- function(data, vars, threshold = 3) {
  outlier_df <- data.frame(Variable = character(), Outlier_Count = integer())
  outlier_values_df <- data.frame()
  
  for (var in vars) {
    x <- data[[var]]
    z_scores <- (x - mean(x, na.rm = TRUE)) / sd(x, na.rm = TRUE)
    
    # Identify outliers based on the Z-score threshold
    outliers <- which(abs(z_scores) > threshold)
    
    # Count outliers
    count_outliers <- length(outliers)
    
    # Store results in a summary dataframe
    outlier_df <- rbind(outlier_df, data.frame(Variable = var, Outlier_Count = count_outliers))
    
    # Store outlier values and Z-scores in a separate dataframe
    if (count_outliers > 0) {
      tmp_df <- data.frame(data[outliers, ], Z_Score = z_scores[outliers])
      tmp_df$Variable <- var
      outlier_values_df <- rbind(outlier_values_df, tmp_df)
      
      print(paste("Variable:", var, "has", count_outliers, "outliers with Z-scores greater than", threshold))
    }
  }
  
  return(list(summary = outlier_df, values = outlier_values_df))
}

# Execute the function
outlier_summary <- find_outliers(df, vs, threshold = 3)

# Splitting the list into two separate dataframes
outlier_count <- outlier_summary$summary
outlier_values_df <- outlier_summary$values

## Calculate Skewness and Kurtosis

library(moments)
library(stats)

calc_descriptive_stats_for_dvs <- function(data, dvs) {
  stats_df <- data.frame(Variable = character(), Skewness = numeric(), 
                         Kurtosis = numeric(), W = numeric(), P_Value = numeric())
  
  for (var in dvs) {
    x <- data[[var]]
    x <- x[!is.na(x)]  # Remove NA values if present
    
    if (length(x) < 3) {
      next  # Skip if less than 3 observations
    }
    
    skewness_value <- skewness(x)
    kurtosis_value <- kurtosis(x)
    shapiro_test <- shapiro.test(x)
    W_value <- shapiro_test$statistic
    p_value <- shapiro_test$p.value
    
    stats_df <- rbind(stats_df, data.frame(Variable = var, Skewness = skewness_value, 
                                           Kurtosis = kurtosis_value, W = W_value, P_Value = p_value))
  }
  
  return(stats_df)
}

# Use function to get stats for all dvs

normality_stats_df <- calc_descriptive_stats_for_dvs(df, vs)

# Replace 'NaN' with NA in the 'ROA' column
df <- df %>% mutate(ROA = ifelse(ROA == 'NaN', NA, ROA))

# Log-transform all variables except Liberalism_Index
df_transformed <- df %>%
  dplyr::mutate_at(vars(AD_Intensity, NewProductIntroduction, Total_Assets), ~log(. + 1))


## ADJUST SIC CATEGORIES FOR AD INTENSITY MODEL

# Create a copy of df_transformed for AD Intensity model
df_transformed_adint <- df_transformed

# Count the number of observations for each SIC category
sic_counts <- table(df_transformed_adint$SIC)

# Identify categories with fewer than 25 observations
low_count_sic <- names(sic_counts[sic_counts < 25])

# Replace low-count SIC categories with 'Other' in df_transformed_adint
df_transformed_adint$SIC <- ifelse(df_transformed_adint$SIC %in% low_count_sic, 'Other', df_transformed_adint$SIC)

# Create dummy variables
dummy_vars <- model.matrix(~ SIC - 1, data = df_transformed_adint)
colnames(dummy_vars) <- paste("SIC_", colnames(dummy_vars), sep = "")

# Add the dummy variables to df_transformed_adint
df_transformed_adint <- cbind(df_transformed_adint, dummy_vars)


# Check the column names to ensure 'Other' is properly represented
print(colnames(dummy_vars))


## MODEL AD INTENSITY

# Load the lme4 package
library(lme4)

# Create the formula for the LMM model
formula <- as.formula("AD_Intensity ~ Liberalism_Index + RD_Intensity + Total_Assets + ROA + SIC_SIC48 + SIC_SIC73 + Is_Male + CMO_Tenure + Age_fyear + (1|Name_Compustat)")

# Fit the linear mixed-effects model with NA handling
lmer_model <- lmer(formula, data = df_transformed_adint)

# Use the base summary() function to obtain p-values
lmer_summary <- summary(lmer_model)

lmer_summary

# Obtain the ANOVA table with p-values
anova_table <- anova(lmer_model)

library(afex)

# Obtain p-values using the afex package
mixed_model <- mixed(lmer_model, data = df_transformed_adint, method = "KR")  # Use the Kenward-Roger method
summary(mixed_model)

# Residual Plots for AD Intensity
residuals <- residuals(lmer_model)
fitted_values <- fitted(lmer_model)

plot(fitted_values, residuals, main="Residuals vs Fitted", xlab="Fitted values", ylab="Residuals")
qqnorm(residuals, main="Q-Q Plot")

# MODEL PRODUCT INTRODUCTION

## ADJUST SIC CATEGORIES FOR NEW PRODUCT MODEL

# Declare the variables that are part of the formula
formula_vars <- c("NewProductIntroduction", "Liberalism_Index", "RD_Intensity", 
                  "Total_Assets", "ROA", "SIC",
                  "Is_Male", "Birthyear", "Age_fyear", "Name_Compustat")

# Remove rows where these specific variables have missing data
df_transformed_newproduct <- df_transformed[complete.cases(df_transformed[, formula_vars]), ]

# Count the number of observations for each SIC category
sic_counts <- table(df_transformed_newproduct$SIC)

# Identify categories with fewer than 10 observations
low_count_sic <- names(sic_counts[sic_counts < 15])

# Replace low-count SIC categories with 'Other'
df_transformed_newproduct$SIC <- ifelse(df_transformed_newproduct$SIC %in% low_count_sic, 'Other', df_transformed_newproduct$SIC)

# Create dummy variables
dummy_vars <- model.matrix(~ SIC - 1, data = df_transformed_newproduct)
colnames(dummy_vars) <- paste("SIC_", colnames(dummy_vars), sep = "")
df_transformed_newproduct <- cbind(df_transformed_newproduct, dummy_vars)

# Check the column names to ensure 'Other' is properly represented
print(colnames(dummy_vars))

colnames(df_transformed_newproduct)

df_transformed_newproduct$Name_Compustat <- factor(df_transformed_newproduct$Name_Compustat)

# Update factor levels for Name_Compustat
df_transformed_newproduct$Name_Compustat <- factor(df_transformed_newproduct$Name_Compustat)

# Create the formula for the Negative Binomial model with NewProductIntroduction as DV
formula_new_product_nb <- as.formula("NewProductIntroduction ~ Liberalism_Index + RD_Intensity + Total_Assets + ROA + SIC_SIC48 + SIC_SIC73 + Is_Male + Birthyear + Age_fyear + (1|Name_Compustat)")

# Fit the generalized linear mixed-effects model with Negative Binomial distribution
nb_model_new_product <- glmer.nb(formula_new_product_nb, data = df_transformed_newproduct)

# Obtain the summary and ANOVA table
nb_summary_new_product <- summary(nb_model_new_product)
nb_anova_table_new_product <- anova(nb_model_new_product)

# Obtain p-values using the afex package
library(afex)
poisson_model_new_product <- mixed(nb_model_new_product, data = df_transformed_newproduct, method = "KR")  # Use the Kenward-Roger method
summary(poisson_model_new_product)

# Extract fixed-effects coefficients for AD Intensity model
fixed_effects_ad_intensity <- lmer_summary$coefficients[, "Estimate"]

# Extract fixed-effects coefficients for New Product Introduction model
fixed_effects_new_product <- summary(nb_model_new_product )$coefficients[, "Estimate"]

# Convert to data frames if you like
fixed_effects_ad_intensity_df <- as.data.frame(fixed_effects_ad_intensity)
fixed_effects_new_product_df <- as.data.frame(fixed_effects_new_product)

## DESCRIPTIVE STATISTICS

# Calculate means and standard deviations for numerical columns
num_cols <- names(df_aggregated)[sapply(df_aggregated, is.numeric)]
means <- sapply(df_aggregated[, num_cols], function(x) mean(x, na.rm = TRUE))
sds <- sapply(df_aggregated[, num_cols], function(x) sd(x, na.rm = TRUE))

# Create summary statistics dataframe
summary_stats_df <- data.frame(mean = means, sd = sds)

# Print the resulting dataframe
print(means_sds_df)

## EXPORT RESULTS


# Add a new column to include row names (index)
summary_stats_df$Variable = rownames(summary_stats_df)

# Reorder the columns to have 'Variable' as the first column
summary_stats_df = summary_stats_df[, c("Variable", "mean", "sd")]

anova_table$Variable = rownames(anova_table)
anova_table_new_product$Variable = rownames(anova_table_new_product)


library(writexl)

# Create an Excel workbook
write_xlsx(
  list(
    "Descriptive Statistics" = summary_stats_df,
    "AD_Intensity" = anova_table,
    "NewProductIntroduction" = anova_table_new_product,
    "Liberalism Crosstab" = crosstab_liberalism_df,
    "Normality Stats" = normality_stats_df
  ),
  path = "Results.xlsx"
)


