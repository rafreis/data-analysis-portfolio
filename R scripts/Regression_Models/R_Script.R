library(psych)
library(ggplot2)
library(e1071)
library(moments)
library(dplyr)
library(tidyverse)
library(car)
library(stargazer)
library(emmeans)
library(flextable)
library(officer)

# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/oklmzer")

# Load the CSV file into a data frame
data <- read.csv("data_ToR.csv", sep = ";") 

perceived_humor_vars <- c("Q2_1", "Q2_2", "Q2_3")
brand_loyalty_vars <- c("Q3_1", "Q3_2", "Q3_3", "Q3_4", "Q3_5", "Q3_6")
purchase_intention_vars <- c("Q4_1", "Q4_2", "Q4_3")
need_for_humor_vars <- c("Q5_1", "Q5_2", "Q5_3", "Q5_4", "Q5_5", "Q5_6")
consumer_skepticism_vars <- c("Q6_1", "Q6_2", "Q6_3", "Q6_4", "Q6_5", "Q6_6", "Q6_7", "Q6_8", "Q6_9")

# Calculate the average of 'Brand Loyalty' questions and create a new column
data$Perceived_Humor_Avg <- rowMeans(data[c("Q2_1", "Q2_2", "Q2_3")], na.rm = TRUE)

# Calculate the average of 'Brand Loyalty' questions and create a new column
data$Brand_Loyalty_Avg <- rowMeans(data[c("Q3_1", "Q3_2", "Q3_3", "Q3_4", "Q3_5", "Q3_6")], na.rm = TRUE)

# Calculate the average of 'Purchase Intention' questions and create a new column
data$Purchase_Intention_Avg <- rowMeans(data[c("Q4_1", "Q4_2", "Q4_3")], na.rm = TRUE)

# Calculate the average of 'Need for Humor' questions and create a new column
data$Need_for_Humor_Avg <- rowMeans(data[c("Q5_1", "Q5_2", "Q5_3", "Q5_4", "Q5_5", "Q5_6")], na.rm = TRUE)

# Calculate the average of 'Consumer Skepticism About Ads' questions and create a new column
data$Consumer_Skepticism_Avg <- rowMeans(data[c("Q6_1", "Q6_2", "Q6_3", "Q6_4", "Q6_5", "Q6_6", "Q6_7", "Q6_8", "Q6_9")], na.rm = TRUE)

## SAMPLE CHARACTERIZATION ##

# Define a function to create a dataframe for each variable
create_df <- function(var, name) {
  df <- data.frame(Variable = name, 
                   Response = data[[var]], 
                   stringsAsFactors = FALSE)
  freq_table <- as.data.frame(table(df$Response))
  names(freq_table) <- c("Response", "Frequency")
  freq_table$Percentage <- (freq_table$Frequency / sum(freq_table$Frequency)) * 100
  freq_table$Variable <- name
  return(freq_table)
}

# Create dataframes for each variable
df_Q7 <- create_df("Q7", "Gender")
df_Q8 <- create_df("Q8", "Age Group")
df_Q9 <- create_df("Q9", "Social Status")

# Combine the dataframes
char_table <- rbind(df_Q7, df_Q8, df_Q9)

# Order the dataframe by variable name
char_table <- char_table[order(char_table$Variable), ]

# Change the column order to put "Variable" first
char_table <- char_table[, c("Variable", "Response", "Frequency", "Percentage")]

# Print the re-ordered table
print(char_table)

## DESCRIPTIVES AND RELIABILITY ANALYSIS ##

# Function to create a list of data for a scale
get_scale_data <- function(scale_vars, scale_name, avg_name) {
  scale_data <- lapply(scale_vars, function(var) {
    return(data.frame(
      Scale = scale_name,
      Variable = var,
      Mean = mean(data[[var]], na.rm = TRUE),
      SD = sd(data[[var]], na.rm = TRUE),
      Alpha = NA
    ))
  })
  
  # Add averaged scale data
  scale_data <- c(scale_data, list(data.frame(
    Scale = scale_name,
    Variable = avg_name,
    Mean = mean(data[[avg_name]], na.rm = TRUE),
    SD = sd(data[[avg_name]], na.rm = TRUE),
    Alpha = psych::alpha(data[scale_vars])$total$raw_alpha
  )))
  
  return(scale_data)
}

# Get data for each scale
perceived_humor_data <- get_scale_data(perceived_humor_vars, "Perceived Humor", "Perceived_Humor_Avg")
brand_loyalty_data <- get_scale_data(brand_loyalty_vars, "Brand Loyalty", "Brand_Loyalty_Avg")
purchase_intention_data <- get_scale_data(purchase_intention_vars, "Purchase Intention", "Purchase_Intention_Avg")
need_for_humor_data <- get_scale_data(need_for_humor_vars, "Need for Humor", "Need_for_Humor_Avg")
consumer_skepticism_data <- get_scale_data(consumer_skepticism_vars, "Consumer Skepticism About Ads", "Consumer_Skepticism_Avg")

# Combine all data
reliability_results <- do.call(rbind, c(perceived_humor_data, brand_loyalty_data, purchase_intention_data, need_for_humor_data, consumer_skepticism_data))

# Print results
print(reliability_results)

## BOXPLOTS AND HISTOGRAMS OF DV's ##

# Boxplot for each averaged scale
avg_data <- data.frame(
  Value = c(data$Perceived_Humor_Avg, data$Brand_Loyalty_Avg, data$Purchase_Intention_Avg, data$Need_for_Humor_Avg, data$Consumer_Skepticism_Avg),
  Scale = rep(c("Perceived Humor", "Brand Loyalty", "Purchase Intention", "Need for Humor", "Consumer Skepticism About Ads"), each = nrow(data))
)

ggplot(avg_data, aes(x = Scale, y = Value)) +
  geom_boxplot() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
  ylab("Average Scale Value") +
  ggtitle("Boxplot of Average Scale Values")

# Histograms for each averaged scale
par(mfrow = c(3, 2)) # 2x2 grid for the plots

hist(data$Perceived_Humor_Avg, main = "Perceived Humor", xlab = "Average Value")
hist(data$Brand_Loyalty_Avg, main = "Brand Loyalty", xlab = "Average Value")
hist(data$Purchase_Intention_Avg, main = "Purchase Intention", xlab = "Average Value")
hist(data$Need_for_Humor_Avg, main = "Need for Humor", xlab = "Average Value")
hist(data$Consumer_Skepticism_Avg, main = "Consumer Skepticism About Ads", xlab = "Average Value")

## NORMALITY CHECK ##

# Calculate skewness, kurtosis and Shapiro-Wilk test
calculate_stats <- function(scale_name, data_vec) {
  skew <- skewness(data_vec, na.rm = TRUE)
  kurt <- kurtosis(data_vec, na.rm = TRUE)
  shapiro_wilk <- shapiro.test(data_vec)$p.value
  
  return(data.frame(
    Scale = scale_name,
    Skewness = skew,
    Kurtosis = kurt,
    `Shapiro-Wilk p-value` = shapiro_wilk
  ))
}

# Get data for each scale
perceived_humor_stats <- calculate_stats("Perceived Humor", data$Perceived_Humor_Avg)
brand_loyalty_stats <- calculate_stats("Brand Loyalty", data$Brand_Loyalty_Avg)
purchase_intention_stats <- calculate_stats("Purchase Intention", data$Purchase_Intention_Avg)
need_for_humor_stats <- calculate_stats("Need for Humor", data$Need_for_Humor_Avg)
consumer_skepticism_stats <- calculate_stats("Consumer Skepticism About Ads", data$Consumer_Skepticism_Avg)

# Combine all data
normality_results <- rbind(perceived_humor_stats, brand_loyalty_stats, purchase_intention_stats, need_for_humor_stats, consumer_skepticism_stats)

# Print results
print(normality_results)

## DESCRIPTIVES BY GROUP ##

# Load required libraries
library(dplyr)
library(tidyverse)

# Recode 'GROUP_HUMOR' to 'Humor' and 'No Humor'
data <- data %>% mutate(GROUP_HUMOR = recode_factor(GROUP_HUMOR, `0` = "No Humor", `1` = "Humor"))

# Calculate means and standard deviations by 'GROUP_HUMOR'
descriptives <- data %>% 
  group_by(GROUP_HUMOR) %>%
  summarise(
    Perceived_Humor_Mean = mean(Perceived_Humor_Avg, na.rm = TRUE),
    Perceived_Humor_SD = sd(Perceived_Humor_Avg, na.rm = TRUE),
    Brand_Loyalty_Mean = mean(Brand_Loyalty_Avg, na.rm = TRUE),
    Brand_Loyalty_SD = sd(Brand_Loyalty_Avg, na.rm = TRUE),
    Purchase_Intention_Mean = mean(Purchase_Intention_Avg, na.rm = TRUE),
    Purchase_Intention_SD = sd(Purchase_Intention_Avg, na.rm = TRUE),
    Need_for_Humor_Mean = mean(Need_for_Humor_Avg, na.rm = TRUE),
    Need_for_Humor_SD = sd(Need_for_Humor_Avg, na.rm = TRUE),
    Consumer_Skepticism_Mean = mean(Consumer_Skepticism_Avg, na.rm = TRUE),
    Consumer_Skepticism_SD = sd(Consumer_Skepticism_Avg, na.rm = TRUE)
  )

# Print descriptives
print(descriptives)

# Pivot table so that 'Humor' and 'No Humor' are columns and statistics are rows
pivoted_descriptives <- descriptives %>%
  pivot_longer(-GROUP_HUMOR, names_to = "Statistic", values_to = "Value") %>%
  pivot_wider(names_from = GROUP_HUMOR, values_from = Value)

# Print pivoted descriptives
print(pivoted_descriptives)

# Recode GROUP_HUMOR as binary numeric variable
data$GROUP_HUMOR <- ifelse(data$GROUP_HUMOR == "Humor", 1, 0)

## REGRESSION ANALYSIS FOR BRAND LOYALTY ##

# Create interaction terms #
data$interaction_need_for_humor <- data$GROUP_HUMOR * data$Need_for_Humor_Avg
data$interaction_consumer_skepticism <- data$GROUP_HUMOR * data$Consumer_Skepticism_Avg

# Run Moderation Analysis

# Model 1: Only GROUP_HUMOR as the independent variable
model_1_brandloyalty <- lm(Brand_Loyalty_Avg ~ GROUP_HUMOR, data = data)
summary(model_1_brandloyalty)

# Model 2: Adding Consumer_Skepticism_Avg and Need_for_Humor_Avg
model_2_brandloyalty <- lm(Brand_Loyalty_Avg ~ GROUP_HUMOR + Need_for_Humor_Avg + Consumer_Skepticism_Avg, data = data)
summary(model_2_brandloyalty)

# Model 3: Adding interaction terms
data$interaction_need_for_humor <- data$GROUP_HUMOR * data$Need_for_Humor_Avg
data$interaction_consumer_skepticism <- data$GROUP_HUMOR * data$Consumer_Skepticism_Avg

model_3_brandloyalty <- lm(Brand_Loyalty_Avg ~ GROUP_HUMOR + Need_for_Humor_Avg + Consumer_Skepticism_Avg + interaction_need_for_humor + interaction_consumer_skepticism, data = data)
summary(model_3_brandloyalty)

# Residual Plot

qqnorm(model_3_brandloyalty$residuals)
qqline(model_3_brandloyalty$residuals)

# Multicolinearity Check

vif(model_2_brandloyalty)

## REGRESSION ANALYSIS FOR PURCHASE INTENTION ##

# Model 1: Only GROUP_HUMOR as the independent variable
model_1_purchaseintention <- lm(Purchase_Intention_Avg ~ GROUP_HUMOR, data = data)
summary(model_1_purchaseintention)

# Model 2: Adding Consumer_Skepticism_Avg and Need_for_Humor_Avg
model_2_purchaseintention <- lm(Purchase_Intention_Avg ~ GROUP_HUMOR + Need_for_Humor_Avg + Consumer_Skepticism_Avg, data = data)
summary(model_2_purchaseintention)

# Model 3: Adding interaction terms
data$interaction_need_for_humor_pi <- data$GROUP_HUMOR * data$Need_for_Humor_Avg
data$interaction_consumer_skepticism_pi <- data$GROUP_HUMOR * data$Consumer_Skepticism_Avg

model_3_purchaseintention <- lm(Purchase_Intention_Avg ~ GROUP_HUMOR + Need_for_Humor_Avg + Consumer_Skepticism_Avg + interaction_need_for_humor_pi + interaction_consumer_skepticism_pi, data = data)
summary(model_3_purchaseintention)

# Checking for assumptions

qqnorm(model_3_purchaseintention$residuals)
qqline(model_3_purchaseintention$residuals)

# Multicolinearity Check

vif(model_2_purchaseintention)

## ANOVA FOR BRAND LOYALTY ##

# Convert GROUP_HUMOR to a factor
data$GROUP_HUMOR <- as.factor(data$GROUP_HUMOR)

# Fit the ANOVA model
aov_model <- aov(Brand_Loyalty_Avg ~ GROUP_HUMOR, data = data)

# Conduct Levene's Test
car::leveneTest(Brand_Loyalty_Avg ~ GROUP_HUMOR, data = data)

# Print the summary of the model
summary(aov_model)

# Calculate the estimated marginal means
emm <- emmeans(aov_model, "GROUP_HUMOR")

# Convert to data frame for plotting
emm_df <- as.data.frame(summary(emm))

# Create plot with ggplot2
ggplot(emm_df, aes(x = GROUP_HUMOR, y = emmean)) +
  geom_point(size = 2) +
  geom_line(aes(group = 1), linetype = "dashed") +
  geom_errorbar(aes(ymin = emmean - SE, ymax = emmean + SE), width = 0.2) +
  labs(x = "Group Humor", y = "Estimated Marginal Mean of Brand Loyalty") +
   theme_bw()

## ANOVA FOR PURCHASE INTENTION ##

# Fit the ANOVA model
aov_model_PI <- aov(Purchase_Intention_Avg ~ GROUP_HUMOR, data = data)

# Convert GROUP_HUMOR to a factor
data$GROUP_HUMOR <- as.factor(data$GROUP_HUMOR)

# Conduct Levene's Test
car::leveneTest(Purchase_Intention_Avg ~ GROUP_HUMOR, data = data)

# Print the summary of the model
summary(aov_model_PI)

# Calculate the estimated marginal means
emm_pi <- emmeans(aov_model_PI, "GROUP_HUMOR")

# Convert to data frame for plotting
emm_df_pi <- as.data.frame(summary(emm_pi))

# Create plot with ggplot2
ggplot(emm_df_pi, aes(x = GROUP_HUMOR, y = emmean)) +
  geom_point(size = 2) +
  geom_line(aes(group = 1), linetype = "dashed") +
  geom_errorbar(aes(ymin = emmean - SE, ymax = emmean + SE), width = 0.2) +
  labs(x = "Group Humor", y = "Estimated Marginal Mean of Purchase Intention") +
  scale_x_discrete(limits = c("0", "1")) +
  theme_bw()

# Exporting Regression outputs

#brand Loyalty
stargazer(model_1_brandloyalty, model_2_brandloyalty, model_3_brandloyalty,
          type = "text",  # use "latex" for LaTeX output, "html" for HTML output
          title = "Regression Results",
          model.names = FALSE,
          align = TRUE,
          out = "model_brand.doc", # specify filename to output to Word
          column.labels = c("Model 1", "Model 2", "Model 3"), # Label the columns
          covariate.labels = c("Humor (Condition)", "Need for Humor", "Consumer Skepticism", "Humor x Need for Humor", "Humor x Consumer Skepticism"), # Label the covariates
          dep.var.caption = "Dependent variable: Brand Loyalty",
          dep.var.labels.include = FALSE,
          single.row = TRUE)  # include each covariate on a single row

#Purchase Intention
stargazer(model_1_purchaseintention, model_2_purchaseintention, model_3_purchaseintention,
          type = "text",  # use "latex" for LaTeX output, "html" for HTML output
          title = "Regression Results",
          model.names = FALSE,
          align = TRUE,
          out = "model_purchint.doc", # specify filename to output to Word
          column.labels = c("Model 1", "Model 2", "Model 3"), # Label the columns
          covariate.labels = c("Humor (Condition)", "Need for Humor", "Consumer Skepticism", "Humor x Need for Humor", "Humor x Consumer Skepticism"), # Label the covariates
          dep.var.caption = "Dependent variable: Purchase Intention",
          dep.var.labels.include = FALSE,
          single.row = TRUE)  # include each covariate on a single row

## Exporting descriptive tables and ANOVAs

# Generate flextable for char_table
char_table_ft <- flextable(char_table)
char_table_ft <- set_caption(char_table_ft, caption = "Characterization Table")
char_table_ft <- fontsize(char_table_ft, size = 12)
char_table_ft <- colformat_num(char_table_ft, digits = 3)

# Generate flextable for reliability_results
reliability_results_ft <- flextable(reliability_results)
reliability_results_ft <- set_caption(reliability_results_ft, caption = "Reliability Analysis Results")

# Generate flextable for normality_results
normality_results_ft <- flextable(normality_results)
normality_results_ft <- normality_results_ft %>%
  colformat_num(columns = c("Skewness", "Kurtosis", "Shapiro.Wilk.p.value"), digits = 3)
normality_results_ft <- set_caption(normality_results_ft, caption = "Normality Test Results")

# Generate flextable for pivoted_descriptives
pivoted_descriptives_ft <- flextable(pivoted_descriptives)
pivoted_descriptives_ft <- set_caption(pivoted_descriptives_ft, caption = "Descriptive Statistics by Group")

# Generate flextable for ANOVAs
aov_model_tidy <- broom::tidy(aov_model)
aov_model_PI_tidy <- broom::tidy(aov_model_PI)

aov_model_ft <- flextable(aov_model_tidy)
aov_model_ft <- set_caption(aov_model_ft, caption = "ANOVA Results - Brand Loyalty")

aov_model_pi_ft <- flextable(aov_model_PI_tidy )
aov_model_pi_ft <- set_caption(aov_model_pi_ft, caption = "ANOVA Results - Purchase Intention")


# Create a new Word document
doc <- read_docx()

# Add flextables to the Word document
doc <- body_add_flextable(doc, value = char_table_ft) %>% 
  body_add_break() %>% 
  body_add_flextable(value = reliability_results_ft) %>% 
  body_add_break() %>% 
  body_add_flextable(value = normality_results_ft) %>% 
  body_add_break() %>% 
  body_add_flextable(value = pivoted_descriptives_ft) %>% 
  body_add_break() %>% 
  body_add_flextable(value = aov_model_ft) %>% 
  body_add_break() %>% 
  body_add_flextable(value = aov_model_pi_ft)

# Save the Word document
print(doc, target = "results_tables.docx")

# Export to Excel

# Combine them into a list
list_of_tables <- list("Characterization Table" = char_table, 
                       "Reliability Results" = reliability_results, 
                       "Normality Results" = normality_results, 
                       "Pivoted Descriptives" = pivoted_descriptives,
                       "AOV Model" = aov_model_tidy,
                       "AOV Model PI" = aov_model_PI_tidy)

# Export to Excel
openxlsx::write.xlsx(list_of_tables, file = "results_tables.xlsx")
