setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/getsellgo/Two Question")

library(readxl)
df <- read_xlsx('WIPO_2Q_Survey_Data_051223.xlsx', sheet = '2Q DATA')

# Get rid of special characters
names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))


# New added Columns

# Calculating RATING_CHANGED
# 1 if the rate changed from before to after, 0 otherwise
df$RATING_CHANGED <- ifelse(df$RATING_BEFORE != df$RATING_AFTER, 1, 0)

# Calculating RATING_CHANGE
# Difference between RATING_AFTER and RATING_BEFORE
df$RATING_CHANGE <- df$RATING_AFTER - df$RATING_BEFORE

# Descriptive Statistics

library(dplyr)
library(tidyr)

# Calculate counts for each transition
transition_counts <- df %>%
  count(RATING_BEFORE, RATING_AFTER)

# Calculate total number of responses
total_responses <- sum(transition_counts$n)

# Create a transition matrix with percentages
transition_matrix <- transition_counts %>%
  mutate(percentage = n / total_responses * 100) %>%
  pivot_wider(names_from = RATING_AFTER, values_from = c(n, percentage), values_fill = list(n = 0, percentage = 0))

# View the transition matrix with counts and percentages
print(transition_matrix)

# VISUALIZATIONS

library(ggplot2)
library(reshape2)
library(dplyr)

# Heat Map
# Melting the transition matrix for percentage data
transition_matrix_melted <- melt(transition_matrix, id.vars = 'RATING_BEFORE', measure.vars = grep("percentage", names(transition_matrix), value = TRUE))

# Adjusting the y-axis labels
transition_matrix_melted$variable <- gsub("percentage_", "", transition_matrix_melted$variable)

# Plotting the heatmap with smoother color transitions
ggplot(transition_matrix_melted, aes(x = RATING_BEFORE, y = variable, fill = value)) +
  geom_tile() +
  geom_text(aes(label = sprintf("%.1f%%", value)), color = "black", size = 4, vjust = 1.5, fontface = "bold") +  # Make text labels bold
  scale_fill_gradient2(low = "white", high = "green", mid = "lightgreen", midpoint = 6) +  # Change low gradient to 0.6%
  labs(x = "Rating Before", y = "Rating After", fill = "Percentage (%)") +
  theme_minimal()


# Paired Bar Chart

# Calculating counts for RATING_BEFORE and RATING_AFTER
before_counts <- as.data.frame(table(df$RATING_BEFORE))
after_counts <- as.data.frame(table(df$RATING_AFTER))

# Renaming columns for clarity
names(before_counts) <- c("Rating", "Before_Survey")
names(after_counts) <- c("Rating", "After_Survey")

# Merging the counts
rating_counts <- merge(before_counts, after_counts, by = "Rating")

# Melting for ggplot
rating_counts_melted <- melt(rating_counts, id.vars = 'Rating')

total_sample <- sum(rating_counts_melted$value)

# Create the plot
p <- ggplot(rating_counts_melted, aes(x = Rating, y = value, fill = variable)) +
  geom_bar(stat = "identity", position = position_dodge()) +
  labs(x = "Rating", y = "Count", fill = "Time") +
  theme_minimal()

# Add a secondary axis for percentages and data labels
p + scale_y_continuous(sec.axis = sec_axis(~ . / total_sample * 100, name = "% of Total Sample")) +
  geom_text(aes(label = paste(round(value / total_sample * 100, 1), "%")), position = position_dodge(width = 0.9), vjust = 1.5)


# Sankey Diagram

library(networkD3)

# Preparing data for Sankey diagram
links <- as.data.frame(table(df$RATING_BEFORE, df$RATING_AFTER))
names(links) <- c("source", "target", "value")

# Adjusting the source and target to be numeric
links$source <- as.numeric(as.factor(links$source)) - 1
links$target <- as.numeric(as.factor(links$target)) - 1 + max(links$source) + 1

# Sort the unique ratings in ascending order
unique_ratings_before <- sort(unique(df$RATING_BEFORE))
unique_ratings_after <- sort(unique(df$RATING_AFTER))

# Nodes list with sorted ratings
nodes <- data.frame(name = c(as.character(unique_ratings_before), as.character(unique_ratings_after)))

# Plotting the Sankey diagram
sankeyNetwork(Links = links, Nodes = nodes, Source = "source", Target = "target", Value = "value", NodeID = "name")


# Boxplot for RATING_CHANGE by RATING_BEFORE
ggplot(df, aes(x=factor(RATING_BEFORE), y=RATING_CHANGE, fill=factor(RATING_BEFORE))) +
  geom_boxplot() +
  labs(title="Boxplot of Rating Change",
       x="Rating Before",
       y="Rating Change") +
  theme_minimal()


# Boxplot for the overall distribution of RATING_CHANGE with mean
ggplot(df, aes(x=factor(1), y=RATING_CHANGE)) +
  geom_boxplot(width=0.2, fill="skyblue") +  # Adjust the width as needed and color the box
  stat_summary(fun=mean, geom="point", shape=20, size=3, color="darkred", aes(group=1)) +  # Add a dot for the mean
  labs(title="Boxplot of Overall Rating Change",
       x="", 
       y="Rating Change") +
  theme_minimal() +
  theme(axis.title.x=element_blank(),
        axis.text.x=element_blank(),
        axis.ticks.x=element_blank())


# Boxplot for RATING_AFTER by RATING_BEFORE
ggplot(df, aes(x = factor(RATING_BEFORE), y = RATING_AFTER, fill = factor(RATING_BEFORE))) +
  geom_boxplot() +
  stat_summary(fun = mean, geom = "point", shape = 20, size = 3, color = "darkred", aes(group = 1)) +
  labs(title = "Boxplot of Rating After Survey",
       x = "Rating Before Survey",
       y = "Rating After Survey") +
  theme_minimal() +
  guides(fill = FALSE)  # Remove the legend for fill

# Boxplot for RATING_CHANGE by QUIZ_SCORE
ggplot(df, aes(x=factor(QUIZ_SCORE), y=RATING_CHANGE, fill=factor(QUIZ_SCORE))) +
  geom_boxplot() +
  stat_summary(fun=mean, geom="point", shape=20, size=3, color="darkred", aes(group=1)) +
  labs(title="Boxplot of Rating Change",
       x="Quiz Score",
       y="Rating Change") +
  theme_minimal()

# PAIRED BAR PLOT FOR RATINGS BY QUIZ SCORE

# Convert factors to characters if they are not already
df$RATING_BEFORE <- as.numeric(as.character(df$RATING_BEFORE))
df$RATING_AFTER <- as.numeric(as.character(df$RATING_AFTER))
df$QUIZ_SCORE <- as.factor(df$QUIZ_SCORE)

# Reshape the data for plotting
library(tidyr)
df_long <- gather(df, key = "Rating_Type", value = "Rating", RATING_BEFORE, RATING_AFTER)

# Calculate mean ratings for each QUIZ_SCORE and Rating_Type
mean_ratings <- aggregate(Rating ~ QUIZ_SCORE + Rating_Type, data = df_long, mean)

# Create the paired bar plot with means
ggplot(mean_ratings, aes(x = QUIZ_SCORE, y = Rating, fill = Rating_Type)) +
  geom_bar(stat = "identity", position = position_dodge()) +
  geom_text(aes(label = round(Rating, 2)), position = position_dodge(width = 0.9), vjust = -0.25) +
  theme_minimal() +
  labs(title = "Mean Ratings Before and After by Quiz Score",
       x = "Quiz Score",
       y = "Mean Rating",
       fill = "Rating Type")


## FREQUENCY ANALYSIS

library(dplyr)

# Frequency and percentage calculations for various variables
calculate_freq_percent <- function(data, variable_name) {
  df_freq <- as.data.frame(table(data))
  df_freq <- df_freq %>%
    mutate(Percentage = Freq / sum(Freq) * 100) %>%
    mutate(Variable = variable_name)
  
  return(df_freq)
}

df_freq_rating_changed <- calculate_freq_percent(df$RATING_CHANGED, "RATING_CHANGED")
df_freq_rating_change <- calculate_freq_percent(df$RATING_CHANGE, "RATING_CHANGE")
df_freq_ans_s <- calculate_freq_percent(df$ANS_S, "ANS_S")
df_freq_ans_d <- calculate_freq_percent(df$ANS_D, "ANS_D")
df_freq_ANS_P <- calculate_freq_percent(df$ANS_P, "ANS_P")
df_freq_quizscore <- calculate_freq_percent(df$QUIZ_SCORE, "Quiz_Score")

# Counts and percentages of negative, zero, and positive values in RATING_CHANGE
count_data <- data.frame(
  data = c("Negative", "Zero", "Positive"),
  Freq = c(sum(df$RATING_CHANGE < 0), sum(df$RATING_CHANGE == 0), sum(df$RATING_CHANGE > 0)),
  Percentage = c(sum(df$RATING_CHANGE < 0), sum(df$RATING_CHANGE == 0), sum(df$RATING_CHANGE > 0)) / nrow(df) * 100,
  Variable = "RATING_CHANGE"
)

# Calculate frequencies and percentages for RATING_BEFORE
df_freq_rating_before <- calculate_freq_percent(df$RATING_BEFORE, "RATING_BEFORE")

# Calculate frequencies and percentages for RATING_AFTER
df_freq_rating_after <- calculate_freq_percent(df$RATING_AFTER, "RATING_AFTER")

# Combine all dataframes into one
all_freq_data <- bind_rows(df_freq_rating_changed, df_freq_rating_change, df_freq_quizscore, df_freq_ans_s, df_freq_ans_d, df_freq_ANS_P, count_data, df_freq_rating_before, df_freq_rating_after)




# NORMALITY CHECK

# Create a histogram for RATING_BEFORE
ggplot(df, aes(x = RATING_BEFORE)) +
  geom_histogram(binwidth = 1, fill = "cyan", color = "black", alpha = 0.7) +
  labs(title = "Histogram of RATING_BEFORE",
       x = "RATING_BEFORE",
       y = "Frequency") +
  theme_minimal()

# Create a histogram for RATING_AFTER
ggplot(df, aes(x = RATING_AFTER)) +
  geom_histogram(binwidth = 1, fill = "magenta", color = "black", alpha = 0.7) +
  labs(title = "Histogram of RATING_AFTER",
       x = "RATING_AFTER",
       y = "Frequency") +
  theme_minimal()

## STATISTICAL ANALYSIS

# MEAN COMPARISON

# Calculate means
mean_before <- mean(df$RATING_BEFORE, na.rm = TRUE)
mean_after <- mean(df$RATING_AFTER, na.rm = TRUE)

# Calculate the percentage change in the mean
percentage_change <- ((mean_after - mean_before) / mean_before) * 100

# Calculate 95% confidence intervals for both means
conf_interval_before <- t.test(df$RATING_BEFORE, conf.level = 0.99)$conf.int
conf_interval_after <- t.test(df$RATING_AFTER, conf.level = 0.99)$conf.int

# Create a data frame for the means and confidence intervals
means_df <- data.frame(
  Timepoint = c("Before Questions", "After Questions"),
  Mean = c(mean_before, mean_after),
  Lower_CI = c(conf_interval_before[1], conf_interval_after[1]),
  Upper_CI = c(conf_interval_before[2], conf_interval_after[2]),
  Perc_Change = percentage_change
)

# Function to convert ratings to percentage
convert_to_percentage <- function(rating) {
  return ((rating - 1) / 4) * 100
}

# Apply the conversion to both 'RATING_BEFORE' and 'RATING_AFTER'
df$PERCENTAGE_BEFORE <- sapply(df$RATING_BEFORE, convert_to_percentage)
df$PERCENTAGE_AFTER <- sapply(df$RATING_AFTER, convert_to_percentage)

# Calculate means in percentage terms
mean_percent_before <- mean(df$PERCENTAGE_BEFORE, na.rm = TRUE)
mean_percent_after <- mean(df$PERCENTAGE_AFTER, na.rm = TRUE)

# Calculate the percentage change in the mean
percentage_change_percent <- ((mean_percent_after - mean_percent_before) / mean_percent_before) * 100

# Calculate 95% confidence intervals for both percentage means
conf_interval_percent_before <- t.test(df$PERCENTAGE_BEFORE, conf.level = 0.99)$conf.int
conf_interval_percent_after <- t.test(df$PERCENTAGE_AFTER, conf.level = 0.99)$conf.int

# Update the data frame for the means and confidence intervals
means_df$Mean_Percentage <- c(mean_percent_before, mean_percent_after)
means_df$Lower_CI_Percentage <- c(conf_interval_percent_before[1], conf_interval_percent_after[1])
means_df$Upper_CI_Percentage <- c(conf_interval_percent_before[2], conf_interval_percent_after[2])
means_df$Perc_Change_Percentage <- percentage_change_percent

# Print the results
print(means_df)
print(paste("Percentage change in the mean: ", round(percentage_change_percent, 2), "%", sep=""))


# Assuming your dataframe is named df
# Convert QUIZ_SCORE to a factor if it isn't already
df$QUIZ_SCORE <- as.factor(df$QUIZ_SCORE)

# Calculate means for PERCENTAGE_BEFORE and PERCENTAGE_AFTER by QUIZ_SCORE
mean_percentage_before_by_quiz <- tapply(df$PERCENTAGE_BEFORE, df$QUIZ_SCORE, mean, na.rm = TRUE)
mean_percentage_after_by_quiz <- tapply(df$PERCENTAGE_AFTER, df$QUIZ_SCORE, mean, na.rm = TRUE)

# Function to calculate 99% confidence interval
calc_99ci <- function(x) {
  ci <- t.test(x, conf.level = 0.99)$conf.int
  return(ci)
}

# Calculate 99% CIs for PERCENTAGE_BEFORE and PERCENTAGE_AFTER by QUIZ_SCORE
ci_percentage_before_by_quiz <- tapply(df$PERCENTAGE_BEFORE, df$QUIZ_SCORE, calc_99ci)
ci_percentage_after_by_quiz <- tapply(df$PERCENTAGE_AFTER, df$QUIZ_SCORE, calc_99ci)

# Combine results into a data frame
meANS_Pi_df <- data.frame(
  QUIZ_SCORE = levels(df$QUIZ_SCORE),
  Mean_Before = mean_percentage_before_by_quiz,
  Lower_CI_Before = sapply(ci_percentage_before_by_quiz, function(x) x[1]),
  Upper_CI_Before = sapply(ci_percentage_before_by_quiz, function(x) x[2]),
  Mean_After = mean_percentage_after_by_quiz,
  Lower_CI_After = sapply(ci_percentage_after_by_quiz, function(x) x[1]),
  Upper_CI_After = sapply(ci_percentage_after_by_quiz, function(x) x[2])
)

# Print the results
print(meANS_Pi_df)

#Plot

# Load the necessary libraries
library(scales) # For percentage format on Y-axis

# Reshape the data from wide to long format for ggplot
long_df <- meANS_Pi_df %>%
  gather(key = "Timepoint", value = "Mean", Mean_Before, Mean_After) %>%
  mutate(QUIZ_SCORE = as.factor(QUIZ_SCORE))

# Plotting
ggplot(long_df, aes(x = QUIZ_SCORE, y = Mean, group = Timepoint, color = Timepoint)) +
  geom_line() +
  geom_point() +
  geom_errorbar(data = meANS_Pi_df, aes(x = QUIZ_SCORE, ymin = Lower_CI_Before, ymax = Upper_CI_Before), color = "gray", width = 0.1, inherit.aes = FALSE) +
  geom_errorbar(data = meANS_Pi_df, aes(x = QUIZ_SCORE, ymin = Lower_CI_After, ymax = Upper_CI_After), color = "gray", width = 0.1, inherit.aes = FALSE) +
  scale_y_continuous(labels = percent_format(scale = 100)) +  # Adjusted scale factor
  labs(title = "Mean Scores Before and After by Survey Score",
       x = "Survey Score",
       y = "Palm Oil Perception (%)",
       color = "Timepoint") +
  theme_minimal() +
  theme(legend.key = element_rect(fill = "white"))

## STATISTICAL MODELS

# Model 1

df$RATING_CHANGE_NUM <- ifelse(df$RATING_CHANGE >= 1, 1, 0)

#Treat Quiz Score as continous

df$QUIZ_SCORE <- as.numeric(as.character(df$QUIZ_SCORE))

df$QUIZ_SCORE <- recode(df$QUIZ_SCORE, 
                        `0.25` = 0, 
                        `0.5` = 1, 
                        `0.75` = 2, 
                        `1` = 3)


probit_regression <- function(data, dependent_var, independent_vars) {
  # Create the formula for the probit model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit the probit model
  model <- glm(formula, data = data, family = binomial(link = "probit"))
  
  # Print the model summary (includes Wald test statistics and Chi-squared statistic)
  print(summary(model))
  
  # Calculating pseudo R-squared values
  r_squared <- with(model, 1 - deviance/null.deviance)
  
  # Check the significance of the model
  model_significance <- summary(model)$coef[2, 4]
  
  # Calculating the change in deviance and its significance
  change_in_deviance <- with(model, null.deviance - deviance)
  change_in_df <- with(model, df.null - df.residual)
  chi_square_p_value <- with(model, pchisq(null.deviance - deviance, df.null - df.residual, lower.tail = FALSE))
  
  cat("\nChange in Deviance: ", change_in_deviance, "\n")
  cat("Degrees of Freedom Difference: ", change_in_df, "\n")
  cat("Chi-Square Test P-Value for Model Significance: ", chi_square_p_value, "\n")
  
  # Extracting coefficients, Odds Ratios, and p-values
  coefficients <- coef(model)
  odds_ratios <- exp(coefficients)
  p_values <- summary(model)$coef[, 4]
  
  # Create a dataframe to store results
  results_df <- data.frame(
    Coefficients = coefficients,
    Odds_Ratios = odds_ratios,
    P_Values = p_values,
    row.names = names(coefficients)
  )
  
  # Print R-squared and overall model summary
  cat("McFadden's Pseudo R-squared: ", r_squared, "\n")
  
  
  return(results_df)
}

dependent_variable <- "RATING_CHANGE_NUM"  # Replace with your dependent variable name
independent_variables <- c("QUIZ_SCORE", "RATING_BEFORE")  # Replace with your independent variable names

model_results_quiz <- probit_regression(df, dependent_variable, independent_variables)
print(model_results_quiz)

dependent_variable <- "RATING_CHANGE_NUM"  # Replace with your dependent variable name
independent_variables <- c("ANS_S", "ANS_D", "ANS_P", "RATING_BEFORE")  # Replace with your independent variable names

model_results_questions <- probit_regression(df, dependent_variable, independent_variables)
print(model_results_questions)

# Create an interaction term as a new column in the dataframe
df$QUIZ_SCORE_RATING_BEFORE_Interaction <- df$QUIZ_SCORE * df$RATING_BEFORE

# Define your dependent and independent variables
dependent_variable <- "RATING_CHANGE_NUM"
independent_variables <- c("QUIZ_SCORE", "RATING_BEFORE", "QUIZ_SCORE_RATING_BEFORE_Interaction")

# Run the probit regression model including the interaction term
model_results_quiz <- probit_regression(df, dependent_variable, independent_variables)
print(model_results_quiz)


## PAIRED SAMPLES T TEST

# Ensure that ANS_S, ANS_D, ANS_P, and QUIZ_SCORE are factors
df$ANS_S <- as.factor(df$ANS_S)
df$ANS_D <- as.factor(df$ANS_D)
df$ANS_P <- as.factor(df$ANS_P)
df$QUIZ_SCORE <- as.factor(df$QUIZ_SCORE)

# Function to calculate the 99% confidence interval for the mean
calc_99ci_mean <- function(x) {
  std_error <- sd(x, na.rm = TRUE) / sqrt(length(na.omit(x)))
  mean_x <- mean(x, na.rm = TRUE)
  error_margin <- qt(0.995, df = length(na.omit(x)) - 1) * std_error
  lower_ci <- mean_x - error_margin
  upper_ci <- mean_x + error_margin
  return(c(lower_ci, upper_ci))
}

# Modify the existing function to include the CI calculation
perform_t_test <- function(subset_data, factor_name = NA, level = NA) {
  test_result <- t.test(subset_data$PERCENTAGE_BEFORE, subset_data$PERCENTAGE_AFTER, paired = TRUE)
  
  # Calculate means and their 99% CIs
  mean_before <- mean(subset_data$PERCENTAGE_BEFORE, na.rm = TRUE)
  mean_after <- mean(subset_data$PERCENTAGE_AFTER, na.rm = TRUE)
  ci_before <- calc_99ci_mean(subset_data$PERCENTAGE_BEFORE)
  ci_after <- calc_99ci_mean(subset_data$PERCENTAGE_AFTER)
  
  return(data.frame(
    Factor = factor_name,
    Level = level,
    Mean_Before = mean_before,
    Mean_After = mean_after,
    Mean_Diff = mean_after - mean_before,
    CI_Before_Lower = ci_before[1],
    CI_Before_Upper = ci_before[2],
    CI_After_Lower = ci_after[1],
    CI_After_Upper = ci_after[2],
    t_value = test_result$statistic,
    p_value = test_result$p.value
  ))
}

# Overall paired t-test for the entire dataset
overall_result <- perform_t_test(df, "Overall", "All")

# Iterate over each factor and perform analyses
factors <- list(ANS_S = levels(df$ANS_S), ANS_D = levels(df$ANS_D), ANS_P = levels(df$ANS_P), QUIZ_SCORE = levels(df$QUIZ_SCORE))
results <- list(Overall_All = overall_result)

for (factor_name in names(factors)) {
  for (level in factors[[factor_name]]) {
    subset_data <- df %>% filter(!!as.name(factor_name) == level)
    analysis_result <- perform_t_test(subset_data, factor_name, level)
    results[[paste(factor_name, level, sep = "_")]] <- analysis_result
  }
}

# Combine all results into one dataframe
all_results_df <- do.call(rbind, results)

# Print the combined results
print(all_results_df)


## Only those that positively change
df_positivechange <- subset(df, RATING_CHANGE_NUM == 1)

## FREQUENCY ANALYSIS

library(dplyr)

# Frequency and percentage calculations for various variables
calculate_freq_percent <- function(data, variable_name) {
  df_freq <- as.data.frame(table(data))
  df_freq <- df_freq %>%
    mutate(Percentage = Freq / sum(Freq) * 100) %>%
    mutate(Variable = variable_name)
  
  return(df_freq)
}

df_freq_rating_changed2 <- calculate_freq_percent(df_positivechange$RATING_CHANGED, "RATING_CHANGED")
df_freq_rating_change2 <- calculate_freq_percent(df_positivechange$RATING_CHANGE, "RATING_CHANGE")
df_freq_ans_s2 <- calculate_freq_percent(df_positivechange$ANS_S, "ANS_S")
df_freq_ans_d2 <- calculate_freq_percent(df_positivechange$ANS_D, "ANS_D")
df_freq_ANS_P2 <- calculate_freq_percent(df_positivechange$ANS_P, "ANS_P")
df_freq_quizscore2 <- calculate_freq_percent(df_positivechange$QUIZ_SCORE, "Quiz_Score")

# Counts and percentages of negative, zero, and positive values in RATING_CHANGE
count_data2 <- data.frame(
  data = c("Negative", "Zero", "Positive"),
  Freq = c(sum(df_positivechange$RATING_CHANGE < 0), sum(df_positivechange$RATING_CHANGE == 0), sum(df_positivechange$RATING_CHANGE > 0)),
  Percentage = c(sum(df_positivechange$RATING_CHANGE < 0), sum(df_positivechange$RATING_CHANGE == 0), sum(df_positivechange$RATING_CHANGE > 0)) / nrow(df_positivechange) * 100,
  Variable = "RATING_CHANGE"
)

# Calculate frequencies and percentages for RATING_BEFORE
df_freq_rating_before2 <- calculate_freq_percent(df_positivechange$RATING_BEFORE, "RATING_BEFORE")

# Calculate frequencies and percentages for RATING_AFTER
df_freq_rating_after2 <- calculate_freq_percent(df_positivechange$RATING_AFTER, "RATING_AFTER")

# Combine all dataframes into one
all_freq_data2 <- bind_rows(df_freq_rating_changed2, df_freq_rating_change2, df_freq_quizscore2, df_freq_ans_s2, df_freq_ans_d2, df_freq_ANS_P2, count_data2, df_freq_rating_before2, df_freq_rating_after2)

# MEAN COMPARISON

# Calculate means
mean_before2 <- mean(df_positivechange$PERCENTAGE_BEFORE, na.rm = TRUE)
mean_after2 <- mean(df_positivechange$PERCENTAGE_AFTER, na.rm = TRUE)

mean_after2 - mean_before2

