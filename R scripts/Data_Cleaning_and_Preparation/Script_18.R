# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dalia/D50")

# Load the openxlsx library
library(openxlsx)
library(dplyr)
library(tidyr)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Verified data D75 and D78_clean.xlsx")

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


# Load necessary libraries
library(dplyr)    # For data manipulation
library(irr)      # For inter-rater reliability (Fleiss' Kappa)

# Define the mapping for the ordinal scores
score_mapping <- c("no" = 0, "marginal" = 1, "mild" = 2, "moderate" = 3, "major" = 4)

# Apply the transformation
df <- df %>%
  mutate(
    Treated.Side.Score.Numeric = as.numeric(score_mapping[Treated.Side.Score]),
    Control.Side.Score.Numeric = as.numeric(score_mapping[Control.Side.Score])
  )

# Perform Wilcoxon signed-rank test
wilcoxon_test <- wilcox.test(
  x = df$Treated.Side.Score.Numeric,
  y = df$Control.Side.Score.Numeric,
  paired = TRUE,
  exact = FALSE
)

# Save Wilcoxon results as a dataframe
wilcoxon_results <- data.frame(
  Test = "Wilcoxon Signed-Rank Test",
  Statistic = wilcoxon_test$statistic,
  P_Value = wilcoxon_test$p.value,
  Method = wilcoxon_test$method,
  Alternative_Hypothesis = wilcoxon_test$alternative
)

# Step 3: Inter-Rater Agreement (Fleiss' Kappa)
kappa_data_treated <- df %>%
  select(Dog.ID, Observer.ID, Treated.Side.Score.Numeric) %>%
  pivot_wider(
    names_from = Observer.ID,
    values_from = Treated.Side.Score.Numeric
  ) %>%
  select(-Dog.ID) %>%
  as.matrix()

kappa_data_control <- df %>%
  select(Dog.ID, Observer.ID, Control.Side.Score.Numeric) %>%
  pivot_wider(
    names_from = Observer.ID,
    values_from = Control.Side.Score.Numeric
  ) %>%
  select(-Dog.ID) %>%
  as.matrix()

kappa_treated <- kappam.fleiss(kappa_data_treated)
kappa_control <- kappam.fleiss(kappa_data_control)

# Save Fleiss' Kappa results as a dataframe
kappa_results <- data.frame(
  Side = c("Treated", "Control"),
  Kappa = c(kappa_treated$value, kappa_control$value),
  Z = c(kappa_treated$statistic, kappa_control$statistic),
  P_Value = c(kappa_treated$p.value, kappa_control$p.value)
)

# Step 4: Descriptive Statistics
treated_summary <- df %>%
  group_by(Treated.Side.Score) %>%
  summarise(Frequency = n()) %>%
  mutate(Side = "Treated")

control_summary <- df %>%
  group_by(Control.Side.Score) %>%
  summarise(Frequency = n()) %>%
  mutate(Side = "Control")

descriptive_results <- bind_rows(
  treated_summary %>%
    rename(Score = Treated.Side.Score),
  control_summary %>%
    rename(Score = Control.Side.Score)
)

# Add supplemental descriptive statistics for Wilcoxon
descriptive_stats <- data.frame(
  Side = c("Treated", "Control"),
  Mean = c(mean(df$Treated.Side.Score.Numeric, na.rm = TRUE), mean(df$Control.Side.Score.Numeric, na.rm = TRUE)),
  Median = c(median(df$Treated.Side.Score.Numeric, na.rm = TRUE), median(df$Control.Side.Score.Numeric, na.rm = TRUE)),
  SD = c(sd(df$Treated.Side.Score.Numeric, na.rm = TRUE), sd(df$Control.Side.Score.Numeric, na.rm = TRUE))
)

# Combine all results into a single Excel file
write.xlsx(
  list(
    Wilcoxon_Test = wilcoxon_results,
    Kappa_Results = kappa_results,
    Descriptive_Results = descriptive_results,
    Supplemental_Stats = descriptive_stats
  ),
  file = "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dalia/analysis_results.xlsx"
)

print("Analysis results saved to Excel.")




library(ggplot2)

# Ensure 'Score' is a factor with the correct order
descriptive_results <- descriptive_results %>%
  mutate(
    Score = factor(Score, levels = c("no", "marginal", "mild", "moderate", "major"))
  )

# Calculate percentages for stacked bar chart
descriptive_results <- descriptive_results %>%
  group_by(Side) %>%
  mutate(Percentage = Frequency / sum(Frequency) * 100)

# Stacked bar plot with percentage labels
stacked_bar_plot <- ggplot(descriptive_results, aes(x = Side, y = Frequency, fill = Score)) +
  geom_bar(stat = "identity", position = "stack") +
  geom_text(
    aes(label = paste0(round(Percentage, 1), "%")),
    position = position_stack(vjust = 0.5),
    size = 3
  ) +
  labs(
    title = "Score Distributions for Treated and Control Sides",
    x = "Side",
    y = "Frequency",
    fill = "Score"
  ) +
  theme_minimal()

# Save the stacked bar plot
ggsave("stacked_bar_plot.png", plot = stacked_bar_plot, width = 4, height = 6)


# Combine treated and control scores into a long format
score_data <- df %>%
  select(Dog.ID, Treated.Side.Score.Numeric, Control.Side.Score.Numeric) %>%
  pivot_longer(cols = c(Treated.Side.Score.Numeric, Control.Side.Score.Numeric),
               names_to = "Side",
               values_to = "Score") %>%
  mutate(Side = factor(Side, levels = c("Treated.Side.Score.Numeric", "Control.Side.Score.Numeric"),
                       labels = c("Treated", "Control")))

# Boxplot
boxplot <- ggplot(score_data, aes(x = Side, y = Score, fill = Side)) +
  geom_boxplot() +
  labs(
    title = "Boxplot of Scores for Treated and Control Sides",
    x = "Side",
    y = "Numeric Score"
  ) +
  theme_minimal()

# Save the boxplot
ggsave("boxplot_scores.png", plot = boxplot, width = 8, height = 6)

library(reshape2)

# Create agreement matrices for heatmap
agreement_matrix_treated <- kappa_data_treated %>%
  as.data.frame() %>%
  cor(use = "pairwise.complete.obs")

agreement_matrix_control <- kappa_data_control %>%
  as.data.frame() %>%
  cor(use = "pairwise.complete.obs")

# Melt data for heatmap plotting
melted_treated <- melt(agreement_matrix_treated)
melted_control <- melt(agreement_matrix_control)

# Heatmap for Treated Side
heatmap_treated <- ggplot(melted_treated, aes(Var1, Var2, fill = value)) +
  geom_tile() +
  geom_text(aes(label = round(value, 2)), color = "black", size = 3) +
  scale_fill_gradient2(
    low = "red", high = "blue", mid = "white",
    midpoint = 0.5, limit = c(0, 1)
  ) +
  labs(
    title = "Inter-Rater Agreement Heatmap (Treated Side)",
    x = "Observer",
    y = "Observer",
    fill = "Agreement"
  ) +
  theme_minimal()

# Save Treated Heatmap
ggsave("heatmap_treated.png", plot = heatmap_treated, width = 8, height = 6)

# Heatmap for Control Side
heatmap_control <- ggplot(melted_control, aes(Var1, Var2, fill = value)) +
  geom_tile() +
  geom_text(aes(label = round(value, 2)), color = "black", size = 3) +
  scale_fill_gradient2(
    low = "red", high = "blue", mid = "white",
    midpoint = 0.5, limit = c(0, 1)
  ) +
  labs(
    title = "Inter-Rater Agreement Heatmap (Control Side)",
    x = "Observer",
    y = "Observer",
    fill = "Agreement"
  ) +
  theme_minimal()

# Save Control Heatmap
ggsave("heatmap_control.png", plot = heatmap_control, width = 8, height = 6)




# Exclude Patient 7 from the dataset
df_filtered <- df %>%
  filter(Dog.ID != "patient7")

# Perform Wilcoxon signed-rank test on the filtered dataset
wilcoxon_test_filtered <- wilcox.test(
  x = df_filtered$Treated.Side.Score.Numeric,
  y = df_filtered$Control.Side.Score.Numeric,
  paired = TRUE,
  exact = FALSE
)

# Save Wilcoxon results as a dataframe
wilcoxon_results_filtered <- data.frame(
  Test = "Wilcoxon Signed-Rank Test",
  Statistic = wilcoxon_test_filtered$statistic,
  P_Value = wilcoxon_test_filtered$p.value,
  Method = wilcoxon_test_filtered$method,
  Alternative_Hypothesis = wilcoxon_test_filtered$alternative
)

# Print the results
print("Wilcoxon Signed-Rank Test Result:")
print(wilcoxon_test_filtered)

# Save the results to Excel
write.xlsx(
  list(Wilcoxon_Test_Exc = wilcoxon_results_filtered),
  file = "Wilcoxon_Patient7.xlsx"
)

print("Filtered Wilcoxon test results saved to Excel.")

# Save the results to Excel
write.xlsx(
  list(Dataframe = df),
  file = "Final Data.xlsx"
)

print("Filtered Wilcoxon test results saved to Excel.")
