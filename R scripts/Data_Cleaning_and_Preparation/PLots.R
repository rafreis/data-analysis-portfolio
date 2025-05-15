library(ggplot2)
library(dplyr)
library(tidyr)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/onkelj99")

# Assuming df_ttest_results is your dataframe name and it's structured as shown previously
plot_comparison_columns <- function(results_df) {
  # Transform the data to a long format for easy use with ggplot2
  results_long <- results_df %>%
    pivot_longer(cols = c(Group1_Mean, Group2_Mean), names_to = "Group", values_to = "Mean") %>%
    mutate(Group = ifelse(Group == "Group1_Mean", paste(Group1, "Mean"), paste(Group2, "Mean"))) %>%
    mutate(Group = gsub("0", "No", Group), Group = gsub("1", "Yes", Group)) %>%  # Replace 0 with No, and 1 with Yes in the group labels
    mutate(Group_Var = factor(Group_Var, levels = unique(Group_Var)))  # Make sure Group_Var is a factor for ordering
  
  # Create the plot
  p <- ggplot(results_long, aes(x = Group_Var, y = Mean, fill = Group)) +
    geom_col(position = "dodge") +
    geom_text(aes(label = round(Mean, 2)), vjust = -0.3, position = position_dodge(width = 0.9), size = 3.5) + # Add data labels
    facet_wrap(~Variable, scales = "free_y") +
    labs(title = "Stacked Comparison of Means by Group and Variable",
         x = "Group",
         y = "Mean Value",
         fill = "Group") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 0.7))  # Adjust text angle for better readability
  
  return(p)
}

# Generate the plot using your test results dataframe
comparison_plot1 <- plot_comparison_columns(df_ttest_results)



# Print the plot if in an interactive session, otherwise save to file
if (interactive()) {
  print(comparison_plot1)
} else {
  ggsave("comparison_plot.png", comparison_plot1, width = 12, height = 8)
}

ggsave("comparison_plot.png", comparison_plot1, width = 12, height = 8)


library(ggplot2)
library(dplyr)
library(tidyr)

# Function to plot two-way disaggregated t-test results focusing only on Mean
plot_two_way_disaggregated_ttest_means <- function(results_df) {
  # Transform the data to a long format focusing only on Mean values
  results_long <- results_df %>%
    pivot_longer(cols = c(Group1_Mean, Group2_Mean),
                 names_to = "Metric_Type", values_to = "Value") %>%
    mutate(Group = case_when(
      Metric_Type == "Group1_Mean" ~ paste(Subset, Group_Var, Group1, "Mean", sep = " - "),
      Metric_Type == "Group2_Mean" ~ paste(Subset, Group_Var, Group2, "Mean", sep = " - ")
    )) %>%
    mutate(Group_Var = factor(Group_Var, levels = unique(Group_Var)))  # Order factors based on the unique appearance
  
  # Create the plot
  p <- ggplot(results_long, aes(x = Group_Var, y = Value, fill = Group)) +
    geom_col(position = "dodge") +
    geom_text(aes(label = round(Value, 2)), vjust = -0.5, position = position_dodge(width = 0.9), size = 3) +
    facet_wrap(~Variable, scales = "free_y") +
    labs(title = "Two-Way Disaggregated T-Test Results by Group and Variable",
         x = "Grouping Variable",
         y = "Value",
         fill = "Group") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 0.5), axis.title.x = element_blank())
  
  return(p)
}

# Example usage with the updated function
df_two_way_ttest_results <- df_ttest_results_twoway  # Assume this dataframe is loaded with the correct data
comparison_plot_two_way_means <- plot_two_way_disaggregated_ttest_means(df_two_way_ttest_results)   

# Print the plot if in an interactive session, otherwise save to file
if (interactive()) {
  print(comparison_plot_two_way_means)
} else {
  ggsave("two_way_comparison_plot_means_only.png", comparison_plot_two_way_means, width = 12, height = 8)
}

ggsave("two_way_comparison_plot_means_only.png", comparison_plot_two_way_means, width = 12, height = 8)




library(ggplot2)
library(dplyr)
library(tidyr)

# Function to plot three-way disaggregated t-test results
plot_three_way_disaggregated_ttest <- function(results_df) {
  # Transform the data to a long format focusing on both Mean and Median values if needed
  results_long <- results_df %>%
    pivot_longer(cols = c(Group1_Mean, Group2_Mean), # Include Median cols if required
                 names_to = "Metric_Type", values_to = "Value") %>%
    mutate(
      Group = case_when(
        grepl("Group1", Metric_Type) ~ paste(Group1, "Group"),
        grepl("Group2", Metric_Type) ~ paste(Group2, "Group")
      ),
      Group = ifelse(Group == "0 Group", "No", "Yes")  # Replacing 0 with No, 1 with Yes
    ) %>%
    mutate(
      Group = paste(Subset, Group_Var, Group, sep = " - "),
      Group_Var = factor(Group_Var, levels = unique(Group_Var))  # Order factors based on the unique appearance
    )
  
  # Create the plot
  p <- ggplot(results_long, aes(x = Group_Var, y = Value, fill = Group)) +
    geom_col(position = "dodge") +
    geom_text(aes(label = round(Value, 2)), vjust = -0.5, position = position_dodge(width = 0.9), size = 3) +
    facet_wrap(~Variable, scales = "free_y") +
    labs(title = "Three-Way Disaggregated T-Test Results by Group and Variable",
         x = "Grouping Variable",
         y = "Value",
         fill = "Group") +
    theme_minimal() +
    theme(axis.text.x = element_text(hjust = 0.5), axis.title.x = element_blank())
  
  return(p)
}

comparison_plot_three_way <- plot_three_way_disaggregated_ttest(df_ttest_results_threeway)   

# Print the plot if in an interactive session, otherwise save to file
if (interactive()) {
  print(comparison_plot_three_way)
} else {
  ggsave("three_way_comparison_plot.png", comparison_plot_three_way, width = 12, height = 8)
}

# Save the plot to file always
ggsave("three_way_comparison_plot.png", comparison_plot_three_way, width = 12, height = 8)




