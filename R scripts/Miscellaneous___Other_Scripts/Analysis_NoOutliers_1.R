# Filter out students 4 and 10
df_filtered <- df[!(df$Student %in% c(4, 10)), ]

## DESCRIPTIVE STATS BY EXAMS

vars <- c("Exam_Score", "Days_studied_of_31", "Total_reviews", "Average_for_days_studied", "Average_total", "Card_New", "Card_Learning", "Card_Relearning", "Card_Young", "Card_Mature", "Card_Suspended")
factors <- c("Exam")
calculate_means_and_sds_by_factors(df_filtered, vars, factors)
descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df_filtered, vars, factors)

## BOXPLOTS

library(ggplot2)

# Specify the factor as a column name
factor_var <- "Exam"  # replace with your actual factor column name

# Loop for generating and saving boxplots
for (var in vars) {
  # Create boxplot for each variable using ggplot2
  p <- ggplot(df_filtered, aes_string(x = factor_var, y = var)) +
    geom_boxplot() +
    stat_summary(fun.y = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = paste("Boxplot of", var, "by", factor_var),
         x = factor_var, y = var)
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = paste("boxplot_nooutliers", var, ".png", sep = ""), plot = p, width = 10, height = 6)
}

# SCATTERPLOTS

library(ggplot2)

baseline_var <- "Exam_Score"
vars_to_plot <- c("Days_studied_of_31", "Total_reviews", "Average_for_days_studied", "Average_total", "Card_New", "Card_Young")
generate_scatterplots(df_filtered, baseline_var, vars_to_plot)

## CORRELATION ANALYSIS

correlation_matrix <- calculate_correlation_matrix(df_filtered, vars, method = "pearson")

# Linear Mixed Model

library(lme4)
library(broom.mixed)
library(afex)

lmm_results <- fit_lmm_and_format("Exam_Score ~ Total_reviews + (1|Exam) + (1|Student)", df_filtered, save_plots = TRUE)

## Export Results

library(openxlsx)

# Example usage
data_list <- list(
  "Descriptive Stats" = descriptive_stats_bygroup, 
  "Correlation" = correlation_matrix, 
  "Model Results" = lmm_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_No_Outliers.xlsx")

## Correlations per student

library(openxlsx)

# Get a list of unique student IDs
unique_students <- unique(df$Student)

# Create an empty list to store correlation matrices for each student
correlation_matrices <- list()

# Create an empty list to store mean exam scores for app users and non-users
mean_scores_list <- list()

# Loop through each student
for (student_id in unique_students) {
  # Subset the data for the current student
  student_data <- df[df$Student == student_id, ]
  
  # Calculate correlations for the current student
  correlation_matrix <- calculate_correlation_matrix(student_data, vars, method = "pearson")
  
  # Add the correlation matrix to the list
  correlation_matrices[[as.character(student_id)]] <- correlation_matrix
  
  # Calculate mean exam scores for app users (Days_Studied > 0) and non-users (Days_Studied == 0)
  app_users_mean <- mean(student_data$Exam_Score[student_data$Days_studied_of_31 > 0], na.rm = TRUE)
  non_app_users_mean <- mean(student_data$Exam_Score[student_data$Days_studied_of_31 == 0], na.rm = TRUE)
  
  # Create a data frame for means
  means_df <- data.frame(Student = as.character(student_id), App_Users = app_users_mean, Non_App_Users = non_app_users_mean)
  
  # Add the means data frame to the list
  mean_scores_list[[as.character(student_id)]] <- means_df
}

# Create a new Excel workbook
wb <- createWorkbook()

# Save correlations in separate tabs for each student
for (student_id in unique_students) {
  # Get the correlation matrix for the current student
  correlation_matrix <- correlation_matrices[[as.character(student_id)]]
  
  # Create a new worksheet for the current student
  addWorksheet(wb, sheetName = paste("Correlations_Student_", student_id))
  
  # Write the correlation matrix to the worksheet
  writeData(wb, sheet = paste("Correlations_Student_", student_id), x = correlation_matrix)
}

# Create a worksheet for mean exam scores
addWorksheet(wb, sheetName = "Mean_Exam_Scores")

# Write mean exam scores for app users and non-users to the worksheet
writeData(wb, sheet = "Mean_Exam_Scores", x = do.call(rbind, mean_scores_list))

# Save the Excel workbook with correlations and mean scores
saveWorkbook(wb, "Correlations_and_Mean_Scores.xlsx", overwrite = TRUE)

## Line plot of scores per user/non user

library(ggplot2)

# Calculate mean exam scores by exam for users and non-users
means_by_exam <- aggregate(df_filtered$Exam_Score, by = list(User = ifelse(df_filtered$Days_studied_of_31 > 0, "User", "Non-User"), Exam = df_filtered$Exam), FUN = mean, na.rm = TRUE)

# Create a line plot
plot <- ggplot(means_by_exam, aes(x = Exam, y = x, color = User)) +
  geom_line() +
  labs(title = "Mean Exam Scores by Exam for Users and Non-Users",
       x = "Exam",
       y = "Mean Exam Score") +
  scale_color_manual(values = c("User" = "blue", "Non-User" = "red")) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))  # Rotate x-axis labels for better readability

# Save the plot as an image file (e.g., PNG)
ggsave("Mean_Exam_Scores_By_Exam_Line_Plot.png", plot, width = 10, height = 6)

# Print the plot
print(plot)



## Line plot of scores per user/non user (no outliers)

library(ggplot2)

# Calculate mean exam scores by exam for users and non-users
means_by_exam <- aggregate(df_filtered$Exam_Score, by = list(User = ifelse(df_filtered$Days_studied_of_31 > 0, "User", "Non-User"), Exam = df_filtered$Exam), FUN = mean, na.rm = TRUE)

# Create a bar plot
plot <- ggplot(means_by_exam, aes(x = Exam, y = x, fill = User)) +
  geom_bar(stat = "identity", position = position_dodge()) +
  labs(title = "Mean Exam Scores by Exam for Users and Non-Users",
       x = "Exam",
       y = "Mean Exam Score") +
  scale_fill_manual(values = c("User" = "blue", "Non-User" = "red")) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))  # Rotate x-axis labels for better readability

# Save the plot as an image file (e.g., PNG)
ggsave("Mean_Exam_Scores_By_Exam_NoOutliers.png", plot, width = 10, height = 6)

# Print the plot
print(plot)
