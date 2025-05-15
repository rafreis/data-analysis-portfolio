setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/maddie_fox")

library(haven)
df <- read_sav("Volunteer Dive Instructor Survey_March 13_ 2024_10.22.sav")

# Install necessary packages if they are not already installed
if (!require("ggplot2")) install.packages("ggplot2")
if (!require("reshape2")) install.packages("reshape2")
library(ggplot2)
library(reshape2)

# Assuming your data frame 'df' is already loaded and structured correctly
# Calculate the correlation matrix


# Since we want a specific set of correlations between matched suffix variables,
# we'll subset the correlation matrix accordingly
variables <- c("Q17_satisfac_matrix_1", "Q17_satisfac_matrix_2", "Q17_satisfac_matrix_3", "Q17_satisfac_matrix_4",
               "Q17_satisfac_matrix_5", "Q17_satisfac_matrix_6", "Q17_satisfac_matrix_7", "Q17_satisfac_matrix_8",
               "Q17_satisfac_matrix_9", "Q17_satisfac_matrix_10", "Q18_satisfied_1", "Q18_satisfied_2",
               "Q18_satisfied_3", "Q18_satisfied_4", "Q18_satisfied_5", "Q18_satisfied_6", "Q18_satisfied_7",
               "Q18_satisfied_8", "Q18_satisfied_9", "Q18_satisfied_10")

cor_matrix <- cor(df[variables], use = "complete.obs")

# Melt the correlation matrix for ggplot2
cor_melted <- melt(cor_matrix)

# Create the heatmap
ggplot(cor_melted, aes(Var1, Var2, fill = value)) +
  geom_tile() +
  scale_fill_gradient2(low = "blue", high = "red", mid = "white", midpoint = 0, limit = c(-1,1), space = "Lab", name="Correlation") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        axis.title.x = element_blank(),
        axis.title.y = element_blank()) +
  labs(title = "Heatmap of Correlation Between Satisfaction and Importance Ratings")



# List of attribute descriptions for better labeling in the plot
attribute_names <- c("Autonomy - Freedom to act independently", "Relationships with colleagues/staff", "Relationships with students", 
                     "Pay - Financial or non-financial perks", "Resources - People, facilities, and/or materials", "Status - Respect received from others", "Intrinsic satisfaction - Work being rewarding", 
                     "Free time away from work", "Administrative support - Support from supervisors", "Community involvement")

# Calculate mean of importance and satisfaction for each attribute
importance_means <- sapply(df[, grep("Q17_satisfac_matrix", names(df))], mean, na.rm = TRUE)
satisfaction_means <- sapply(df[, grep("Q18_satisfied", names(df))], mean, na.rm = TRUE)

# Create a dataframe for plotting
comparison_df <- data.frame(Importance = importance_means, Satisfaction = satisfaction_means)
comparison_df$Attribute <- attribute_names

# Assuming 'df' is your dataframe and 'comparison_df' has been created as before

# Precompute mean and standard deviation for importance and satisfaction
importance_mean <- mean(comparison_df$Importance, na.rm = TRUE)
importance_sd <- sd(comparison_df$Importance, na.rm = TRUE)
satisfaction_mean <- mean(comparison_df$Satisfaction, na.rm = TRUE)

# Plotting with ggplot2
library(ggplot2)
ggplot(comparison_df, aes(x = Importance, y = Satisfaction, label = Attribute)) +
  geom_point(aes(color = Satisfaction - Importance), size = 4) +  # Color points based on the gap
  geom_text(aes(hjust = ifelse(Importance < importance_mean - importance_sd, 0, 1.1), 
                vjust = ifelse(Satisfaction < satisfaction_mean, -0.5, 0.5)), 
            label = comparison_df$Attribute, size = 3, check_overlap = TRUE) +  
  scale_color_gradient2(low = "red", mid = "yellow", high = "green", midpoint = 0) +
  labs(title = "Comparison of Importance vs Satisfaction for Job Components",
       x = "Importance Mean Score",
       y = "Satisfaction Mean Score") +
  theme_minimal() +
  geom_smooth(method = "lm", se = FALSE, color = "blue") +
  xlim(3.5,5) +
  ylim (2.5,5) + # Add a linear regression line without SE
  theme(plot.margin = unit(c(1, 1, 1, 2), "lines"))


# BAR CHARTS
# Calculate means and standard errors for importance and satisfaction
importance_means <- sapply(df[, grep("Q17_satisfac_matrix", names(df))], mean, na.rm = TRUE)
satisfaction_means <- sapply(df[, grep("Q18_satisfied", names(df))], mean, na.rm = TRUE)
importance_ses <- sapply(df[, grep("Q17_satisfac_matrix", names(df))], function(x) sd(x, na.rm = TRUE) / sqrt(length(x)))
satisfaction_ses <- sapply(df[, grep("Q18_satisfied", names(df))], function(x) sd(x, na.rm = TRUE) / sqrt(length(x)))

# Prepare a data frame for ggplot
bar_data <- data.frame(
  Attribute = rep(attribute_names, 2),
  Score = c(importance_means, satisfaction_means),
  SE = c(importance_ses, satisfaction_ses),
  Type = rep(c("Importance", "Satisfaction"), each = length(importance_means))
)

# Bar chart with error bars
ggplot(bar_data, aes(x = Attribute, y = Score, fill = Type)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), width = 0.7) +
  geom_errorbar(aes(ymin = Score - SE, ymax = Score + SE), width = 0.25, position = position_dodge(width = 0.8)) +
  scale_fill_manual(values = c("blue", "orange")) +
  labs(title = "Mean Scores with Confidence Intervals for Importance and Satisfaction",
       x = "Attribute", y = "Mean Score") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) + 
  theme(plot.margin = unit(c(1, 1, 1, 7), "lines"))


# Calculate the differences between importance and satisfaction means
differences <- satisfaction_means - importance_means 

# Prepare a data frame for ggplot
difference_data <- data.frame(
  Attribute = attribute_names,
  Difference = differences
)

# Paired difference plot
ggplot(difference_data, aes(x = Attribute, y = Difference, fill = Difference > 0)) +
  geom_bar(stat = "identity") +
  scale_fill_manual(values = c("red", "green"), name = "Difference", labels = c("Satisfaction < Importance", "Satisfaction > Importance")) +
  labs(title = "Paired Difference Between Importance and Satisfaction",
       x = "Attribute", y = "Difference (Satisfaction - Importance)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))+ 
  theme(plot.margin = unit(c(1, 1, 1, 7), "lines"))


