setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/miles2nowhere")

library(readxl)
df <- read_xlsx('Indiana County Data.xlsx')

# Get rid of spaces in column names
names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("/", "_", names(df))

# Load necessary library
library(tidyverse)

# Select relevant columns and handle missing values
df_selected <- df %>% 
  select(Median_HH_Income, Unemployment_Rate, Population_Growth, Poverty_Rate) %>%
  na.omit()

# Scale the variables
df_scaled <- scale(df_selected)

# Factor Analysis
fa_result <- factanal(df_scaled, factors = 1)
print(fa_result)

# PCA
pca_result <- prcomp(df_scaled, center = TRUE, scale. = TRUE)
summary(pca_result)

# Inverting the factor loadings
inverted_loadings <- -fa_result$loadings[,1]

# Ensure df_scaled is a numeric matrix
df_scaled_matrix <- as.matrix(df_scaled)

# Calculate factor scores using inverted loadings
# No need to center and scale again as df_scaled is already processed
factor_scores <- df_scaled_matrix %*% matrix(inverted_loadings, ncol = 1)

# Adding the factor scores to the original dataframe
df$economic_success <- factor_scores

# Find the minimum and maximum of the 'economic_success' column
min_val <- min(df$economic_success)
max_val <- max(df$economic_success)

# Rescale the 'economic_success' column to range from 0 to 100
df$economic_success_scaled <- (df$economic_success - min_val) / (max_val - min_val) * 100


## CORRELATION

# Recoding Yes/No variables to 1/0
df$Natural_Amenties <- ifelse(df$Natural_Amenties == "Yes", 1, 0)
df$University_25_Mins <- ifelse(df$University_25_Mins == "Yes", 1, 0)
df$Interstate <- ifelse(df$Interstate == "Yes", 1, 0)
df$Mins_Urban_Center <- ifelse(df$`25_Mins_Urban_Center` == "Yes", 1, 0)

# Calculate correlations
calculate_correlation_matrix <- function(data, variables, method = "pearson") {
  # Ensure the method is valid
  if (!method %in% c("pearson", "spearman")) {
    stop("Method must be either 'pearson' or 'spearman'")
  }
  
  # Subset the data for the specified variables
  data_subset <- data[variables]
  
  # Calculate the correlation matrix
  corr_matrix <- cor(data_subset, method = method, use = "complete.obs")
  
  # Initialize a matrix to hold formatted correlation values
  corr_matrix_formatted <- matrix(nrow = nrow(corr_matrix), ncol = ncol(corr_matrix))
  rownames(corr_matrix_formatted) <- colnames(corr_matrix)
  colnames(corr_matrix_formatted) <- colnames(corr_matrix)
  
  # Calculate p-values and apply formatting with stars
  for (i in 1:nrow(corr_matrix)) {
    for (j in 1:ncol(corr_matrix)) {
      if (i == j) {  # Diagonal elements (correlation with itself)
        corr_matrix_formatted[i, j] <- format(round(corr_matrix[i, j], 3), nsmall = 3)
      } else {
        p_value <- cor.test(data_subset[[i]], data_subset[[j]], method = method)$p.value
        stars <- ifelse(p_value < 0.001, "***", 
                        ifelse(p_value < 0.01, "**", 
                               ifelse(p_value < 0.05, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}

vars = c("Health_Ranking", "Natural_Amenties", "Mins_Urban_Center", "University_25_Mins", "Manufacturing","Education", "Interstate", "economic_success")
# Convert each variable to numeric
df[vars] <- lapply(df[vars], as.numeric)
correlation_matrix <- calculate_correlation_matrix(df, vars, method = "pearson")

# Load necessary library
library(ggplot2)

# Scatterplot for economic_success vs Health_Ranking
plot1 <- ggplot(df, aes(x = economic_success_scaled, y = Health_Ranking)) +
  geom_point() +
  theme_minimal() +
  labs(title = "Scatterplot of Economic Success vs Health Ranking",
       x = "Economic Success",
       y = "Health Ranking")

# Save plot1
ggsave("economic_success_vs_health_ranking.png", plot1, width = 8, height = 6)

# Scatterplot for economic_success vs Education
plot2 <- ggplot(df, aes(x = economic_success_scaled, y = Education)) +
  geom_point() +
  theme_minimal() +
  labs(title = "Scatterplot of Economic Success vs Education",
       x = "Economic Success",
       y = "Education")

# Save plot2
ggsave("economic_success_vs_education.png", plot2, width = 8, height = 6)


library(openxlsx)


# Create a new workbook
wb <- createWorkbook()

# Add sheets to the workbook
addWorksheet(wb, "DataFrame")
addWorksheet(wb, "Correlation Matrix")

# Write data to the respective sheets
writeData(wb, sheet = "DataFrame", df)
writeData(wb, sheet = "Correlation Matrix", correlation_matrix)

# Save the workbook
saveWorkbook(wb, "Data_and_Correlation_Matrix.xlsx", overwrite = TRUE)