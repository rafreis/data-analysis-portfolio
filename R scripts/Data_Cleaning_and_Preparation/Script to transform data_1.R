# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/mariaheze")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df_3 <- read.xlsx("chart_review_250_patients.xlsx")

# Read data from a specific sheet of an Excel file
df_1_pre <- read.xlsx("Pre_Test_Providers.xlsx")

df_1_post <- read.xlsx("Post_Test_Providers.xlsx")

# Get rid of special characters

names(df_3) <- gsub(" ", "_", trimws(names(df_3)))
names(df_3) <- gsub("\\s+", "_", trimws(names(df_3), whitespace = "[\\h\\v\\s]+"))
names(df_3) <- gsub("\\(", "_", names(df_3))
names(df_3) <- gsub("\\)", "_", names(df_3))
names(df_3) <- gsub("\\-", "_", names(df_3))
names(df_3) <- gsub("/", "_", names(df_3))
names(df_3) <- gsub("\\\\", "_", names(df_3)) 
names(df_3) <- gsub("\\?", "", names(df_3))
names(df_3) <- gsub("\\'", "", names(df_3))
names(df_3) <- gsub("\\,", "_", names(df_3))
names(df_3) <- gsub("\\$", "", names(df_3))
names(df_3) <- gsub("\\+", "_", names(df_3))

# Trim all values

# Loop over each column in the dataframe
df_3 <- data.frame(lapply(df_3, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))



# Trim all values and convert to lowercase

# Loop over each column in the dataframe
df_3 <- data.frame(lapply(df_3, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
    # Convert character strings to lowercase
    x <- tolower(x)
  }
  return(x)
}))



library(dplyr)
library(tidyr)

# Rename columns for consistency
df_1_post <- df_1_post %>% rename_with(~ gsub("_Post-Test", "", .), -Question)
df_1_pre <- df_1_pre %>% rename_with(~ gsub("_Pre-Test", "", .), -Question)

# Merge pre- and post-test datasets
df_1 <- left_join(
  df_1_pre %>% pivot_longer(-Question, names_to = "Provider", values_to = "Pre"),
  df_1_post %>% pivot_longer(-Question, names_to = "Provider", values_to = "Post"),
  by = c("Question", "Provider")
)

# Create wide format with columns like Q1_Pre, Q1_Post, etc.
df_1 <- df_1 %>%
  pivot_longer(cols = c(Pre, Post), names_to = "Test", values_to = "Score") %>%
  unite("Question_Test", Question, Test, sep = "_") %>%
  pivot_wider(names_from = Question_Test, values_from = Score)

# View the final dataset
print(df_1)



names(df_1) <- gsub(" ", "_", trimws(names(df_1)))
names(df_1) <- gsub("\\s+", "_", trimws(names(df_1), whitespace = "[\\h\\v\\s]+"))
names(df_1) <- gsub("\\(", "_", names(df_1))
names(df_1) <- gsub("\\)", "_", names(df_1))
names(df_1) <- gsub("\\-", "_", names(df_1))
names(df_1) <- gsub("/", "_", names(df_1))
names(df_1) <- gsub("\\\\", "_", names(df_1)) 
names(df_1) <- gsub("\\?", "", names(df_1))
names(df_1) <- gsub("\\'", "", names(df_1))
names(df_1) <- gsub("\\,", "_", names(df_1))
names(df_1) <- gsub("\\$", "", names(df_1))
names(df_1) <- gsub("\\+", "_", names(df_1))

# Trim all values

# Loop over each column in the dataframe
df_1 <- data.frame(lapply(df_1, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))



# Trim all values and convert to lowercase

# Loop over each column in the dataframe
df_1 <- data.frame(lapply(df_1, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
    # Convert character strings to lowercase
    x <- tolower(x)
  }
  return(x)
}))

library(openxlsx)

# Export df_1 as an Excel file
write.csv(df_1, file = "df_1.csv", row.names = FALSE)

