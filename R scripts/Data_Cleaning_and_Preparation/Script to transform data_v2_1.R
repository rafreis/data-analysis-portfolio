# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/mariaheze")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df_2 <- read.xlsx("_Updated_Provider_PHQ9_Utilization.xlsx")

# Read data from a specific sheet of an Excel file
df_1 <- read.xlsx("Updated_Pretest_Posttest_Scores.xlsx")


# Get rid of special characters

names(df_2) <- gsub(" ", "_", trimws(names(df_2)))
names(df_2) <- gsub("\\s+", "_", trimws(names(df_2), whitespace = "[\\h\\v\\s]+"))
names(df_2) <- gsub("\\(", "_", names(df_2))
names(df_2) <- gsub("\\)", "_", names(df_2))
names(df_2) <- gsub("\\-", "_", names(df_2))
names(df_2) <- gsub("/", "_", names(df_2))
names(df_2) <- gsub("\\\\", "_", names(df_2)) 
names(df_2) <- gsub("\\?", "", names(df_2))
names(df_2) <- gsub("\\'", "", names(df_2))
names(df_2) <- gsub("\\,", "_", names(df_2))
names(df_2) <- gsub("\\$", "", names(df_2))
names(df_2) <- gsub("\\+", "_", names(df_2))

# Trim all values

# Loop over each column in the dataframe
df_2 <- data.frame(lapply(df_2, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))



# Trim all values and convert to lowercase

# Loop over each column in the dataframe
df_2 <- data.frame(lapply(df_2, function(x) {
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

df_1 <- df_1 %>%
  pivot_longer(
    cols = -Question, # All columns except 'Question'
    names_to = c("Provider", "Test"), # Split column names into Provider and Test
    names_pattern = "(Provider_\\d+)_(Pre|Post)-Test", # Extract Provider and Test (Pre/Post)
    values_to = "Score"
  ) %>%
  unite("Question_Test", Question, Test, sep = "_") %>% # Combine Question and Test into a single column
  pivot_wider(
    names_from = "Question_Test", # Spread Question_Test into separate columns
    values_from = "Score" # Fill values in these columns
  ) %>%
  arrange(Provider)


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

