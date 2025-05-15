# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/rayanal")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("CleanData.xlsx")

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

colnames(df)


# List of source columns from which values will be transposed
source_columns <- c("Manipulation.Check2_1", "Cont_Eva2", "Invol_Cont2_1", "Invol_Cont2_2", "Invol_Cont2_3", 
                    "Cog_Ima2_1", "Cog_Ima2_2", "Cog_Ima2_3", "Cog_Ima2_4", "Cog_Ima2_5", "Cog_Ima2_6", 
                    "Cog_Ima2_7", "Cog_Ima2_8", "Cog_Ima2_9", "Cog_Ima2_10", "Aff_Ima2_1", "Aff_Ima2_2", 
                    "Aff_Ima2_3", "Aff_Ima2_4", "InT2_1", "InT2_2", "InT2_3", "Att_TikTok2_1", "Att_TikTok2_2", 
                    "Att_TikTok2_3", "Manipulation.Check3_1", "Cont_Eva3", "Invol_Cont3_1", "Invol_Cont3_2", 
                    "Invol_Cont3_3", "Cog_Ima3_1", "Cog_Ima3_2", "Cog_Ima3_3", "Cog_Ima3_4", "Cog_Ima3_5", 
                    "Cog_Ima3_6", "Cog_Ima3_7", "Cog_Ima3_8", "Cog_Ima3_9", "Cog_Ima3_10", "Aff_Ima3_1", 
                    "Aff_Ima3_2", "Aff_Ima3_3", "Aff_Ima3_4", "InT3_1", "InT3_2", "InT3_3", "Att_TikTok3_1", 
                    "Att_TikTok3_2", "Att_TikTok3_3", "Manipulation.Check4_1", "Cont_Eva4", "Invol_Cont4_1", 
                    "Invol_Cont4_2", "Invol_Cont4_3", "Cog_Ima4_1", "Cog_Ima4_2", "Cog_Ima4_3", "Cog_Ima4_4", 
                    "Cog_Ima4_5", "Cog_Ima4_6", "Cog_Ima4_7", "Cog_Ima4_8", "Cog_Ima4_9", "Cog_Ima4_10", 
                    "Aff_Ima4_1", "Aff_Ima4_2", "Aff_Ima4_3", "Aff_Ima4_4", "InT4_1", "InT4_2", "InT4_3", 
                    "Att_TikTok4_1", "Att_TikTok4_2", "Att_TikTok4_3")

# List of target columns
target_columns <- c("Manipulation.Check", "Cont_Eva", "Invol_Cont_1", "Invol_Cont_2", "Invol_Cont_3", 
                    "Cog_Ima_1.1", "Cog_Ima_2.1", "Cog_Ima_3.1", "Cog_Ima_4.1", "Cog_Ima_5.1", "Cog_Ima_6.1", 
                    "Cog_Ima_7.1", "Cog_Ima_8.1", "Cog_Ima_9.1", "Cog_Ima_10.1", "Aff_Ima_1.1", "Aff_Ima_2.1", 
                    "Aff_Ima_3.1", "Aff_Ima_4.1", "InT_1.1", "InT_2.1", "InT_3.1", "Att_TikTok_1", "Att_TikTok_2", 
                    "Att_TikTok_3", "Manipulation.Check", "Cont_Eva", "Invol_Cont_1", "Invol_Cont_2", 
                    "Invol_Cont_3", "Cog_Ima_1.1", "Cog_Ima_2.1", "Cog_Ima_3.1", "Cog_Ima_4.1", "Cog_Ima_5.1", 
                    "Cog_Ima_6.1", "Cog_Ima_7.1", "Cog_Ima_8.1", "Cog_Ima_9.1", "Cog_Ima_10.1", "Aff_Ima_1.1", 
                    "Aff_Ima_2.1", "Aff_Ima_3.1", "Aff_Ima_4.1", "InT_1.1", "InT_2.1", "InT_3.1", "Att_TikTok_1", 
                    "Att_TikTok_2", "Att_TikTok_3", "Manipulation.Check", "Cont_Eva", "Invol_Cont_1", 
                    "Invol_Cont_2", "Invol_Cont_3", "Cog_Ima_1.1", "Cog_Ima_2.1", "Cog_Ima_3.1", "Cog_Ima_4.1", 
                    "Cog_Ima_5.1", "Cog_Ima_6.1", "Cog_Ima_7.1", "Cog_Ima_8.1", "Cog_Ima_9.1", "Cog_Ima_10.1", 
                    "Aff_Ima_1.1", "Aff_Ima_2.1", "Aff_Ima_3.1", "Aff_Ima_4.1", "InT_1.1", "InT_2.1", "InT_3.1", 
                    "Att_TikTok_1", "Att_TikTok_2", "Att_TikTok_3")

# Perform the transposition
for (i in seq_along(source_columns)) {
  df[[target_columns[i]]] <- ifelse(df[[target_columns[i]]] == "" | is.na(df[[target_columns[i]]]), 
                                    df[[source_columns[i]]], 
                                    df[[target_columns[i]]])
}


# Remove the source columns from the dataset
df <- df[ , !(names(df) %in% source_columns)]

library(dplyr)

# Recode agreement scale to 1-5 throughout the dataset
recode_agreement_scale <- function(x) {
  recoded_values <- ifelse(x == "Strongly disagree", 1,
                           ifelse(x == "Disagree", 2,
                                  ifelse(x == "Somewhat disagree", 3,
                                         ifelse(x == "Neither agree nor disagree", 4,
                                                ifelse(x == "Somewhat agree", 5,
                                                       ifelse(x == "Agree", 6,
                                                              ifelse(x == "Strongly agree", 7, x)))))))
  return(recoded_values)
}

# Step 2: List of columns to apply the recoding
columns_to_recode <- c("Know_1", "Know_2", "Know_3", "Cog_Ima_1", "Cog_Ima_2", "Cog_Ima_3", "Cog_Ima_4", 
                       "Cog_Ima_5", "Cog_Ima_6", "Cog_Ima_7", "Cog_Ima_8", "Cog_Ima_9", "Cog_Ima_10", 
                       "InT_1", "InT_2", "InT_3", "Manipulation.Check", "Invol_Cont_1", "Invol_Cont_2", 
                       "Invol_Cont_3", "Cog_Ima_1.1", "Cog_Ima_2.1", "Cog_Ima_3.1", "Cog_Ima_4.1", 
                       "Cog_Ima_5.1", "Cog_Ima_6.1", "Cog_Ima_7.1", "Cog_Ima_8.1", "Cog_Ima_9.1", 
                       "Cog_Ima_10.1", "InT_1.1", "InT_2.1", "InT_3.1", "Att_TikTok_1", "Att_TikTok_2", 
                       "Att_TikTok_3")

# Step 3: Apply recoding to the specified columns
df[columns_to_recode] <- lapply(df[columns_to_recode], function(col) {
  if (is.character(col)) {
    return(recode_agreement_scale(col))
  } else {
    return(col)  # Keep original values for non-character columns
  }
})

# Export the dataframe to a CSV file
write.csv(df, "Data.csv", row.names = FALSE)
