setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/rarchie")

library(readxl)
df <- read_xlsx('Survey_Key_and_Data.xlsx', sheet = 'Survey Data Select Items')

## DATA PROcESSING

library(dplyr)

# Recoding the variables
df <- df %>%
  mutate(
    YearsT_14 = case_when(
      YearsT_14 == 1 ~ "None",
      YearsT_14 == 2 ~ "Less than 5 yrs",
      YearsT_14 == 3 ~ "6-10 yrs",
      YearsT_14 == 4 ~ "11-15 yrs",
      YearsT_14 == 5 ~ "16+ years",
      TRUE ~ as.character(YearsT_14)
    ),
    YearsK3_15 = case_when(
      YearsK3_15 == 1 ~ "Less than 5 yrs",
      YearsK3_15 == 2 ~ "6-10 yrs",
      YearsK3_15 == 3 ~ "11-15 yrs",
      YearsK3_15 == 4 ~ "16-20 yrs",
      YearsK3_15 == 5 ~ "21+ yrs",
      TRUE ~ as.character(YearsK3_15)
    ),
    YearsP_16 = case_when(
      YearsP_16 == 1 ~ "Less than 5 yrs",
      YearsP_16 == 2 ~ "6-10 yrs",
      YearsP_16 == 3 ~ "11-15 yrs",
      YearsP_16 == 4 ~ "16-20 yrs",
      YearsP_16 == 5 ~ "21+ yrs",
      TRUE ~ as.character(YearsP_16)
    )
  )


# Compute Quiz Score

# Assuming your dataframe is named df
df$Quiz_Score <- with(df, (
  (ScreenerUse29 == 3) +
    (TierUse30 == 2) +
    (Initials32 == 2) +
    (MTSSUse33 == 2) +
    (MTSSWhy34 == 3)
))


library(openxlsx)

# Specify the file path and name for the Excel file
file_path <- "cleaned_data.xlsx"

# Write df to an Excel file
write.xlsx(df, file = file_path)
