setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/HungryHowies")

library(openxlsx)
df <- read.xlsx("317CleanData.xlsx")
colnames(df)


datecolumn <- "date"
Stringcolumns <- c("SheetName", "day", "paidout.description")

library(dplyr)
library(lubridate)
library(tidyr)

df <- df %>%
  mutate(date = case_when(
    # Assuming numeric values represent Excel serial dates for restaurant 369
    Restaurant == 369 & !is.na(as.numeric(date)) ~ as.POSIXct(as.Date(as.numeric(date), origin = "1899-12-30")),
    # Directly parse the standard datetime strings for restaurant 317
    Restaurant == 317 & is.character(date) ~ as.POSIXct(date, format = "%Y-%m-%d %H:%M:%S", tz = "UTC"),
    # Fallback for unexpected formats or other restaurants
    TRUE ~ as.POSIXct(NA)
  ))

# Optional: If you need to ensure all dates are in the same timezone or format them as character
df$date <- format(df$date, "%Y-%m-%d %H:%M:%S", tz = "UTC")



df$SheetName <- as.character(df$SheetName)
df$day <- as.character(df$day)
df$paidout.description <- as.character(df$paidout.description)


numeric_columns <- setdiff(names(df), c("date", "SheetName", "day", "paidout.description"))
df[numeric_columns] <- lapply(df[numeric_columns], function(x) as.numeric(as.character(x)))


tolerance <- 1e-5

# Apply the condition to all numeric columns
df <- df %>%
  mutate(across(where(is.numeric), ~if_else(abs(.) < tolerance, 0, .)))


# Export the cleaned and standardized dataframe back to an Excel file
write.xlsx(df, "CleanedData_Export.xlsx", overwrite = TRUE)
