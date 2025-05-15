setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/HungryHowies")

library(openxlsx)
df <- read.xlsx("HH317 Consolidated.xlsx")



# Create a function to rename duplicates with a suffix
rename_duplicates <- function(names) {
  unique_names <- character(length(names))
  dup_counts <- rep(0, length(names))
  
  for (i in seq_along(names)) {
    name <- names[i]
    dup_counts[i] <- sum(names[1:i] == name)
    if (dup_counts[i] > 1) {
      unique_names[i] <- paste0(name, "_", dup_counts[i])
    } else {
      unique_names[i] <- name
    }
  }
  
  unique_names
}

# Apply this function to the column names of your dataframe
names(df) <- rename_duplicates(names(df))



library(dplyr)
library(lubridate)

# Assuming df is your dataframe
# Convert Date column to Date type if it's not already
df <- df %>%
  mutate(Date = as.Date(Date, origin = "1899-12-30"), # Convert Excel serial date to Date
         MonthYear = format(Date, "%B '%y")) # Create MonthYear column

# Create MonthYear column with format "January '24"
df <- df %>% 
  mutate(MonthYear = format(Date, "%B '%y"))

# Add the restaurant column with a fixed value of 317 at the beginning
df <- df %>% 
  mutate(restaurant = 317) %>% 
  select(restaurant, MonthYear, everything())

day_map <- setNames(c("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"),
                    c("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"))

# Assuming df$Day contains abbreviated day names like "Mon", "Tue", etc.
df$Day <- day_map[df$Day]


# Assuming df$Date is already a Date object
df$Date <- as.POSIXct(df$Date, format="%Y-%m-%d", tz="UTC")
# Convert Date to character with "00:00:00" added
df$Date <- paste(as.character(df$Date), "00:00:00")


library(openxlsx)

# Define the path and name of the Excel file you want to create
output_file_path <- "317CleanData.xlsx"

# Write the dataframe to an Excel file
write.xlsx(df, file = output_file_path)




