
library(tidyverse)

# Creating the data frame
data <- tibble(
  Item = c("Latte", "Latte", "Iced Latte", "Americano", "Americano", "Iced Americano", "Cappuccino", "Cappuccino", "Caramel Macchiato", "Caramel Macchiato", "Iced Caramel Macchiato", "Espresso Macchiato", "Espresso Macchiato", "Espresso", "Espresso", "Drip Coffee", "Drip Coffee", "Tea", "Tea", "Iced Tea", "Hot Cocoa", "Hot Cocoa", "Iced Cocoa"),
  Variation = c("Medium", "Large", "", "Medium", "Large", "", "Medium", "Large", "Medium", "Large", "", "1 shot", "2 shots", "1 shot", "2 shots", "Medium", "Large", "Medium", "Large", "", "Medium", "Large", ""),
  Coffee = c(19, 22, 21.5, 19, 22, 19, 19, 22, 19, 22, 21.5, 19, 16, 19, 16, 24.132, 30.165, NA, NA, NA, NA, NA, NA),
  Whole_Milk = c(14, 16, 16, 14, 16, 16, 12, 14, 14, 16, 16, 2, 2, 2, 2, 12, 14, 14, 16, 16, 16, 18, 16)
)

# Converting from wide to long format
long_data <- data %>% 
  pivot_longer(
    cols = c(Coffee, Whole_Milk), 
    names_to = "Ingredient", 
    values_to = "Quantity",
    values_drop_na = TRUE
  )

# View the long format data
print(long_data)


library(readxl)
library(tidyverse)

# Replace the path and sheet name according to your file structure
data2 <- read_excel("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/freshbaguette/Ingredients.xlsx", sheet = "Sheet1")


# Converting from wide to long format
long_data2 <- data2 %>%
  pivot_longer(
    cols = -...1,  # Assuming 'Item' is the name of the column with the product names
    names_to = "Ingredient", 
    values_to = "Quantity",
    values_drop_na = TRUE  # Drop NA values if ingredients are not used in some products
  )


# Load the package
library(writexl)


# Define the list of data frames with sheet names as the list names
data_list <- list(
  LongData1 = long_data,
  LongData2 = long_data2
)

# Specify the path to save the Excel file
file_path <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/freshbaguette/Forecast_Ingredients.xlsx"

# Write the data frames to the Excel file
write_xlsx(data_list, file_path)

