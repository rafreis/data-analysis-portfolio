setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/KUSHU1940")

data <- read.csv("cleaned_dataA_8.csv", sep = ";")

# Load libraries
library(dplyr)
library(stringr)

# Clean the household columns by removing any text characters and keeping only numbers
data_clean <- data %>%
  mutate(across(starts_with("Households_DryRecycling_"), ~as.numeric(gsub("[^0-9]", "", .))))

# Initialize an empty dataframe to store the result
result_anyletter <- data.frame()

# Number of schemes (update this based on your data)
n_schemes <- 5

# Initialize an empty dataframe to store the result
result_anyletter <- data.frame()

for (i in 1:n_schemes) {
  # Create a column that checks if the PropertyType and PropertySubType conditions are met for the current scheme
  check_col <- paste0("check_conditions_scheme", i)
  
  property_types_cols <- grep(paste0("PropertyTypes_DryRecyclingScheme", i, "_[A-I]_PropertyTypes"), names(data_clean), value = TRUE)
  property_subtypes_cols <- grep(paste0("PropertyTypes_DryRecyclingScheme", i, "_[A-I]_PropertySubTypes"), names(data_clean), value = TRUE)
  
  data_clean <- data_clean %>%
    rowwise() %>%
    mutate(!!rlang::sym(check_col) := any(c_across(all_of(property_types_cols)) == "Standard") & any(c_across(all_of(property_subtypes_cols)) == "Standard"))
  
  
  # Calculate the sum for the current scheme, considering the new check column
  temp_result_anyletter <- data_clean %>%
    group_by(LocalAuth) %>%
    summarise(
      !!paste0("Total_Households_Scheme", i) := sum(
        if_else(
          get(paste0("check_conditions_scheme", i)) &
            get(paste0("Frequency_DryRecycling_Scheme", i)) == "Weekly" &
            get(paste0("DryRecyclingScheme", i)) == "Multi-stream",
          get(paste0("Households_DryRecycling_Scheme", i)), 0
        ),
        na.rm = TRUE
      )
    )
  
  # Merge this with the overall result_anyletter
  if (nrow(result_anyletter) == 0) {
    result_anyletter <- temp_result_anyletter
  } else {
    result_anyletter <- full_join(result_anyletter, temp_result_anyletter, by = "LocalAuth")
  }
}

# Replace NA with 0
result_anyletter[is.na(result_anyletter)] <- 0

# Calculate the total sum across all schemes
result_anyletter <- result_anyletter %>%
  mutate(Total_Households_All_Schemes = rowSums(select(., starts_with("Total_Households_Scheme")), na.rm = TRUE))























