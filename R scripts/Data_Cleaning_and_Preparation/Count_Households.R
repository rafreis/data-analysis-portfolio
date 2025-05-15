setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/KUSHU1940")

data <- read.csv("data_DryRecycling2.csv", sep = ";")

# Load libraries
library(dplyr)
library(stringr)

# Clean the household columns by removing any text characters and keeping only numbers
data_clean <- data %>%
  mutate(across(starts_with("Households_DryRecycling_"), ~as.numeric(gsub("[^0-9]", "", .))))

# Initialize an empty dataframe to store the results
result <- data.frame()

# Number of schemes (update this based on your data)
n_schemes <- 5

# Initialize an empty dataframe to store the results
result_sameletter <- data.frame()

for (i in 1:n_schemes) {
  # Create a column that checks if the PropertyType and PropertySubType conditions are met for the current scheme
  check_col <- paste0("check_conditions_scheme", i)
  
  data_clean <- data_clean %>%
    rowwise() %>%
    mutate(!!rlang::sym(check_col) := any(
      mapply(
        function(pt_col, pst_col) {
          any(c_across(all_of(pt_col)) == "Standard") & any(c_across(all_of(pst_col)) == "Restricted")
        },
        paste0("PropertyTypes_DryRecyclingScheme", i, "_", LETTERS[1:9], "_PropertyTypes"),
        paste0("PropertyTypes_DryRecyclingScheme", i, "_", LETTERS[1:9], "_PropertySubTypes")
      )
    ))
  
  # Calculate the sum for the current scheme, considering the new check column
  temp_result <- data_clean %>%
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
  
  # Merge this with the overall result
  if (nrow(result_sameletter) == 0) {
    result_sameletter <- temp_result
  } else {
    result_sameletter <- full_join(result_sameletter, temp_result, by = "LocalAuth")
  }
}

# Replace NA with 0
result_sameletter[is.na(result_sameletter)] <- 0

# Calculate the total sum across all schemes
result_sameletter <- result_sameletter %>%
  mutate(Total_Households_All_Schemes = rowSums(select(., starts_with("Total_Households_Scheme")), na.rm = TRUE))


# Import the package
library(openxlsx)

# Write the dataframe to Excel
write.xlsx(result_sameletter, "Result.xlsx")





















