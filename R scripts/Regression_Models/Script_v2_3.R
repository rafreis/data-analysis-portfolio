# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/gabyrios15")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("LA TABLA2.xlsx", sheet = "Sheet1")

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

str(df)

# Load necessary libraries
library(dplyr)
library(tidyr)

# Function to count distinct values for each Case.ID across all columns
count_distinct_values <- function(df) {
  df %>%
    group_by(Case.ID) %>%  # Group by Case.ID
    summarise(across(everything(), ~ n_distinct(na.omit(.)), .names = "{col}_distinct")) %>%
    ungroup()  # Remove grouping for cleaner output
}

# Apply function to dataframe
distinct_counts_table <- count_distinct_values(df)

# Print the result
print(distinct_counts_table)


# Load necessary libraries
library(dplyr)
library(tidyr)

# Define columns to dichotomize
columns_to_dichotomize <- c("Trauma.M", "Circum.D", "Class", "Side", "Skeletal.R", 
                            "Sub.Cat.Skeletal._Face_", "Sub.Cat.Skeletal._Vault_", 
                            "Axial", "Appendicular")

# Function to process the dataset
process_dataset <- function(df, columns_to_dichotomize) {
  message("Step 1: Checking valid columns (excluding fully NA ones)")
  
  # Define valid columns (ignore fully NA columns)
  valid_columns <- columns_to_dichotomize[sapply(df[columns_to_dichotomize], function(x) any(!is.na(x)))]
  message("Valid columns: ", paste(valid_columns, collapse = ", "))
  
  message("Step 2: Converting all selected columns to character")
  
  # Convert all selected columns to character to avoid pivoting issues
  df[valid_columns] <- lapply(df[valid_columns], as.character)
  
  message("Step 3: Reshaping data before counting occurrences")
  
  # Reshape data into long format for counting
  long_data <- df %>%
    select(Case.ID, all_of(valid_columns)) %>%
    pivot_longer(cols = all_of(valid_columns), names_to = "Variable", values_to = "Value") %>%
    filter(!is.na(Value))  # Remove NAs before counting
  
  message("Step 4: Counting occurrences per Case.ID and Value")
  
  # Count occurrences per Case.ID and Value
  count_data <- long_data %>%
    group_by(Case.ID, Variable, Value) %>%
    summarise(Count = n(), .groups = "drop") %>%
    mutate(VarName = paste0(Variable, "_", Value)) %>%
    select(Case.ID, VarName, Count) %>%
    pivot_wider(names_from = VarName, values_from = Count, values_fill = 0)
  
  message("Step 5: Extracting random values for other columns")
  
  # Retain one random value for other columns
  random_values <- df %>%
    group_by(Case.ID) %>%
    summarise(across(-all_of(valid_columns), ~ {
      non_na_values <- na.omit(.)
      if (length(non_na_values) > 0) sample(non_na_values, 1) else NA
    }, .names = "{col}"), .groups = "drop")
  
  message("Step 6: Merging processed data")
  final_df <- left_join(random_values, count_data, by = "Case.ID")
  
  message("Final dataset created successfully!")
  return(final_df)
}

# Apply function
df_processed <- process_dataset(df, columns_to_dichotomize)



# Load necessary libraries
library(dplyr)
library(tidyr)

# Define columns to dichotomize
columns_to_dichotomize <- c("Trauma.M", "Circum.D", "Class", "Side", "Skeletal.R", 
                            "Sub.Cat.Skeletal._Face_", "Sub.Cat.Skeletal._Vault_", 
                            "Axial", "Appendicular")

# Function to process the dataset
process_dataset <- function(df, columns_to_dichotomize) {
  message("Step 1: Checking valid columns (excluding fully NA ones)")
  
  # Define valid columns (ignore fully NA columns)
  valid_columns <- columns_to_dichotomize[sapply(df[columns_to_dichotomize], function(x) any(!is.na(x)))]
  message("Valid columns: ", paste(valid_columns, collapse = ", "))
  
  message("Step 2: Converting all selected columns to character")
  
  # Convert all selected columns to character to avoid pivoting issues
  df[valid_columns] <- lapply(df[valid_columns], as.character)
  
  message("Step 3: Reshaping data before encoding")
  
  # Reshape data into long format for encoding
  long_data <- df %>%
    select(Case.ID, all_of(valid_columns)) %>%
    pivot_longer(cols = all_of(valid_columns), names_to = "Variable", values_to = "Value") %>%
    filter(!is.na(Value))  # Remove NAs before encoding
  
  message("Step 4: Creating binary indicators per Case.ID and Value")
  
  # Create binary indicator (1 if present, 0 otherwise)
  binary_data <- long_data %>%
    distinct(Case.ID, Variable, Value) %>%  # Ensure uniqueness
    mutate(VarName = paste0(Variable, "_", Value), Presence = 1) %>%
    select(Case.ID, VarName, Presence) %>%
    pivot_wider(names_from = VarName, values_from = Presence, values_fill = 0)  # Fill missing values with 0
  
  message("Step 5: Extracting random values for other columns")
  
  # Retain one random value for other columns
  random_values <- df %>%
    group_by(Case.ID) %>%
    summarise(across(-all_of(valid_columns), ~ {
      non_na_values <- na.omit(.)
      if (length(non_na_values) > 0) sample(non_na_values, 1) else NA
    }, .names = "{col}"), .groups = "drop")
  
  message("Step 6: Merging processed data")
  final_df <- left_join(random_values, binary_data, by = "Case.ID")
  
  message("Final dataset created successfully!")
  return(final_df)
}

# Apply function
df_processed <- process_dataset(df, columns_to_dichotomize)

str(df_processed)


create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each category variable
  for (category in categories) {
    # Ensure the category variable is a factor
    data[[category]] <- factor(data[[category]])
    
    # Calculate counts
    counts <- table(data[[category]])
    
    # Create a dataframe for this category
    freq_table <- data.frame(
      "Category" = rep(category, length(counts)),
      "Level" = names(counts),
      "Count" = as.integer(counts),
      stringsAsFactors = FALSE
    )
    
    # Calculate and add percentages
    freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
    
    # Add the result to the list
    all_freq_tables[[category]] <- freq_table
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

vars <- c("Age","Sex", "Stature", "Weight_lb_", "Year")
df_freq <- create_frequency_tables(df_processed, vars)

# Replace Year values < 2017 with NA
df_processed <- df_processed %>%
  mutate(Year = ifelse(Year < 2017, NA, Year))

# Function to categorize Age, Stature, and Weight
categorize_numeric_vars <- function(df) {
  message("Step 1: Categorizing Age into Numeric Ranges")
  
  df <- df %>%
    mutate(
      Age_Group = case_when(
        Age >= 18 & Age <= 29 ~ "18-29",
        Age >= 30 & Age <= 39 ~ "30-39",
        Age >= 40 & Age <= 49 ~ "40-49",
        Age >= 50 & Age <= 59 ~ "50-59",
        Age >= 60 ~ "60+",
        TRUE ~ NA_character_  # Keep NA if Age is missing
      )
    )
  
  message("Step 2: Categorizing Stature into Numeric Ranges")
  
  df <- df %>%
    mutate(Stature_Category = case_when(
      Stature < 64 ~ "<64",
      Stature >= 64 & Stature <= 69 ~ "64-69",
      Stature >= 70 & Stature <= 75 ~ "70-75",
      Stature > 75 ~ ">75",
      TRUE ~ NA_character_
    ))
  
  message("Step 3: Categorizing Weight into Numeric Ranges")
  
  df <- df %>%
    mutate(Weight_Category = case_when(
      Weight_lb_ < 125 ~ "<125",
      Weight_lb_ >= 125 & Weight_lb_ < 175 ~ "125-174",
      Weight_lb_ >= 175 & Weight_lb_ < 225 ~ "175-224",
      Weight_lb_ >= 225 ~ "≥225",
      TRUE ~ NA_character_
    ))
  
  return(df)
}

# Apply function
df_processed <- categorize_numeric_vars(df_processed)

str(df_processed)



chi_square_analysis_multiple <- function(data, row_vars, col_vars) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over column variables
  for (col_var in col_vars) {
    # Ensure the column variable is a factor
    data[[col_var]] <- factor(data[[col_var]])
    
    # Iterate over row variables
    for (row_var in row_vars) {
      # Ensure the row variable is a factor
      data[[row_var]] <- factor(data[[row_var]])
      
      # Create a crosstab with percentages
      crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 2) * 100
      
      # Check if the crosstab is 2x2
      is_2x2_table <- all(dim(table(data[[row_var]], data[[col_var]])) == 2)
      
      # Perform chi-square test with correction for 2x2 tables
      chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]), correct = !is_2x2_table)
      
      # Convert crosstab to a dataframe
      crosstab_df <- as.data.frame.matrix(crosstab)
      
      # Create a dataframe for this pair of variables
      for (level in levels(data[[row_var]])) {
        level_df <- data.frame(
          "Row_Variable" = row_var,
          "Row_Level" = level,
          "Column_Variable" = col_var,
          check.names = FALSE
        )
        
        if (nrow(crosstab_df) > 0 && !is.null(crosstab_df[level, , drop = FALSE])) {
          level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
        }
        
        level_df$Chi_Square <- chi_square_test$statistic
        level_df$P_Value <- chi_square_test$p.value
        
        # Add the result to the list
        results[[paste0(row_var, "_", col_var, "_", level)]] <- level_df
      }
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}


# Define row (independent) and column (dependent) variables for crosstab analysis
row_vars <- list(
  "Trauma Mechanism" = c("Trauma.M_BFT", "Trauma.M_SFT", "Trauma.M_GSW", "Trauma.M_Thermal"),
  "Skeletal Region" = c("Skeletal.R_Cranial", "Skeletal.R_Post"),
  "Age Group" = "Age_Group",
  "Sex" = "Sex",
  "Year" = "Year",
  "Stature" = "Stature_Category",
  "Weight" = "Weight_Category"
)

col_vars <- list(
  "Skeletal Region" = c("Skeletal.R_Cranial", "Skeletal.R_Post"),
  "Circumstance of Death" = c("Circum.D_Homicide", "Circum.D_Suicide", "Circum.D_Accident"),
  "Trauma Mechanism" = c("Trauma.M_BFT", "Trauma.M_SFT", "Trauma.M_GSW", "Trauma.M_Thermal")
)

# Apply chi-square analysis function
chi_square_results <- chi_square_analysis_multiple(df_processed, unlist(row_vars), unlist(col_vars))

# View results
print(chi_square_results)

colnames(df_processed)

df_freq2 <- create_frequency_tables(df_processed, c("Age_Group"         ,                                                                    
                                                    "Stature_Category"   ,                                                                   
                                                    "Weight_Category",
                                                    "Year",
                                                    "Sex"))

## Export Results

library(openxlsx)

save_apa_formatted_excel <- function(data_list, filename) {
  wb <- createWorkbook()  # Create a new workbook
  
  for (i in seq_along(data_list)) {
    # Define the sheet name
    sheet_name <- names(data_list)[i]
    if (is.null(sheet_name)) sheet_name <- paste("Sheet", i)
    addWorksheet(wb, sheet_name)  # Add a sheet to the workbook
    
    # Convert matrices to data frames, if necessary
    data_to_write <- if (is.matrix(data_list[[i]])) as.data.frame(data_list[[i]]) else data_list[[i]]
    
    # Include row names as a separate column, if they exist
    if (!is.null(row.names(data_to_write))) {
      data_to_write <- cbind("Index" = row.names(data_to_write), data_to_write)
    }
    
    # Write the data to the sheet
    writeData(wb, sheet_name, data_to_write)
    
    # Define styles
    header_style <- createStyle(textDecoration = "bold", border = "TopBottom", borderColour = "black", borderStyle = "thin")
    bottom_border_style <- createStyle(border = "bottom", borderColour = "black", borderStyle = "thin")
    
    # Apply styles
    addStyle(wb, sheet_name, style = header_style, rows = 1, cols = 1:ncol(data_to_write), gridExpand = TRUE)
    addStyle(wb, sheet_name, style = bottom_border_style, rows = nrow(data_to_write) + 1, cols = 1:ncol(data_to_write), gridExpand = TRUE)
    
    # Set column widths to auto
    setColWidths(wb, sheet_name, cols = 1:ncol(data_to_write), widths = "auto")
  }
  
  # Save the workbook
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Example usage
data_list <- list(
  "Frequency Table" = df_freq2, 
  "Crosstabs" = chi_square_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")










# Second Order
# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/gabyrios15/order 2")

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Table 3.xlsx", sheet = "Sheet1")

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

str(df)

# Load necessary libraries
library(dplyr)
library(tidyr)

# Function to count distinct values for each Case.ID across all columns
count_distinct_values <- function(df) {
  df %>%
    group_by(Case.ID) %>%  # Group by Case.ID
    summarise(across(everything(), ~ n_distinct(na.omit(.)), .names = "{col}_distinct")) %>%
    ungroup()  # Remove grouping for cleaner output
}

# Apply function to dataframe
distinct_counts_table <- count_distinct_values(df)

# Print the result
print(distinct_counts_table)


# Load necessary libraries
library(dplyr)
library(tidyr)

colnames(df)

# Define columns to dichotomize
columns_to_dichotomize <- c("Trauma.M", "Circum.D", "Class", "Bone.Side", "Skeletal.Region", 
                            "Skeletal.Cranial.Subcategory._Face_", "Skeletal.Cranial.Subcategory._Vault_", 
                            "Skeletal.Cranial.Subcategory._Axial_", "Skeletal.Cranial.Subcategory._Appendicular_")

# Load the package
library(writexl)

# Export to Excel
write_xlsx(df_processed, "df_processed.xlsx")

# Function to process the dataset
process_dataset <- function(df, columns_to_dichotomize) {
  message("Step 1: Checking valid columns (excluding fully NA ones)")
  
  # Define valid columns (ignore fully NA columns)
  valid_columns <- columns_to_dichotomize[sapply(df[columns_to_dichotomize], function(x) any(!is.na(x)))]
  message("Valid columns: ", paste(valid_columns, collapse = ", "))
  
  message("Step 2: Converting all selected columns to character")
  
  # Convert all selected columns to character to avoid pivoting issues
  df[valid_columns] <- lapply(df[valid_columns], as.character)
  
  message("Step 3: Reshaping data before encoding")
  
  # Reshape data into long format for encoding
  long_data <- df %>%
    select(Case.ID, all_of(valid_columns)) %>%
    pivot_longer(cols = all_of(valid_columns), names_to = "Variable", values_to = "Value") %>%
    filter(!is.na(Value))  # Remove NAs before encoding
  
  message("Step 4: Creating binary indicators per Case.ID and Value")
  
  # Create binary indicator (1 if present, 0 otherwise)
  binary_data <- long_data %>%
    distinct(Case.ID, Variable, Value) %>%  # Ensure uniqueness
    mutate(VarName = paste0(Variable, "_", Value), Presence = 1) %>%
    select(Case.ID, VarName, Presence) %>%
    pivot_wider(names_from = VarName, values_from = Presence, values_fill = 0)  # Fill missing values with 0
  
  message("Step 5: Extracting random values for other columns")
  
  # Retain one random value for other columns
  random_values <- df %>%
    group_by(Case.ID) %>%
    summarise(across(-all_of(valid_columns), ~ {
      non_na_values <- na.omit(.)
      if (length(non_na_values) > 0) sample(non_na_values, 1) else NA
    }, .names = "{col}"), .groups = "drop")
  
  message("Step 6: Merging processed data")
  final_df <- left_join(random_values, binary_data, by = "Case.ID")
  
  message("Final dataset created successfully!")
  return(final_df)
}

# Apply function
df_processed <- process_dataset(df, columns_to_dichotomize)

colnames(df_processed)

colnames(df_processed)
# Replace Year values < 2017 with NA
df_processed <- df_processed %>%
  mutate(Year = ifelse(Year < 2017, NA, Year))


# Function to categorize Age, Stature, and Weight
categorize_numeric_vars <- function(df) {
  message("Step 1: Categorizing Age into Numeric Ranges")
  
  df <- df %>%
    mutate(
      Age_Group = case_when(
        Age >= 18 & Age <= 29 ~ "18-29",
        Age >= 30 & Age <= 39 ~ "30-39",
        Age >= 40 & Age <= 49 ~ "40-49",
        Age >= 50 & Age <= 59 ~ "50-59",
        Age >= 60 ~ "60+",
        TRUE ~ NA_character_  # Keep NA if Age is missing
      )
    )
  
  message("Step 2: Categorizing Stature into Numeric Ranges")
  
  df <- df %>%
    mutate(Height_Category = case_when(
      Height._meters_ < 1.63 ~ "<1.63m",
      Height._meters_ >= 1.63 & Height._meters_ <= 1.75 ~ "1.63-1.75m",
      Height._meters_ > 1.75 ~ ">1.75m",
      TRUE ~ NA_character_
    ))
  
  message("Step 3: Categorizing Weight into Numeric Ranges (Kg)")
  
  df <- df %>%
    mutate(Weight_Category = case_when(
      Weight._kg_ < 57 ~ "<57kg",
      Weight._kg_ >= 57 & Weight._kg_ < 79 ~ "57-79kg",
      Weight._kg_ >= 79 & Weight._kg_ < 102 ~ "79-102kg",
      Weight._kg_ >= 102 ~ "≥102kg",
      TRUE ~ NA_character_
    ))
  
  
  return(df)
}
colnames(df_processed)

# Apply function
df_processed <- categorize_numeric_vars(df_processed)


# Load necessary libraries
library(dplyr)
library(tidyr)
library(broom)
library(nnet)

# Function for cross-tabulations with chi-square tests for dichotomized data
# Function for cross-tabulations with chi-square tests for dichotomized data
perform_crosstab_dicho <- function(df, row_var, col_var) {
  # Remove rows with missing data for the specific variables
  df_filtered <- df %>% filter(!is.na(.data[[row_var]]), !is.na(.data[[col_var]]))
  
  # Check if there are sufficient unique values to perform the test
  if (n_distinct(df_filtered[[row_var]]) < 2 | n_distinct(df_filtered[[col_var]]) < 2) {
    return(NULL) # Skip this crosstab if there aren't enough unique values
  }
  
  tbl <- table(df_filtered[[row_var]], df_filtered[[col_var]])
  
  # Ensure the table has at least 2 rows and 2 columns before performing chi-square
  if (nrow(tbl) < 2 | ncol(tbl) < 2) {
    return(NULL) # Skip the test if there aren't enough observations
  }
  
  # Perform chi-square test
  chi_sq <- chisq.test(tbl)
  
  tbl_row_percent <- prop.table(tbl, margin = 1) * 100
  tbl_col_percent <- prop.table(tbl, margin = 2) * 100
  
  df_counts <- as.data.frame(tbl)
  
  combined_table <- df_counts %>%
    rename(RowLevel = Var1, ColumnLevel = Var2, Count = Freq) %>%
    mutate(
      RowPercent = round(mapply(function(x, y) tbl_row_percent[x, y], RowLevel, ColumnLevel), 2),
      ColPercent = round(mapply(function(x, y) tbl_col_percent[x, y], RowLevel, ColumnLevel), 2),
      RowVariable = row_var,
      ColVariable = col_var,
      ChiSquare = chi_sq$statistic,
      p.value = chi_sq$p.value
    )
  
  return(combined_table)
}

colnames(df_processed)

names(df_processed) <- gsub(" ", "_", trimws(names(df_processed)))
names(df_processed) <- gsub("\\s+", "_", trimws(names(df_processed), whitespace = "[\\h\\v\\s]+"))
names(df_processed) <- gsub("\\(", "_", names(df_processed))
names(df_processed) <- gsub("\\)", "_", names(df_processed))
names(df_processed) <- gsub("\\-", "_", names(df_processed))
names(df_processed) <- gsub("/", "_", names(df_processed))
names(df_processed) <- gsub("\\\\", "_", names(df_processed)) 
names(df_processed) <- gsub("\\?", "", names(df_processed))
names(df_processed) <- gsub("\\'", "", names(df_processed))
names(df_processed) <- gsub("\\,", "_", names(df_processed))
names(df_processed) <- gsub("\\$", "", names(df_processed))
names(df_processed) <- gsub("\\+", "_", names(df_processed))

colnames(df_processed)

# Comprehensive crosstab definitions for all required analyses
# Comprehensive crosstab definitions for all required analyses

crosstabs <- list(
  # 1. Trauma Mechanism vs Circumstances of Death
  c("Trauma.M_BFT", "Circum.D_Suicide"),
  c("Trauma.M_SFT", "Circum.D_Homicide"),
  c("Trauma.M_GSW", "Circum.D_Homicide"),
  c("Trauma.M_THERMAL", "Circum.D_Accident"),
  c("Trauma.M_RTC", "Circum.D_Accident"),
  c("Trauma.M_SFT_Thermal", "Circum.D_Suicide"),
  
  # 2. Trauma Mechanism vs Skeletal Region (Cranial, Post-Cranial)
  c("Trauma.M_BFT", "Skeletal.Region_Cranial"),
  c("Trauma.M_BFT", "Skeletal.Region_Post_Cranial"),
  c("Trauma.M_GSW", "Skeletal.Region_Cranial"),
  c("Trauma.M_GSW", "Skeletal.Region_Post_Cranial"),
  c("Trauma.M_SFT", "Skeletal.Region_Cranial"),
  c("Trauma.M_SFT", "Skeletal.Region_Post_Cranial"),
  c("Trauma.M_THERMAL", "Skeletal.Region_Cranial"),
  c("Trauma.M_THERMAL", "Skeletal.Region_Post_Cranial"),
  
  # 3. Sex vs Circumstances of Death within Class
  c("Sex", "Class_Fall"),
  c("Sex", "Class_FireArms"),
  c("Sex", "Class_Asphyxia"),
  c("Sex", "Class_Aggresion"),
  c("Sex", "Class_Trauma"),
  c("Sex", "Class_Car__Conductor_"),
  c("Sex", "Class_Car__Pedestrian_"),
  c("Sex", "Class_Car__Passenger_seat_"),
  c("Sex", "Class_Motocycle"),
  c("Sex", "Class_Intoxication"),
  c("Sex", "Class_Asphyxia_by_drowning"),
  c("Sex", "Class_Asphyxia_by_chocking_on_food"),
  
  # 4. Sex vs Trauma Mechanism
  c("Sex", "Trauma.M_BFT"),
  c("Sex", "Trauma.M_GSW"),
  c("Sex", "Trauma.M_SFT"),
  c("Sex", "Trauma.M_THERMAL"),
  c("Sex", "Trauma.M_RTC"),
  
  # 5. Age Groups / Sex vs Trauma Mechanism
  c("Age_Group", "Trauma.M_BFT"),
  c("Age_Group", "Trauma.M_GSW"),
  c("Age_Group", "Trauma.M_SFT"),
  c("Age_Group", "Trauma.M_THERMAL"),
  c("Age_Group", "Trauma.M_RTC"),
  c("Sex", "Trauma.M_BFT"),
  c("Sex", "Trauma.M_GSW"),
  c("Sex", "Trauma.M_SFT"),
  c("Sex", "Trauma.M_THERMAL"),
  
  # 6. Age Groups / Sex vs Circumstances of Death
  c("Age_Group", "Circum.D_Suicide"),
  c("Age_Group", "Circum.D_Homicide"),
  c("Age_Group", "Circum.D_Accident"),
  c("Sex", "Circum.D_Suicide"),
  c("Sex", "Circum.D_Homicide"),
  c("Sex", "Circum.D_Accident"),
  
  # 7. Skeletal Region vs Circumstances of Death
  c("Skeletal.Region_Cranial", "Circum.D_Suicide"),
  c("Skeletal.Region_Cranial", "Circum.D_Homicide"),
  c("Skeletal.Region_Cranial", "Circum.D_Accident"),
  c("Skeletal.Region_Post_Cranial", "Circum.D_Suicide"),
  c("Skeletal.Region_Post_Cranial", "Circum.D_Homicide"),
  c("Skeletal.Region_Post_Cranial", "Circum.D_Accident"),
  
  # 8 & 9. Bone Side vs Circumstance of Death
  c("Bone.Side_Right_Side", "Circum.D_Suicide"),
  c("Bone.Side_Left_Side", "Circum.D_Suicide"),
  c("Bone.Side_No_Side", "Circum.D_Suicide"),
  c("Bone.Side_Right_Side", "Circum.D_Homicide"),
  c("Bone.Side_Left_Side", "Circum.D_Homicide"),
  c("Bone.Side_No_Side", "Circum.D_Homicide"),
  c("Bone.Side_Right_Side", "Circum.D_Accident"),
  c("Bone.Side_Left_Side", "Circum.D_Accident"),
  c("Bone.Side_No_Side", "Circum.D_Accident"),
  
  # 10. Class vs Skeletal Region/Subcategories
  c("Class_Fall", "Skeletal.Region_Cranial"),
  c("Class_Fall", "Skeletal.Region_Post_Cranial"),
  c("Class_FireArms", "Skeletal.Region_Cranial"),
  c("Class_FireArms", "Skeletal.Region_Post_Cranial"),
  c("Class_Car__Conductor_", "Skeletal.Region_Post_Cranial"),
  c("Class_Car__Pedestrian_", "Skeletal.Region_Post_Cranial"),
  c("Class_Car__Passenger_seat_", "Skeletal.Region_Post_Cranial"),
  
  # 11. Skeletal Cranial Subcategories vs Class and Circumstances of Death
  c("Skeletal.Cranial.Subcategory._Vault__Temporal", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Vault__Frontal", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Vault__Parietal", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Vault__Occipital", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Vault__Sphenoid", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Vault__Sphenoid", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Face__Orbit", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Face__Mandible", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Face__Maxilla", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Face__Ethmoid", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Face__Nasal", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Face__Vomer", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Face__Zygomatic_Arch", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Face__Zygomatic_Arch", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Face__Lacrimal_Bone", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Face__Palatine_Bone", "Circum.D_Homicide"),
  
  # 12. Skeletal Post-Cranial (Axial/Appendicular) vs Class and Circumstances
  c("Skeletal.Cranial.Subcategory._Axial__Ribs", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Axial__Ribs", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Axial__Vertebral_Column", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Axial__Vertebral_Column", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Axial__Vertebral_Column", "Circum.D_Accident"),
  c("Skeletal.Cranial.Subcategory._Axial__Sternum", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Axial__Sternum", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Axial__Sternum", "Circum.D_Accident"),
  c("Skeletal.Cranial.Subcategory._Axial__Thyroid_Cartilage", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Axial__Hyoid", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Axial__Sacrum_and_Coccyx", "Circum.D_Suicide"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Clavicle", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Scapula", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Scapula", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Radius", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Ulna", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Humerus", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Humerus", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Humerus", "Circum.D_Accident"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Metacarpals", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Phalanges__fingers_", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Hip_Bones", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Hip_Bones", "Circum.D_Accident"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Femur", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Femur", "Circum.D_Accident"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Patella", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Tibia", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Tibia", "Circum.D_Accident"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Fibula", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Fibula", "Circum.D_Accident"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Tarsal_Bones", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Metatarsals", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Metatarsals", "Circum.D_Homicide"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Phalanges__Toes_", "Class_Fall"),
  c("Skeletal.Cranial.Subcategory._Appendicular__Phalanges__Toes_", "Circum.D_Homicide"),
  
  # 13. Year / Sex vs Circumstances of Death
  c("Year", "Circum.D_Suicide"),
  c("Year", "Circum.D_Homicide"),
  c("Year", "Circum.D_Accident"),
  c("Sex", "Circum.D_Suicide"),
  c("Sex", "Circum.D_Homicide"),
  c("Sex", "Circum.D_Accident"),
  
  # 14. Circumstances / Trauma vs Height
  c("Height_Category", "Circum.D_Suicide"),
  c("Height_Category", "Circum.D_Homicide"),
  c("Height_Category", "Circum.D_Accident"),
  c("Height_Category", "Trauma.M_BFT"),
  c("Height_Category", "Trauma.M_GSW"),
  c("Height_Category", "Trauma.M_SFT"),
  
  # 15. Circumstances / Trauma vs Weight
  c("Weight_Category", "Circum.D_Suicide"),
  c("Weight_Category", "Circum.D_Homicide"),
  c("Weight_Category", "Circum.D_Accident"),
  c("Weight_Category", "Trauma.M_BFT"),
  c("Weight_Category", "Trauma.M_GSW"),
  c("Weight_Category", "Trauma.M_SFT")
)

crosstabs <- list(
c("Trauma.M_BFT", "Circum.D_Homicide"),
c("Trauma.M_BFT", "Circum.D_Accident"), # for your BFT-Accident hypothesis
c("Trauma.M_SFT", "Circum.D_Accident"),
c("Trauma.M_GSW", "Circum.D_Suicide"),
c("Trauma.M_THERMAL", "Circum.D_Suicide"),
c("Trauma.M_THERMAL", "Circum.D_Homicide"))


# Perform all crosstabs and combine results, skipping invalid cases
all_crosstabs <- bind_rows(lapply(crosstabs, function(vars) {
  result <- perform_crosstab_dicho(df_processed, vars[1], vars[2])
  if (!is.null(result)) return(result) # Only return non-null results
}))

colnames(df_processed)



# Create a vector mapping shorter column names back to the original dataset names
col_rename_map <- c(
  # Demographics
  "HeightCat" = "Height_Category",
  "WeightCat" = "Weight_Category",
  "AgeGroup" = "Age_Group",
  "Sex" = "Sex",
  
  # Trauma Mechanisms
  "Trauma_BFT" = "Trauma.M_BFT",
  "Trauma_SFT" = "Trauma.M_SFT",
  "Trauma_GSW" = "Trauma.M_GSW",
  "Trauma_Thermal" = "Trauma.M_THERMAL",
  "Trauma_RTC" = "Trauma.M_RTC",
  "Trauma_SFT_Thermal" = "Trauma.M_SFT_Thermal",
  
  # Circumstances of Death
  "Circum_Suicide" = "Circum.D_Suicide",
  "Circum_Homicide" = "Circum.D_Homicide",
  "Circum_Accident" = "Circum.D_Accident",
  
  # Skeletal Regions
  "Skeletal_Cranial" = "Skeletal.Region_Cranial",
  "Skeletal_PostCranial" = "Skeletal.Region_Post_Cranial",
  
  # Skeletal Cranial Subcategories
  "Vault_Temporal" = "Skeletal.Cranial.Subcategory._Vault__Temporal",
  "Vault_Frontal" = "Skeletal.Cranial.Subcategory._Vault__Frontal",
  "Vault_Parietal" = "Skeletal.Cranial.Subcategory._Vault__Parietal",
  "Vault_Occipital" = "Skeletal.Cranial.Subcategory._Vault__Occipital",
  "Vault_Sphenoid" = "Skeletal.Cranial.Subcategory._Vault__Sphenoid",
  "Face_Orbit" = "Skeletal.Cranial.Subcategory._Face__Orbit",
  "Face_Mandible" = "Skeletal.Cranial.Subcategory._Face__Mandible",
  "Face_Maxilla" = "Skeletal.Cranial.Subcategory._Face__Maxilla",
  "Face_Ethmoid" = "Skeletal.Cranial.Subcategory._Face__Ethmoid",
  "Face_Nasal" = "Skeletal.Cranial.Subcategory._Face__Nasal",
  "Face_Vomer" = "Skeletal.Cranial.Subcategory._Face__Vomer",
  "Face_Zygomatic" = "Skeletal.Cranial.Subcategory._Face__Zygomatic_Arch",
  "Face_Lacrimal" = "Skeletal.Cranial.Subcategory._Face__Lacrimal_Bone",
  "Face_Palatine" = "Skeletal.Cranial.Subcategory._Face__Palatine_Bone",
  
  # Skeletal Post-Cranial Subcategories (Axial & Appendicular)
  "Axial_Ribs" = "Skeletal.Cranial.Subcategory._Axial__Ribs",
  "Axial_Vertebrae" = "Skeletal.Cranial.Subcategory._Axial__Vertebral_Column",
  "Axial_Sternum" = "Skeletal.Cranial.Subcategory._Axial__Sternum",
  "Axial_Thyroid" = "Skeletal.Cranial.Subcategory._Axial__Thyroid_Cartilage",
  "Axial_Hyoid" = "Skeletal.Cranial.Subcategory._Axial__Hyoid",
  "Axial_Sacrum" = "Skeletal.Cranial.Subcategory._Axial__Sacrum_and_Coccyx",
  "Append_Clavicle" = "Skeletal.Cranial.Subcategory._Appendicular__Clavicle",
  "Append_Scapula" = "Skeletal.Cranial.Subcategory._Appendicular__Scapula",
  "Append_Radius" = "Skeletal.Cranial.Subcategory._Appendicular__Radius",
  "Append_Ulna" = "Skeletal.Cranial.Subcategory._Appendicular__Ulna",
  "Append_Humerus" = "Skeletal.Cranial.Subcategory._Appendicular__Humerus",
  "Append_Metacarpals" = "Skeletal.Cranial.Subcategory._Appendicular__Metacarpal",
  "Append_Carpals" = "Skeletal.Cranial.Subcategory._Appendicular__Carpals",
  "Append_Phal_Fingers" = "Skeletal.Cranial.Subcategory._Appendicular__Phalanges__fingers_",
  "Append_Hip" = "Skeletal.Cranial.Subcategory._Appendicular__Hip_Bones",
  "Append_Femur" = "Skeletal.Cranial.Subcategory._Appendicular__Femur",
  "Append_Patella" = "Skeletal.Cranial.Subcategory._Appendicular__Patella",
  "Append_Tibia" = "Skeletal.Cranial.Subcategory._Appendicular__Tibia",
  "Append_Fibula" = "Skeletal.Cranial.Subcategory._Appendicular__Fibula",
  "Append_Tarsals" = "Skeletal.Cranial.Subcategory._Appendicular__Tarsal_Bones",
  "Append_Metatarsals" = "Skeletal.Cranial.Subcategory._Appendicular__Metatarsals",
  "Append_Phal_Toes" = "Skeletal.Cranial.Subcategory._Appendicular__Phalanges__Toes_",
  
  # Bone Side
  "Bone_Right" = "Bone.Side_Right_Side",
  "Bone_Left" = "Bone.Side_Left_Side",
  "Bone_NoSide" = "Bone.Side_No_Side"
)


# Apply only the valid renaming
df_processed <- df_processed %>% rename(col_rename_map)

colnames(df_processed)

# Updated predictor list covering all analysis dimensions
predictors <- c(
  # Demographics
  "Sex", "AgeGroup", "HeightCat", "WeightCat",
  
  # Trauma Mechanisms
  "Trauma_BFT", "Trauma_SFT", "Trauma_GSW", "Trauma_Thermal", "Trauma_RTC", "Trauma_SFT_Thermal",
  
  # Skeletal Regions
  "Skeletal_Cranial", "Skeletal_PostCranial",
  
  # Skeletal Subcategories - Cranial
  "Vault_Temporal", "Vault_Frontal", "Vault_Parietal", "Vault_Occipital", "Vault_Sphenoid",
  "Face_Orbit", "Face_Mandible", "Face_Maxilla", "Face_Ethmoid", "Face_Nasal",
  "Face_Vomer", "Face_Zygomatic", "Face_Lacrimal", "Face_Palatine",
  
  # Skeletal Subcategories - Post-Cranial (Axial & Appendicular)
  "Axial_Ribs", "Axial_Vertebrae", "Axial_Sternum", "Axial_Thyroid",
  "Axial_Hyoid", "Axial_Sacrum",
  "Append_Clavicle", "Append_Scapula", "Append_Radius", "Append_Ulna",
  "Append_Humerus", "Append_Metacarpals", "Append_Phal_Fingers", "Append_Hip",
  "Append_Femur", "Append_Patella", "Append_Tibia", "Append_Fibula",
  "Append_Tarsals", "Append_Metatarsals", "Append_Phal_Toes", "Append_Carpals",
  
  # Bone Side
  "Bone_Right", "Bone_Left", "Bone_NoSide"
)


run_logistic_model <- function(df, outcome_var, predictors) {
  # Remove rows where the outcome variable is NA
  df <- df[!is.na(df[[outcome_var]]), ]
  
  # Ensure that the outcome variable has both 0 and 1
  if (length(unique(df[[outcome_var]])) < 2) {
    message(paste("Skipping model for", outcome_var, "- only one class present."))
    return(NULL)
  }
  
  # Identify predictors with sufficient variability and non-missing values
  valid_predictors <- predictors[sapply(df[predictors], function(col) {
    if (all(is.na(col))) return(FALSE)  # Exclude predictors with all NA values
    if (length(unique(na.omit(col))) <= 1) return(FALSE)  # Exclude predictors with no variability
    return(TRUE)
  })]
  
  # Ensure at least one predictor remains
  if (length(valid_predictors) == 0) {
    message(paste("No valid predictors remaining for", outcome_var, "- skipping model."))
    return(NULL)
  }
  
  # Check for predictors that perfectly predict the outcome
  perfect_predictors <- valid_predictors[sapply(valid_predictors, function(var) {
    length(unique(df[[var]][df[[outcome_var]] == 1])) == 1 | 
      length(unique(df[[var]][df[[outcome_var]] == 0])) == 1
  })]
  
  # Remove perfect predictors
  valid_predictors <- setdiff(valid_predictors, perfect_predictors)
  
  if (length(valid_predictors) == 0) {
    message(paste("All predictors perfectly predict", outcome_var, "- skipping model."))
    return(NULL)
  }
  
  # Construct formula dynamically
  formula <- as.formula(paste(outcome_var, "~", paste(valid_predictors, collapse = "+")))
  
  # Fit logistic regression model with error handling
  model <- tryCatch(
    {
      glm(formula, data = df, family = binomial)
    },
    error = function(e) {
      message("Error fitting model for ", outcome_var, ": ", e$message)
      return(NULL)
    }
  )
  
  if (is.null(model)) {
    return(NULL)
  }
  
  # Extract model summary and fit indices
  model_summary <- tidy(model, exponentiate = TRUE, conf.int = TRUE)
  fit_indices <- glance(model)
  
  return(list(summary = model_summary, fit = fit_indices))
}




# Run logistic regression models using the updated predictors
logistic_suicide <- run_logistic_model(df_processed, "Circum_Suicide", predictors)
logistic_homicide <- run_logistic_model(df_processed, "Circum_Homicide", predictors)
logistic_accident <- run_logistic_model(df_processed, "Circum_Accident", predictors)

# Combine logistic results into tidy dataframe
logistic_results <- bind_rows(
  logistic_suicide$summary %>% mutate(Model = "Suicide"),
  logistic_homicide$summary %>% mutate(Model = "Homicide"),
  logistic_accident$summary %>% mutate(Model = "Accident")
)

# Print model fit indices
fit_indices <- bind_rows(
  logistic_suicide$fit %>% mutate(Model = "Suicide"),
  logistic_homicide$fit %>% mutate(Model = "Homicide"),
  logistic_accident$fit %>% mutate(Model = "Accident")
)

print(logistic_results)
print(fit_indices)


colnames(df_processed)


# Example usage
data_list <- list(
  "Crosstabs" = all_crosstabs, 
  "Logit Models" = logistic_results,
  "Fit Indices" = fit_indices
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_v5.xlsx")



# Load the package
library(writexl)

# Export to Excel
write_xlsx(all_crosstabs, "all_crosstabs2.xlsx")



crosstabs <- list(
  c("Trauma.M_BFT", "Circum.D_Homicide"),
  c("Trauma.M_BFT", "Circum.D_Accident"), # for your BFT-Accident hypothesis
  c("Trauma.M_SFT", "Circum.D_Accident"),
  c("Trauma.M_GSW", "Circum.D_Suicide"),
  c("Trauma.M_THERMAL", "Circum.D_Suicide"),
  c("Trauma.M_THERMAL", "Circum.D_Homicide"),
  c("Trauma.M_SFT", "Circum.D_Suicide"))


# Perform all crosstabs and combine results, skipping invalid cases
all_crosstabs <- bind_rows(lapply(crosstabs, function(vars) {
  result <- perform_crosstab_dicho(df_processed, vars[1], vars[2])
  if (!is.null(result)) return(result) # Only return non-null results
}))
