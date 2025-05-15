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

# Load necessary libraries
library(dplyr)
library(tidyr)
library(broom)
library(nnet)

# Function for cross-tabulations with chi-square tests for dichotomized data
perform_crosstab_dicho <- function(df, row_var, col_var) {
  tbl <- table(df[[row_var]], df[[col_var]])
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

# Comprehensive crosstab definitions for all required analyses
crosstabs <- list(
  # 1. Trauma mechanism vs Circumstances of Death
  c("Trauma.M_BFT", "Circum.D_Suicide"),
  c("Trauma.M_SFT", "Circum.D_Homicide"),
  c("Trauma.M_GSW", "Circum.D_Homicide"),
  c("Trauma.M_Thermal", "Circum.D_Accident"),
  
  # 2. Trauma mechanism vs Skeletal Region (Cranial, Post-Cranial)
  c("Trauma.M_BFT", "Skeletal.R_Cranial"),
  c("Trauma.M_BFT", "Skeletal.R_Post"),
  
  # 3. Sex vs Circumstances of Death within Class (examples: Fall, Firearms)
  c("Sex", "Class_Fall"),
  c("Sex", "Class_FireArms"),
  
  # 4. Sex vs Trauma Mechanism
  c("Sex", "Trauma.M_BFT"),
  c("Sex", "Trauma.M_GSW"),
  
  # 5. Age groups / Sex vs Trauma mechanism (example shown for BFT)
  c("Age_Group", "Trauma.M_BFT"),
  c("Sex", "Trauma.M_BFT"),
  
  # 6. Age groups / Sex vs Circumstances of Death
  c("Age_Group", "Circum.D_Suicide"),
  c("Sex", "Circum.D_Suicide"),
  
  # 7. Skeletal region vs Circumstances of Death
  c("Skeletal.R_Cranial", "Circum.D_Suicide"),
  c("Skeletal.R_Post", "Circum.D_Suicide"),
  
  # 8 & 9. Side / Skeletal Region vs Circumstance of Death (examples for sides)
  c("Side_1", "Circum.D_Suicide"),
  c("Side_2", "Circum.D_Suicide"),
  
  # 10. Class vs Skeletal Region/Subcategories
  c("Class_Fall", "Skeletal.R_Cranial"),
  c("Class_Fall", "Skeletal.R_Post"),
  
  # 11. Skeletal cranial subcategory vs Class and Circumstances of Death
  c("Sub.Cat.Skeletal._Face__1", "Class_Fall"),
  c("Sub.Cat.Skeletal._Vault__1", "Circum.D_Suicide"),
  
  # 12. Skeletal post-cranial (axial/appendicular) vs Class and circumstances
  c("Axial_3", "Class_Fall"),
  c("Appendicular_4", "Circum.D_Suicide"),
  
  # 13. Year/Sex vs Circumstances of death
  c("Year", "Circum.D_Suicide"),
  c("Sex", "Circum.D_Homicide"),
  
  # 14. Circumstances/Trauma vs Height
  c("Stature_Category", "Circum.D_Suicide"),
  c("Stature_Category", "Trauma.M_BFT"),
  
  # 15. Circumstances/Trauma vs Pounds
  c("Weight_Category", "Circum.D_Suicide"),
  c("Weight_Category", "Trauma.M_BFT")
)


# Perform all crosstabs and combine results
all_crosstabs <- bind_rows(lapply(crosstabs, function(vars) perform_crosstab_dicho(df_processed, vars[1], vars[2])))


# Logistic regression models (binary for each circumstance)
run_logistic_model <- function(df, outcome_var, predictors) {
  formula <- as.formula(paste(outcome_var, "~", paste(predictors, collapse = "+")))
  model <- glm(formula, data = df, family = binomial)
  model_summary <- tidy(model, exponentiate = TRUE, conf.int = TRUE)
  fit_indices <- glance(model)
  return(list(summary = model_summary, fit = fit_indices))
}

predictors <- c("Sex", "Age_Group", "Stature_Category", "Weight_Category", "Trauma.M_BFT", "Trauma.M_SFT", "Trauma.M_GSW", 
                "Trauma.M_Thermal", "Skeletal.R_Cranial", "Skeletal.R_Post", "Sub.Cat.Skeletal._Face__1", "Sub.Cat.Skeletal._Vault__1",
                "Axial_3", "Appendicular_4", "Side_1", "Side_2")

logistic_suicide <- run_logistic_model(df_processed, "Circum.D_Suicide", predictors)
logistic_homicide <- run_logistic_model(df_processed, "Circum.D_Homicide", predictors)
logistic_accident <- run_logistic_model(df_processed, "Circum.D_Accident", predictors)

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




# Example usage
data_list <- list(
  "Crosstabs" = all_crosstabs, 
  "Logit Models" = logistic_results,
  "Fit Indices" = fit_indices
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_v3.xlsx")
