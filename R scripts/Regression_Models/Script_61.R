# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/welletssee")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Upload for analysis Fiverr (1).xlsx", sheet = "Culture positive scrapes")

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
str(df)

library(dplyr)
library(lubridate)

# Convert Excel date number to Date format
df$Date.Of.Specimen <- as.Date(df$Date.Of.Specimen, origin = "1899-12-30")

# Define the periods based on date ranges
df <- df %>%
  mutate(Period = case_when(
    Date.Of.Specimen >= as.Date("2014-01-01") & Date.Of.Specimen <= as.Date("2020-03-23") ~ "Pre-lockdown",
    Date.Of.Specimen >= as.Date("2020-03-24") & Date.Of.Specimen <= as.Date("2021-09-30") ~ "Lockdown",
    Date.Of.Specimen >= as.Date("2021-10-01") & Date.Of.Specimen <= as.Date("2023-12-31") ~ "Post-lockdown"
  ))


library(dplyr)

# Function to analyze antibiotic trends across periods
analyze_antibiotic <- function(df, antibiotic_name) {
  df %>%
    filter(!is.na(get(antibiotic_name)), get(antibiotic_name) != "Not tested") %>%
    group_by(Period, Antibiotic_Response = get(antibiotic_name)) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    mutate(Percentage = Count / sum(Count) * 100) %>%
    arrange(Period, Antibiotic_Response)
}

# Apply this function to each antibiotic and combine the results
antibiotic_list <- c("Amikacin", "Azithromycin", "Ceftazidime", "Cefuroxime", 
                     "Chloramphenicol", "Ciprofloxacin", "Clarithromycin", "Clindamycin",
                     "Cotrimoxazole", "Doxycycline", "Fusidic.Acid", "Gentamicin",
                     "Levofloxacin", "Moxifloxacin", "Mupirocin", "Ofloxacin", 
                     "Tobramycin", "Trimethoprim", "Vancomycin")

# Function to standardize responses
standardize_responses <- function(df, antibiotic_cols) {
  df %>%
    mutate(across(all_of(antibiotic_cols), ~ ifelse(. == "s", "S", .)))
}

# Apply the standardization function to your dataframe
df <- standardize_responses(df, antibiotic_list)

# Create an empty data frame to store results
antibiotic_trends <- data.frame()

# Loop through each antibiotic and bind the results
for (antibiotic in antibiotic_list) {
  antibiotic_data <- analyze_antibiotic(df, antibiotic)
  antibiotic_data$Antibiotic <- antibiotic  # Add the antibiotic name to the results
  antibiotic_trends <- rbind(antibiotic_trends, antibiotic_data)
}

# View the results
print(antibiotic_trends)


library(dplyr)

# Calculate the distribution of Organism Types across periods
organism_type_trends <- df %>%
  group_by(Period, Organism.type) %>%
  summarise(Count = n(), .groups = 'drop') %>%
  mutate(Percentage = Count / sum(Count) * 100)

# View the results
print(organism_type_trends)

# Function to analyze trends for a specific antibiotic by Organism Type
analyze_by_organism_and_antibiotic <- function(df, antibiotic_name) {
  df %>%
    filter(!is.na(get(antibiotic_name)), get(antibiotic_name) != "Not tested") %>%
    group_by(Period, Organism.type, Antibiotic_Response = get(antibiotic_name)) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    mutate(Percentage = Count / sum(Count) * 100) %>%
    arrange(Period, Organism.type, Antibiotic_Response)
}

# Create an empty data frame to store results
organism_antibiotic_trends <- data.frame()

# Loop through each antibiotic and bind the results
for (antibiotic in antibiotic_list) {
  antibiotic_data <- analyze_by_organism_and_antibiotic(df, antibiotic)
  antibiotic_data$Antibiotic <- antibiotic  # Add the antibiotic name to the results
  organism_antibiotic_trends <- rbind(organism_antibiotic_trends, antibiotic_data)
}

# View the results
print(organism_antibiotic_trends)

library(dplyr)
library(tidyr)

# Gather all antibiotic columns into a long format
df_long <- df %>%
  select(Period, all_of(antibiotic_list)) %>%
  pivot_longer(cols = -Period, names_to = "Antibiotic", values_to = "Response") %>%
  filter(Response %in% c("R", "S", "I"))  # Filter out 'Not tested' and NA values

# Calculate the percentage of each response type (R, S, I) across periods
antibiotic_response_trends <- df_long %>%
  group_by(Period, Response) %>%
  summarise(Count = n(), .groups = 'drop') %>%
  group_by(Period) %>%
  mutate(Total = sum(Count)) %>%
  ungroup() %>%
  mutate(Percentage = (Count / Total) * 100)

# View the results
print(antibiotic_response_trends)



# Chi-Square tests
chi_square_analysis_pre_post <- function(data, row_vars) {
  results <- list()  # Initialize an empty list to store results
  
  # Filter to only Pre-lockdown and Post-lockdown
  data <- data %>%
    filter(Period %in% c("Pre-lockdown", "Post-lockdown")) %>%
    droplevels()
  
  # Iterate over row variables
  for (row_var in row_vars) {
    cat("Processing row variable:", row_var, "\n")
    
    # Filter out 'Not tested', NA values, and 'I' responses
    modified_data <- data %>%
      filter(!is.na(.[[row_var]]), .[[row_var]] != "Not tested", .[[row_var]] != "I") %>%
      mutate(!!row_var := factor(.[[row_var]], levels = c("R", "S")))
    
    cat("Levels in", row_var, "after filtering:", levels(modified_data[[row_var]]), "\n")
    
    # Create a contingency table for the Pre-lockdown vs. Post-lockdown
    contingency_table <- table(modified_data[[row_var]], modified_data$Period, dnn = c(row_var, "Period"))
    
    cat("Contingency Table for", row_var, ":\n")
    print(contingency_table)
    
    # Apply Fisher's Exact Test if the table dimensions are 2x2
    if (ncol(contingency_table) > 0 && all(dim(contingency_table) == 2)) {
      fisher_test <- fisher.test(contingency_table)
      p_val = fisher_test$p.value
      
      # Create data frame with results
      for (lvl in levels(modified_data[[row_var]])) {
        for (per in colnames(contingency_table)) {
          level_df <- data.frame(
            Row_Variable = row_var,
            Row_Level = lvl,
            Period = per,
            Period_Comparison = "Pre-lockdown vs. Post-lockdown",
            Count = as.numeric(contingency_table[lvl, per]),
            Percent = as.numeric(prop.table(contingency_table, margin = 2)[lvl, per] * 100),
            P_Value = p_val,
            stringsAsFactors = FALSE
          )
          # Add the result to the list
          results[[paste0(row_var, "_", per, "_", lvl)]] <- level_df
        }
      }
    } else {
      cat("Not suitable for Fisher's Exact Test due to table dimensions or data sparsity for", row_var, "\n")
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}

row_variables <- antibiotic_list

# Apply the function to the dataframe
df_fisher_antibiotic_period <- chi_square_analysis_pre_post(df, row_variables)
print(df_fisher_antibiotic_period)


# Percentages relative to the column
chi_square_analysis_multiple <- function(data, row_vars, col_var) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over row variables
  for (row_var in row_vars) {
    # Ensure that the variables are factors
    data[[row_var]] <- factor(data[[row_var]])
    data[[col_var]] <- factor(data[[col_var]])
    
    # Create a crosstab with percentages
    crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 2) * 100
    
    # Check if the crosstab is 2x2
    is_2x2_table <- all(dim(table(data[[row_var]], data[[col_var]])) == 2)
    
    # Perform chi-square test with correction for 2x2 tables
    chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]), correct = is_2x2_table)
    
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
      
      level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
      level_df$Chi_Square <- chi_square_test$statistic
      level_df$P_Value <- chi_square_test$p.value
      
      # Add the result to the list
      results[[paste0(row_var, "_", level)]] <- level_df
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}

row_variables <- "Organism.type"
column_variable <- "Period"

df_chisq_organism <- chi_square_analysis_multiple(df, row_variables, column_variable)
print(df_chisq_organism)


df_chisq_combinedantibiotics <- data.frame(
  Period = c("Lockdown", "Lockdown", "Post-lockdown", "Post-lockdown", "Post-lockdown", "Pre-lockdown", "Pre-lockdown", "Pre-lockdown"),
  Response = c("R", "S", "I", "R", "S", "I", "R", "S"),
  Count = c(29, 414, 63, 159, 1051, 46, 395, 2500),
  Total = c(443, 443, 1273, 1273, 1273, 2941, 2941, 2941),
  Percentage = c(6.546275, 93.453725, 4.948940, 12.490181, 82.560880, 1.564094, 13.430806, 85.005100)
)

# Create the contingency table
contingency_table <- xtabs(Count ~ Response + Period, data = df_chisq_combinedantibiotics)
print("Contingency Table:")
print(contingency_table)

# Perform the Chi-squared test
chi_test <- chisq.test(contingency_table)

# Print Chi-squared results
cat("Chi-squared Statistic:", chi_test$statistic, "\n")
cat("P-value:", chi_test$p.value, "\n")

# Append the Chi-squared statistic and p-value to the original data frame
df_chisq_combinedantibiotics$Chi_squared_statistic <- chi_test$statistic
df_chisq_combinedantibiotics$Chi_squared_p_value <- chi_test$p.value

# Show the updated data frame
print(df_chisq_combinedantibiotics)




# Organism type and antibiotics

chi_square_analysis_pre_post <- function(data, row_vars, organism_types) {
  results <- list()  # Initialize an empty list to store results
  
  # Iterate over each organism type
  for (organism_type in organism_types) {
    cat("Processing for Organism Type:", organism_type, "\n")
    
    # Filter data for the current organism type
    data_filtered <- data %>%
      filter(Organism.type == organism_type, Period %in% c("Pre-lockdown", "Post-lockdown")) %>%
      droplevels()
    
    # Iterate over row variables (antibiotics)
    for (row_var in row_vars) {
      cat("Processing row variable:", row_var, "\n")
      
      # Filter out 'Not tested', NA values, and 'I' responses
      modified_data <- data_filtered %>%
        filter(!is.na(.[[row_var]]), .[[row_var]] != "Not tested", .[[row_var]] != "I") %>%
        mutate(!!row_var := factor(.[[row_var]], levels = c("R", "S")))
      
      cat("Levels in", row_var, "after filtering:", levels(modified_data[[row_var]]), "\n")
      
      # Create a contingency table for the Pre-lockdown vs. Post-lockdown
      contingency_table <- table(modified_data[[row_var]], modified_data$Period, dnn = c(row_var, "Period"))
      
      cat("Contingency Table for", row_var, ":\n")
      print(contingency_table)
      
      # Apply Fisher's Exact Test if the table dimensions are 2x2
      if (ncol(contingency_table) > 0 && all(dim(contingency_table) == 2)) {
        fisher_test <- fisher.test(contingency_table)
        p_val = fisher_test$p.value
        
        # Create data frame with results
        for (lvl in levels(modified_data[[row_var]])) {
          for (per in colnames(contingency_table)) {
            level_df <- data.frame(
              Organism_Type = organism_type,
              Row_Variable = row_var,
              Row_Level = lvl,
              Period = per,
              Period_Comparison = "Pre-lockdown vs. Post-lockdown",
              Count = as.numeric(contingency_table[lvl, per]),
              Percent = as.numeric(prop.table(contingency_table, margin = 2)[lvl, per] * 100),
              P_Value = p_val,
              stringsAsFactors = FALSE
            )
            # Add the result to the list
            results[[paste0(organism_type, "_", row_var, "_", per, "_", lvl)]] <- level_df
          }
        }
      } else {
        cat("Not suitable for Fisher's Exact Test due to table dimensions or data sparsity for", row_var, "\n")
      }
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}

row_variables <- antibiotic_list
organism_types <- c("Gram Positive bacteria", "Gram Negative bacteria", "Protozoa", "Fungus")  # Specify your organism types

# Apply the function to the dataframe
df_fisher_antibiotic_organism_period <- chi_square_analysis_pre_post(df, row_variables, organism_types)
print(df_fisher_antibiotic_organism_period)


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
  "Combined Antibiotics" = antibiotic_response_trends,
  "All antibiotics" = antibiotic_trends,
  "Organism and Antibiotics" = organism_antibiotic_trends,
  "Organism Only" = organism_type_trends,
  "Chisq - Combined antibiotics" = df_chisq_combinedantibiotics,
  "Chisq - Organism only" = df_chisq_organism,
  "Fisher - All antibiotics" = df_fisher_antibiotic_period,
  "Fisher - Org and Antibiotics" = df_fisher_antibiotic_organism_period
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")




# Read data from a specific sheet of an Excel file
df_negative <- read.xlsx("Upload for analysis Fiverr (1).xlsx", sheet = "Culture negative scrapes")


# Convert Excel date number to Date format
df_negative$Date.Of.Specimen <- as.Date(df_negative$Date.Of.Specimen, origin = "1899-12-30")

# Define the periods based on date ranges
df_negative <- df_negative %>%
  mutate(Period = case_when(
    Date.Of.Specimen >= as.Date("2014-01-01") & Date.Of.Specimen <= as.Date("2020-03-23") ~ "Pre-lockdown",
    Date.Of.Specimen >= as.Date("2020-03-24") & Date.Of.Specimen <= as.Date("2021-09-30") ~ "Lockdown",
    Date.Of.Specimen >= as.Date("2021-10-01") & Date.Of.Specimen <= as.Date("2023-12-31") ~ "Post-lockdown"
  ))


str(df_negative)

# Add a 'Result' column indicating positive and negative cultures
df$Result <- "Positive"
df_negative$Result <- "Negative"

# Combine the two dataframes
df_combined <- bind_rows(df, df_negative)

# Check structure
str(df_combined)

# Calculate the count and percentage of positive/negative results per period
df_result_trends <- df_combined %>%
  group_by(Period, Result) %>%
  summarise(Count = n(), .groups = 'drop') %>%
  group_by(Period) %>%
  mutate(Total = sum(Count),
         Percentage = (Count / Total) * 100)

# View the result
print(df_result_trends)

library(openxlsx)

# Define the function
chi_square_analysis_multiple <- function(data, row_var, col_var) {
  results <- list() # Initialize an empty list to store results
  
  # Ensure that the variables are factors
  data[[row_var]] <- factor(data[[row_var]])
  data[[col_var]] <- factor(data[[col_var]])
  
  # Create a crosstab with percentages
  crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 2) * 100
  
  # Check if the crosstab is 2x2
  is_2x2_table <- all(dim(table(data[[row_var]], data[[col_var]])) == 2)
  
  # Perform chi-square test with correction for 2x2 tables
  chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]), correct = is_2x2_table)
  
  # Convert crosstab to a dataframe
  crosstab_df <- as.data.frame.matrix(crosstab)
  
  # Create a dataframe for the variable
  for (level in levels(data[[row_var]])) {
    level_df <- data.frame(
      "Row_Variable" = row_var,
      "Row_Level" = level,
      "Column_Variable" = col_var,
      check.names = FALSE
    )
    
    level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
    level_df$Chi_Square <- chi_square_test$statistic
    level_df$P_Value <- chi_square_test$p.value
    
    # Add the result to the list
    results[[paste0(row_var, "_", level)]] <- level_df
  }
  
  # Combine all results into a single dataframe
  return(do.call(rbind, results))
}

# Apply the function to calculate the chi-square results for the "Result" (Positive/Negative) across "Period"
df_chisq_results <- chi_square_analysis_multiple(df_combined, "Result", "Period")

# Export the Chi-Square results to an Excel file
save_chisq_results <- function(data, filename) {
  wb <- createWorkbook()
  addWorksheet(wb, "Chi-Square Results")
  
  writeData(wb, "Chi-Square Results", data)
  
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Save the results to Excel
save_chisq_results(df_chisq_results, "Chi_Square_Results_Positive_Negative_Period.xlsx")

# View the results
print(df_chisq_results)


# Calculate mean age and age range per period for the df (which already contains positive scrapes)
age_stats <- df %>%
  group_by(Period) %>%
  summarise(
    Mean_Age = mean(Age.of.patient, na.rm = TRUE),
    Min_Age = min(Age.of.patient, na.rm = TRUE),
    Max_Age = max(Age.of.patient, na.rm = TRUE),
    Age_Range = Max_Age - Min_Age
  )

# View the results
print(age_stats)

# Perform ANOVA to check for statistically significant differences in mean age across periods
anova_result <- aov(Age.of.patient ~ Period, data = df)

# Get the summary of the ANOVA result
summary(anova_result)

# Combine the age statistics with the ANOVA p-value
anova_p_value <- summary(anova_result)[[1]]$`Pr(>F)`[1]

# Add ANOVA result to the statistics table
age_stats$ANOVA_P_Value <- anova_p_value

# Export the results to an Excel file
save_age_stats <- function(data, filename) {
  wb <- createWorkbook()
  addWorksheet(wb, "Age Statistics")
  
  writeData(wb, "Age Statistics", data)
  
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Save the age statistics to Excel
save_age_stats(age_stats, "Age_Statistics_Period_Positive_Scrapes.xlsx")



# Print all distinct values of the Organism.Result column
distinct_organism_results <- unique(df$Organism.Result)

# Print the distinct values
print(distinct_organism_results)



# Load necessary libraries
library(dplyr)

# Define the organisms to filter and group subspecies (with names matching the dataframe)
species_groups <- c(
  # Gram Positives
  "Staphylococcus epidermidis", "Staphylococcus aureus", "Staphylococcuscapitis", 
  "Staphylococcus warneri", "Streptococcus pneumoniae", "Staphylococcus hominis", 
  "Staphylococcus pasteuri", "Streptococcusmitis/oralis", "Micrococcus luteus", "Micrococcus sp",
  
  # Gram Negatives
  "Pseudomonas aeruginosa", "Moraxella osloensis", "Moraxella nonliquefaciens", 
  "Moraxella lacunata", "Serratia liquifaciens", "Serratia marcescens", "Serratia species",
  
  # Protozoa
  "Acanthamoeba sp",
  
  # Fungi
  "Aspergillus fumigatus", "Candida albicans", "Candida parapsilosis", 
  "Exophiala dermatitidis", "Fusarium spp"
)

grouped_species <- c(
  "Staphylococcus epidermidis", "Staphylococcus aureus", "Staphylococcuscapitis", 
  "Staphylococcus warneri", "Streptococcus pneumoniae", "Staphylococcus hominis", 
  "Staphylococcus pasteuri", "Streptococcusmitis/oralis", "Micrococcus",  # Grouped Micrococcus species
  
  "Pseudomonas aeruginosa", "Moraxella",  # Grouped Moraxella species
  "Serratia",  # Grouped Serratia species
  
  "Acanthamoeba sp",  # Protozoa
  
  "Aspergillus fumigatus", "Candida albicans", "Candida parapsilosis", 
  "Exophiala dermatitidis", "Fusarium spp"  # Fungus
)

# Group subspecies (Micrococcus, Moraxella, Serratia)
df_grouped <- df %>%
  mutate(Organism_Group = case_when(
    Organism.Result %in% c("Micrococcus luteus", "Micrococcus sp") ~ "Micrococcus",
    Organism.Result %in% c("Moraxella osloensis", "Moraxella nonliquefaciens", "Moraxella lacunata") ~ "Moraxella",
    Organism.Result %in% c("Serratia liquifaciens", "Serratia marcescens", "Serratia species") ~ "Serratia",
    TRUE ~ Organism.Result
  )) %>%
  # Filter only the species of interest from the grouped list
  filter(Organism_Group %in% grouped_species)


# View distinct values to confirm correct filtering
distinct_species <- unique(df_grouped$Organism_Group)
print(distinct_species)


# Calculate percentage of each species relative to the total count of species during each period
species_percentages <- df_grouped %>%
  group_by(Period, Organism_Group) %>%
  summarise(Count = n(), .groups = 'drop') %>%
  group_by(Period) %>%
  mutate(Total = sum(Count),
         Percentage = (Count / Total) * 100) %>%
  arrange(Period, Organism_Group)

# View the calculated percentages
print(species_percentages)

# Create a contingency table of species counts across periods
contingency_table <- table(df_grouped$Organism_Group, df_grouped$Period)

# Perform the Chi-Square test
chi_test <- chisq.test(contingency_table)

# Output Chi-Square results
cat("Chi-Square Statistic:", chi_test$statistic, "\n")
cat("P-value:", chi_test$p.value, "\n")


# Function to export the percentages and Chi-Square test results to an Excel file
save_species_percentages <- function(data, chi_test, filename) {
  wb <- createWorkbook()
  addWorksheet(wb, "Species Percentages")
  
  # Write species percentages data
  writeData(wb, "Species Percentages", data)
  
  # Write Chi-Square test results
  addWorksheet(wb, "Chi-Square Results")
  chi_test_results <- data.frame(
    "Chi-Square Statistic" = chi_test$statistic,
    "P-value" = chi_test$p.value
  )
  writeData(wb, "Chi-Square Results", chi_test_results)
  
  # Save the workbook
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Save the results to Excel
save_species_percentages(species_percentages, chi_test, "Species_Percentages_and_Chi_Square.xlsx")

# View the final percentages table
print(species_percentages)

# Function to analyze antibiotic trends across periods for a given species and antibiotic
analyze_antibiotic_by_species <- function(df, species_name, antibiotic_name) {
  df %>%
    filter(Organism_Group == species_name, !is.na(get(antibiotic_name)), get(antibiotic_name) %in% c("R", "S", "I")) %>%
    group_by(Period, Antibiotic_Response = get(antibiotic_name)) %>%
    summarise(Count = n(), .groups = 'drop') %>%
    complete(Antibiotic_Response = c("R", "S", "I"), Period = unique(df$Period), fill = list(Count = 0)) %>%
    # Calculate the percentage relative to the total responses in each period
    group_by(Period) %>%
    mutate(Percentage = Count / sum(Count) * 100) %>%
    arrange(Period, Antibiotic_Response)
}



# Create an empty data frame to store results
antibiotic_sensitivity_results <- data.frame()

# Loop through each species and antibiotic to calculate sensitivity percentages
for (species in distinct_species) {
  for (antibiotic in antibiotic_list) {
    species_antibiotic_data <- analyze_antibiotic_by_species(df_grouped, species, antibiotic)
    if (nrow(species_antibiotic_data) > 0) {  # Only include non-empty tables
      species_antibiotic_data$Species <- species  # Add the species name to the results
      species_antibiotic_data$Antibiotic <- antibiotic  # Add the antibiotic name to the results
      antibiotic_sensitivity_results <- rbind(antibiotic_sensitivity_results, species_antibiotic_data)
    }
  }
}

# View the results
print(antibiotic_sensitivity_results)


# Function to perform Chi-Square or Fisher's Exact Test comparing the full distribution of Antibiotic Response (R, S, I) across periods
perform_chi_square_or_fisher <- function(df, species_name, antibiotic_name) {
  # Filter for the species and antibiotic response (R, S, I)
  df_filtered <- df %>%
    filter(Organism_Group == species_name, !is.na(get(antibiotic_name)), get(antibiotic_name) %in% c("R", "S", "I"))
  
  # Create a contingency table comparing Periods to Antibiotic Sensitivity (R, S, I)
  contingency_table <- table(df_filtered$Period, df_filtered[[antibiotic_name]])
  
  # Print the contingency table to see how Periods and Responses are distributed
  cat("\nContingency table for species:", species_name, "and antibiotic:", antibiotic_name, "\n")
  print(contingency_table)
  
  # Check if the table has at least 2 rows (Periods) and 2 columns (Antibiotic Responses with non-zero counts)
  if (nrow(contingency_table) < 2 | ncol(contingency_table) < 2) {
    return(list(p.value = NA, test_type = NA, statistic = NA, periods = unique(df_filtered$Period)))  # Skip this test if there's not enough variation
  }
  
  # If the contingency table has any 0 counts, Fisher's Exact Test is preferred
  if (any(contingency_table == 0)) {
    test_result <- tryCatch({
      test <- fisher.test(contingency_table)
      list(p.value = test$p.value, statistic = NA, test_type = "Fisher's Exact Test", periods = unique(df_filtered$Period))
    }, error = function(e) {
      return(list(p.value = NA, test_type = NA, statistic = NA, periods = unique(df_filtered$Period)))
    })
  } else {
    # If no 0 counts, perform Chi-Square Test
    test_result <- tryCatch({
      test <- chisq.test(contingency_table)
      list(p.value = test$p.value, statistic = test$statistic, test_type = "Chi-Square Test", periods = unique(df_filtered$Period))
    }, error = function(e) {
      return(list(p.value = NA, test_type = NA, statistic = NA, periods = unique(df_filtered$Period)))
    })
  }
  
  return(test_result)
}




# Create an empty data frame to store the test results
chi_square_fisher_results <- data.frame()

# Loop through each species and antibiotic to perform the test comparing the full distribution of R, S, I across periods
for (species in distinct_species) {
  for (antibiotic in antibiotic_list) {
    test_result <- perform_chi_square_or_fisher(df_grouped, species, antibiotic)
    
    # Store the periods involved as a concatenated string
    periods_involved <- paste(test_result$periods, collapse = ", ")
    
    test_data <- data.frame(Species = species, Antibiotic = antibiotic, 
                            P_Value = test_result$p.value, 
                            Test_Type = test_result$test_type,
                            Chi_Square_Stat = test_result$statistic,
                            Periods_Involved = periods_involved)
    
    chi_square_fisher_results <- rbind(chi_square_fisher_results, test_data)
  }
}

# Merge the test results back into the antibiotic sensitivity table
antibiotic_sensitivity_results <- antibiotic_sensitivity_results %>%
  left_join(chi_square_fisher_results, by = c("Species", "Antibiotic"))

# View the final results with 0 counts and test results
print(antibiotic_sensitivity_results)


# Define the function to export the results to Excel
save_antibiotic_sensitivity_results <- function(data, filename) {
  # Create a new Excel workbook
  wb <- createWorkbook()
  
  # Add a worksheet named "Antibiotic Sensitivity"
  addWorksheet(wb, "Antibiotic Sensitivity")
  
  # Write the data to the worksheet
  writeData(wb, "Antibiotic Sensitivity", data)
  
  # Save the workbook to the specified file
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Call the function to export antibiotic_sensitivity_results to Excel
save_antibiotic_sensitivity_results(antibiotic_sensitivity_results, "Antibiotic_Sensitivity_Results.xlsx")

# Confirmation message
cat("Antibiotic sensitivity results successfully saved to 'Antibiotic_Sensitivity_Results.xlsx'.")

