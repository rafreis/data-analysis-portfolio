setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/avanr1975")

library(readxl)
df <- read_xlsx('FinalData.xlsx')

# Get rid of spaces in column names

names(df) <- gsub(" ", "_", names(df))

# List of variables to be dummyfied
variables_to_dummyfy <- c("Age_Groups", "Diabetes", "Hypertension", "Asthma", "hISTORY_OF_PREVIOUS_SURGERY", "PSA_LEVELS", "GLEASON")

# Convert to factors
vars = c("Age_Groups", "Diabetes","Hypertension","Asthma","hISTORY_OF_PREVIOUS_SURGERY","PSA_LEVELS","GLEASON")

# Loop to convert each variable in 'vars' to a factor, handling missing values
for (var in vars) {
  # Check if the variable exists in the dataframe
  if (var %in% names(df)) {
    # Convert the variable to a factor, keeping NA as a level if they exist
    df[[var]] <- factor(df[[var]], exclude = NULL)
  } else {
    # Print a message if the variable is not found in the dataframe
    message(sprintf("Variable '%s' not found in the dataframe.", var))
  }
}

# Loop to create dummy variables
for (var in variables_to_dummyfy) {
  # Check if the variable exists in the dataframe
  if (var %in% names(df)) {
    # Convert to factor, keeping NA as a level
    df[[var]] <- factor(df[[var]], exclude = NULL)
    
    # Specify the contrasts to create a dummy variable for each level
    contrast_list <- list()
    contrast_list[[var]] <- contrasts(df[[var]], contrasts = FALSE)
    
    # Creating dummy variables including all levels
    dummy_vars <- model.matrix(reformulate(var), data = df, contrasts.arg = contrast_list)
    
    # Removing the intercept column (first column) from dummy_vars
    dummy_vars <- dummy_vars[, -1]
    # Check if dummy_vars is not a matrix (i.e., has only one column)
    if (!is.matrix(dummy_vars)) {
      # Convert it to a matrix
      dummy_vars <- matrix(dummy_vars, ncol = 1)
      colname <- paste0(var, "_", levels(df[[var]])[2])  # Assuming the first level is excluded
      colnames(dummy_vars) <- colname
    } else {
      # Renaming dummy variables to include the original variable name for clarity
      colnames(dummy_vars) <- gsub(paste0("^", var), paste0(var, "_"), colnames(dummy_vars))
    }
    
    # Adding dummy variables to the original dataframe
    df <- cbind(df, dummy_vars)
  } else {
    # Print a message if the variable is not found in the dataframe
    message(sprintf("Variable '%s' not found in the dataframe.", var))
  }
}

## FREQUENCY TABLE

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

freq_df <- create_frequency_tables(df, vars)

## CHI-SQUARE ANALYSIS

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


row_variables <- c("Age_Groups", "Diabetes", "Hypertension","Asthma","hISTORY_OF_PREVIOUS_SURGERY","PSA_LEVELS","GLEASON")
column_variable <- "Surgery"  

chisquare_result_df <- chi_square_analysis_multiple(df, row_variables, column_variable)
print(result_df)

## PROBIT REGRESSION

probit_regression <- function(data, dependent_var, independent_vars) {
  # Create the formula for the probit model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit the probit model
  model <- glm(formula, data = data, family = binomial(link = "probit"))
  
  # Print the model summary (includes Wald test statistics and Chi-squared statistic)
  print(summary(model))
  
  # Calculating pseudo R-squared values
  r_squared <- with(model, 1 - deviance/null.deviance)
  
  # Check the significance of the model
  model_significance <- summary(model)$coef[2, 4]
  
  # Calculating the change in deviance and its significance
  change_in_deviance <- with(model, null.deviance - deviance)
  change_in_df <- with(model, df.null - df.residual)
  chi_square_p_value <- with(model, pchisq(null.deviance - deviance, df.null - df.residual, lower.tail = FALSE))
  
  cat("\nChange in Deviance: ", change_in_deviance, "\n")
  cat("Degrees of Freedom Difference: ", change_in_df, "\n")
  cat("Chi-Square Test P-Value for Model Significance: ", chi_square_p_value, "\n")
  
  # Extracting coefficients, Odds Ratios, and p-values
  coefficients <- coef(model)
  odds_ratios <- exp(coefficients)
  p_values <- summary(model)$coef[, 4]
  
  # Create a dataframe to store results
  results_df <- data.frame(
    Coefficients = coefficients,
    Odds_Ratios = odds_ratios,
    P_Values = p_values,
    row.names = names(coefficients)
  )
  
  # Print R-squared and overall model summary
  cat("McFadden's Pseudo R-squared: ", r_squared, "\n")
  
  
  return(results_df)
}

names(df) <- gsub(">=", "olderthan", names(df))
names(df) <- gsub("-", "_", names(df))

#Recode DV
df$Surgery <- as.numeric(df$Surgery == "Yes") 

dependent_variable <- "Surgery"
independent_variables <- c("Age_Groups_65_74","Age_Groups_olderthan75" ,"Diabetes_yes" ,"Hypertension_yes" ,"Asthma_yes"  ,"hISTORY_OF_PREVIOUS_SURGERY_yes",
                           "PSA_LEVELS_high", "PSA_LEVELS_medium","PSA_LEVELS_NA","GLEASON_high" ,"GLEASON_medium")

model_results <- probit_regression(df, dependent_variable, independent_variables)
print(model_results)

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
  "Cleaned Data" = df,
  "Frequency Table" = freq_df, 
  "Chi-Square Results" = chisquare_result_df, 
  "Probit Model results" = model_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")


