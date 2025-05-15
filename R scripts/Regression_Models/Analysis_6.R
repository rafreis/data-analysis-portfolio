setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/intkabmah")

library(openxlsx)
df <- read.xlsx("Data.xlsx")

# Get rid of special characters

names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))
names(df) <- gsub(" ", "_", names(df))

library(dplyr)

df <- df %>%
  mutate(Interval.between.current.pregnancy.and.last.pregnancy = case_when(
    Interval.between.current.pregnancy.and.last.pregnancy == "&gt;5 years" ~ "more than 5 years",
    Interval.between.current.pregnancy.and.last.pregnancy == "&lt; 1 year" ~ "less than 1 year",
    TRUE ~ Interval.between.current.pregnancy.and.last.pregnancy
  ))

# Define Variables of interest

vars <- c("Group"                  ,                               "Maternal.Age"    ,                                     
          "Gravidity"                      ,                       "Parity"          ,                                     
          "Abortions"                       ,                      "Interval.between.current.pregnancy.and.last.pregnancy",
          "Indication.of.last.LSCS"         ,                      "Used.IUCD.after.last.birth"                  ,         
          "H_O.PID"                         ,                      "H_O.UTI.since.last.birth"                     ,        
          "current.pregnancy.by.ART"         ,                     "Treatment.options.for.scar.site.pregnancy"     ,       
          "Clinical.presentation"             ,                    "Level.of.BHCG.AT.DIAGNOSIS"                    ,      
          "lEVEL.OF.BHCG.1.week.after.tm"     )

# Frequency Tables

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


df_freq <- create_frequency_tables(df, vars)

create_segmented_frequency_tables <- function(data, vars, segmenting_categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each segmenting category
  for (segment in segmenting_categories) {
    # Iterate over each variable for which frequencies are to be calculated
    for (var in vars) {
      # Ensure both the segmenting category and variable are treated as factors
      segment_factor <- factor(data[[segment]])
      var_factor <- factor(data[[var]])
      
      # Calculate counts segmented by the segmenting category
      counts <- table(segment_factor, var_factor)
      
      # Melt the table into a long format for easier handling
      freq_table <- as.data.frame(counts)
      names(freq_table) <- c("Segment", "Level", "Count")
      
      # Add the variable name and calculate percentages
      freq_table$Variable <- var
      freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

vars2 <- c("Maternal.Age"    ,                                     
          "Gravidity"                      ,                       "Parity"          ,                                     
          "Abortions"                       ,                      "Interval.between.current.pregnancy.and.last.pregnancy",
          "Indication.of.last.LSCS"         ,                      "Used.IUCD.after.last.birth"                  ,         
          "H_O.PID"                         ,                      "H_O.UTI.since.last.birth"                     ,        
          "current.pregnancy.by.ART"         ,                     "Treatment.options.for.scar.site.pregnancy"     ,       
          "Clinical.presentation"             ,                    "Level.of.BHCG.AT.DIAGNOSIS"                    ,      
          "lEVEL.OF.BHCG.1.week.after.tm"     )

df_segmented_freq <- create_segmented_frequency_tables(df, vars2, "Group")


# Chi-Square Comparison

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

row_variables <- vars2
column_variable <- "Group"  # Replace with your column variable name

df_chisquare_result <- chi_square_analysis_multiple(df, row_variables, column_variable)


# Risk Factors Modelling

# Group levels that are not representative

library(dplyr)

df_recoded <- df %>%
  mutate(Gravidity = case_when(
    Gravidity == "3 and Below" ~ "5 and Below",
    Gravidity == "4-5" ~ "5 and Below",
    TRUE ~ Gravidity
  )) %>%
  mutate(Parity = case_when(
    Parity == "4" ~ "4-5",
    TRUE ~ Parity
  )) %>%
  mutate(Abortions = case_when(
    Abortions == "1-2" ~ "Yes",
    Abortions == "3 and above" ~ "Yes",
    TRUE ~ Abortions
  )) %>%
  mutate(Interval.between.current.pregnancy.and.last.pregnancy = case_when(
    Interval.between.current.pregnancy.and.last.pregnancy == "more than 5 years" ~ "2 or more years",
    Interval.between.current.pregnancy.and.last.pregnancy == "2-5 years" ~ "2 or more years",
    TRUE ~ Interval.between.current.pregnancy.and.last.pregnancy
  )) %>%
  mutate(Indication.of.last.LSCS = case_when(
    Indication.of.last.LSCS == "anyother" ~ "Any Other",
    Indication.of.last.LSCS == "Fetal distress" ~ "Any Other",
    TRUE ~ Indication.of.last.LSCS
  )) %>%
  mutate(Clinical.presentation = case_when(
    Clinical.presentation == "referred after failed tm from hospital" ~ "Assymptomatic or referred after failed tm",
    Clinical.presentation == "asymptomatic" ~ "Assymptomatic or referred after failed tm",
    TRUE ~ Clinical.presentation
  ))


## PROBIT REGRESSION

#Change reference category

# df_recode$Nivel_máximo_de_estudios <- relevel(df_recode$Nivel_máximo_de_estudios, ref = "Primary")

library(DescTools)


probit_regression <- function(data, dependent_var, independent_vars) {
  # Create the formula for the probit model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit the probit model
  model <- glm(formula, data = data, family = binomial(link = "probit"))
  
  # Print the model summary (includes Wald test statistics and Chi-squared statistic)
  print(summary(model))
  
  # Ensure the dependent variable is a numeric binary variable
  data[[dependent_var]] <- as.numeric(data[[dependent_var]])
  
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
  
  if(is.numeric(model$deviance) && is.numeric(model$null.deviance)) {
    r_squared <- 1 - model$deviance / model$null.deviance
    cat("McFadden's Pseudo R-squared: ", r_squared, "\n")
  } else {
    cat("Cannot calculate pseudo R-squared: deviance values are not numeric.\n")
  }
  
  
  return(results_df)
}

#df_recoded$Group <- ifelse(df_recoded$Group == "Group1(scar site pregnancy)", 1, 0) # Define binary variable to use as DV
dependent_variable <- "Group" 

vars3 <- c("Used.IUCD.after.last.birth"                  ,         
                    "H_O.PID"                         ,                      "H_O.UTI.since.last.birth"                     ,        
                    "current.pregnancy.by.ART"         ,                   "Interval.between.current.pregnancy.and.last.pregnancy",      
                    "Clinical.presentation" )

df_recoded <- df_recoded %>%
  mutate_at(vars(one_of(vars3)), factor)

independent_variables <- vars3

model_results <- probit_regression(df_recoded, dependent_variable, independent_variables)

vars4 <- c("Maternal.Age", "Used.IUCD.after.last.birth"                  ,         
           "H_O.PID"                         ,                      "H_O.UTI.since.last.birth"                     ,        
           "current.pregnancy.by.ART"         ,                   "Interval.between.current.pregnancy.and.last.pregnancy",      
           "Clinical.presentation" )

model_results_age <- probit_regression(df_recoded, dependent_variable, vars4)

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
  "Cleaned Data" = df_recoded,
  "Frequency Table" = df_freq, 
  "Disaggregated Freq Table" = df_segmented_freq,
  "Chi-Square Results" = df_chisquare_result, 
  "Probit Model results" = model_results,
  "Probit Model results - Age" = model_results_age
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")

