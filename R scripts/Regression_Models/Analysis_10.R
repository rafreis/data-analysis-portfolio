setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/oliversogard")

library(openxlsx)
df <- read.xlsx("Fiver ASES Study Data (1).xlsx")

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

colnames(df)

chisq_vars <- c("Gender", "anatomic.or.reverse.total.shoulder.arthroplasty", "Patient.Arm"    ,  "Diabetes"   ,                                    
                "Tobacco.Use"                      ,               "Previous.Cardiovascular.Event"   ,               
                "Periprosthetic.Joint.infection.within.6.weeks"  , "Postoperative.blood.transfusion"   ,             
                "Periprosthetic.fracture.within.6.weeks"      ,    "Surgery.Revision.within.6.weeks"  ,              
                "Readmission.within.30.days")

anova_vars <- c("Age", "BMI", "Operation.Time", "ASES.Patient.Satisfaction.Score.Change")

iv <- ("Resident.Present.During.Surgery")


#DESCRIPTIVE STATISTICS

library(dplyr)

calculate_descriptive_stats_bygroups <- function(data, desc_vars, cat_column) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each category
  for (var in desc_vars) {
    # Group data by the categorical column
    grouped_data <- data %>%
      group_by(!!sym(cat_column)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE)
        
      ) %>%
      mutate(Variable = var)
    
    # Append the results for each variable and category to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}

desc_table_disag <- calculate_descriptive_stats_bygroups(df, anova_vars, iv)
desc_table_2 <- calculate_descriptive_stats_bygroups(df, anova_vars, "anatomic.or.reverse.total.shoulder.arthroplasty")

calculate_descriptive_stats_bytwogroups <- function(data, desc_vars, cat_column1, cat_column2) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category1 = character(),
    Category2 = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each combination of categories
  for (var in desc_vars) {
    # Group data by both categorical columns
    grouped_data <- data %>%
      group_by(!!sym(cat_column1), !!sym(cat_column2)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE),
        .groups = 'drop' # Drop grouping structure after summarising
      ) %>%
      mutate(Variable = var) %>%
      rename(Category1 = !!sym(cat_column1), Category2 = !!sym(cat_column2))
    
    # Append the results for each variable and category combination to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}

desc_table_disagtwo <- calculate_descriptive_stats_bytwogroups(df, anova_vars, iv, "anatomic.or.reverse.total.shoulder.arthroplasty")


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

freq_df <- create_frequency_tables(df, chisq_vars)


library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

desc_table <- calculate_descriptive_stats(df, anova_vars)

# CHI_SQ ANALYSIS

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

chisq_result_df <- chi_square_analysis_multiple(df, chisq_vars, iv)
print(chisq_result_df)


# Function to perform one-way ANOVA
perform_one_way_anova <- function(data, response_vars, factor) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  
  # Iterate over response variables
  for (var in response_vars) {
    # One-way ANOVA
    anova_model <- aov(reformulate(factor, response = var), data = data)
    model_summary <- summary(anova_model)
    
    # Extract relevant statistics
    model_results <- data.frame(
      Variable = var,
      Effect = factor,
      Sum_Sq = model_summary[[1]]$"Sum Sq"[1],  # Extract Sum Sq for the factor
      Mean_Sq = model_summary[[1]]$"Mean Sq"[1],  # Extract Mean Sq for the factor
      Df = model_summary[[1]]$Df[1],  # Extract Df for the factor
      FValue = model_summary[[1]]$"F value"[1],  # Extract F value for the factor
      pValue = model_summary[[1]]$"Pr(>F)"[1]  # Extract p-value for the factor
    )
    
    anova_results <- rbind(anova_results, model_results)
  }
  
  return(anova_results)
}


# Perform one-way ANOVA
one_way_anova_results <- perform_one_way_anova(df, anova_vars, iv)

# ANALYSIS DISAGGREGATED BY SURGERY

chi_square_analysis_subset <- function(data, row_vars, col_var, subset_col) {
  results <- list() # Initialize an empty list to store results
  
  subset_levels <- unique(data[[subset_col]])
  
  for (subset_level in subset_levels) {
    subset_data <- data[data[[subset_col]] == subset_level, ]
    
    for (row_var in row_vars) {
      subset_data[[row_var]] <- factor(subset_data[[row_var]])
      subset_data[[col_var]] <- factor(subset_data[[col_var]])
      
      crosstab <- prop.table(table(subset_data[[row_var]], subset_data[[col_var]]), margin = 2) * 100
      is_2x2_table <- all(dim(table(subset_data[[row_var]], subset_data[[col_var]])) == 2)
      chi_square_test <- chisq.test(table(subset_data[[row_var]], subset_data[[col_var]]), correct = is_2x2_table)
      crosstab_df <- as.data.frame.matrix(crosstab)
      
      for (level in levels(subset_data[[row_var]])) {
        level_df <- data.frame(
          Subset_Level = subset_level,
          Row_Variable = row_var,
          Row_Level = level,
          Column_Variable = col_var,
          check.names = FALSE
        )
        
        level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
        level_df$Chi_Square <- chi_square_test$statistic
        level_df$P_Value <- chi_square_test$p.value
        
        results[[paste0(subset_level, "_", row_var, "_", level)]] <- level_df
      }
    }
  }
  
  return(do.call(rbind, results))
}


chisq_result_disag_df <- chi_square_analysis_subset(df, chisq_vars, iv, "anatomic.or.reverse.total.shoulder.arthroplasty" )
print(chisq_result_disag_df)


perform_one_way_anova_subset <- function(data, response_vars, factor, subset_col) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  
  subset_levels <- unique(data[[subset_col]])
  
  for (subset_level in subset_levels) {
    subset_data <- data[data[[subset_col]] == subset_level, ]
    
    for (var in response_vars) {
      anova_model <- aov(reformulate(factor, response = var), data = subset_data)
      model_summary <- summary(anova_model)
      
      model_results <- data.frame(
        Subset_Level = subset_level,
        Variable = var,
        Effect = factor,
        Sum_Sq = model_summary[[1]]$"Sum Sq"[1],
        Mean_Sq = model_summary[[1]]$"Mean Sq"[1],
        Df = model_summary[[1]]$Df[1],
        FValue = model_summary[[1]]$"F value"[1],
        pValue = model_summary[[1]]$"Pr(>F)"[1]
      )
      
      anova_results <- rbind(anova_results, model_results)
    }
  }
  
  return(anova_results)
}

one_way_anova_disag_results <- perform_one_way_anova_subset(df, anova_vars, iv,"anatomic.or.reverse.total.shoulder.arthroplasty")

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
  "Descriptives" = desc_table,
  "Descriptives - Disag" = desc_table_disag,
  "Descriptives - Disag 2" = desc_table_disagtwo,
  "Chi-Square Results" = chisq_result_df, 
  "Chi-Square Results Disag" = chisq_result_disag_df, 
  "ANOVA Results" = one_way_anova_results,
  "ANOVA Results Disag" = one_way_anova_disag_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")
