setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dimplepatel")

library(openxlsx)
library(tidyr)
library(dplyr)
library(VGAM)

df <- read.xlsx("AllSTRIPatientsBiopsied_v.xlsx")

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

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))


# Create a new column 'Obesity' where BMI > 30 is marked as 1 (Obese) and others as 0 (Not Obese)
df$Obesity <- ifelse(df$BMI > 30, 1, 0)

vars_ivsmodel <- c("FIB_4", "ELF", "FAST", "Fibroscan.CAP" ,"Fibroscan.E._kPA_")

vars_control <- c("Type.II.Diabetic"      ,         "Hypertension"    ,              
                  "High.Cholesterol", "Obesity")

var_dv <- "Liver.Biopsy.Fibrosis.Stage"

# Function to diagnose missing data
diagnose_missing_data <- function(df, vars) {
  missing_summary <- df %>%
    select(all_of(vars)) %>%
    summarise(across(everything(), ~ sum(is.na(.)))) %>%
    pivot_longer(cols = everything(), names_to = "Variable", values_to = "Missing_Values")
  
  total_rows <- nrow(df)
  
  missing_summary <- missing_summary %>%
    mutate(Percentage_Missing = (Missing_Values / total_rows) * 100)
  
  return(missing_summary)
}

missing_data_summary <- diagnose_missing_data(df, c(vars_ivsmodel, vars_control))

# DESCRIPTIVES

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

df_freq <- create_frequency_tables(df, c(vars_control, var_dv, "Gender"))

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


df_freq_segmented <- create_segmented_frequency_tables(df, vars_control, var_dv)

library(moments)

# Function to calculate descriptive statistics with the count of non-missing observations
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    N = integer(),      # Number of non-missing observations
    Mean = numeric(),
    SEM = numeric(),    # Standard Error of the Mean
    SD = numeric(),     # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    n <- length(na.omit(variable_data))  # Calculate number of non-missing observations
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(n)
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      N = n,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}


df_descriptive_stats <- calculate_descriptive_stats(df, vars_ivsmodel)


# ANOVAS

# Function to perform one-way ANOVA
perform_one_way_anova <- function(data, response_vars, factors) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  means_results <- data.frame()  # Initialize an empty dataframe to store means for each factor level
  
  # Iterate over response variables
  for (var in response_vars) {
    # Iterate over factors
    for (factor in factors) {
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
      
      # Calculate means for each level of the factor
      factor_means <- aggregate(data[[var]], list(data[[factor]]), mean, na.rm = TRUE)
      names(factor_means) <- c(factor, "Mean")
      factor_means$Variable <- var
      factor_means$Effect <- factor
      
      # Combine results into the main dataframe
      anova_results <- rbind(anova_results, model_results)
      means_results <- rbind(means_results, factor_means)
    }
  }
  
  # Combine anova_results and means_results into a single list for output
  return(list(ANOVA_Results = anova_results, Means = means_results))
}


# Example usage
# Define your response variables (dependent variables) and the factor (independent variable)
response_vars <- vars_ivsmodel  # Add more variables as needed
factor <- var_dv # Specify the factor (independent variable)

# Perform one-way ANOVA
df_one_way_anova_results <- perform_one_way_anova(df, response_vars, factor)
df_one_way_anova_results <- df_one_way_anova_results$ANOVA_Results
df_means_factors <- df_one_way_anova_results$Means

# REGRESSION MODELS

library(dplyr)

calculate_z_scores <- function(data, vars, id_var) {
  # Calculate z-scores for each variable
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~scale(.) %>% as.vector, .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  return(z_score_data)
}

id_var <- "Patient.ID"
z_scores_df <- calculate_z_scores(df, vars_ivsmodel, id_var)


# Replace 'yes' with 1 and 'no' with 0 in specified columns
columns_to_replace <- c("Hypertension", "High.Cholesterol")
df[columns_to_replace] <- lapply(df[columns_to_replace], function(x) ifelse(x == "Yes", 1, 0))

columns_to_replace <- "Site"
df[columns_to_replace] <- lapply(df[columns_to_replace], function(x) ifelse(x == "Edinburg", 1, 0))

# For the "Type.II.Diabetic" column, code 'Yes' and 'Pre-Diabetic' as 1, and 'No' as 0
df$Type.II.Diabetic <- ifelse(df$Type.II.Diabetic %in% c("Yes", "Pre-Diabetic"), 1, ifelse(df$Type.II.Diabetic == "No", 0, NA))

vars_control <- c("Type.II.Diabetic",     "Hypertension"    ,              
                  "High.Cholesterol", "Obesity")


# Ensure your dependent variable is a factor
df$Liver.Biopsy.Fibrosis.Stage <- factor(df$Liver.Biopsy.Fibrosis.Stage)

df$Liver.Biopsy.Fibrosis.Stage <- factor(df$Liver.Biopsy.Fibrosis.Stage, levels = c("Stage 0", "Stage 1", "Stage 2", "Stage 3", "Stage 4"))

# Check levels of the dependent variable
levels(df$Liver.Biopsy.Fibrosis.Stage)


# Optionally, set a specific level as the reference
# This example sets 'Stage 0' as the reference level
df$Liver.Biopsy.Fibrosis.Stage <- relevel(df$Liver.Biopsy.Fibrosis.Stage, ref = "Stage 0")


multinomial_regression_vgam <- function(data, dependent_var, independent_vars) {
  
  # Create the formula for the multinomial logistic regression model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit the multinomial logistic regression model using vglm
  model <- vglm(formula, family = multinomial(refLevel = "Stage 0"), data = data, na.action = na.omit)
  
  # Print the summary of the model
  summary_model <- summary(model, coef = TRUE)
  print(summary_model)
  
  # Extract coefficients and compute standard errors, z-values, and p-values
  coefficients <- coef(summary_model)
  
  # Create a dataframe to store coefficient results
  coef_df <- as.data.frame(coefficients)
  
  # Calculate Odds Ratios from the log odds (coefficients)
  coef_df$OddsRatios <- exp(coef_df$Estimate)
  
  # Extract model fit statistics
  deviance <- slot(model, "criterion")$deviance
  log_likelihood <- slot(model, "criterion")$loglikelihood
  residual_deviance <- slot(model, "ResSS")
  df_residual <- slot(model, "df.residual")
  df_total <- slot(model, "df.total")
  iter <- slot(model, "iter")
  rank <- slot(model, "rank")
  
  # Create a dataframe to store model fit statistics
  fit_stats_df <- data.frame(
    Deviance = deviance,
    LogLikelihood = log_likelihood,
    ResidualDeviance = residual_deviance,
    DF_Residual = df_residual,
    DF_Total = df_total,
    Iterations = iter,
    Rank = rank
  )
  
  # Returning a list of two dataframes
  return(list(Coefficients = coef_df, FitStatistics = fit_stats_df))
}



# Example usage with your data
dependent_var <- var_dv

independent_vars_1 <- c("FIB_4", "Site")
independent_vars_2 <- c("ELF", "Site")
independent_vars_3 <- c("FAST", "Site")
independent_vars_4 <- c("Fibroscan.CAP" ,"Fibroscan.E._kPA_", "Site")
independent_vars_5 <- c(vars_control, "Site")

model_results1 <- multinomial_regression_vgam(df, dependent_var, independent_vars_1)
model_results2 <- multinomial_regression_vgam(df, dependent_var, independent_vars_2)
model_results3 <- multinomial_regression_vgam(df, dependent_var, independent_vars_3)

df_fibrooutlier <- df[df$Patient.ID != "0173-0115", ]
model_results4 <- multinomial_regression_vgam(df_fibrooutlier, dependent_var, independent_vars_4)
model_results5 <- multinomial_regression_vgam(df, dependent_var, independent_vars_5)

df_modelcoef1 <- model_results1$Coefficients
df_modelfit1 <- model_results1$FitStatistics

df_modelcoef2 <- model_results2$Coefficients
df_modelfit2 <- model_results2$FitStatistics

df_modelcoef3 <- model_results3$Coefficients
df_modelfit3 <- model_results3$FitStatistics

df_modelcoef4 <- model_results4$Coefficients
df_modelfit4 <- model_results4$FitStatistics

df_modelcoef5 <- model_results5$Coefficients
df_modelfit5 <- model_results5$FitStatistics


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
  "Data" = df,
  "Descriptive Statistics" = df_descriptive_stats,
  "Frequency Table" = df_freq,
  "Frequency Table - per Stage" = df_freq_segmented,
  "Descriptives by Stage" <- df_means_factors,
  "ANOVA" = df_one_way_anova_results,
  "Model - FIB" = df_modelcoef1,
  "Model - ELF" = df_modelcoef2,
  "Model - FAST" = df_modelcoef3,
  "Model - Fibroscan" = df_modelcoef4,
  "Model - Comorbidities" = df_modelcoef5,
  "ModelFit - FIB" = df_modelfit1,
  "ModelFit - ELF" = df_modelfit2,
  "ModelFit - FAST" = df_modelfit3,
  "ModelFit - Fibroscan" = df_modelfit4,
  "ModelFit - Comorbidities" = df_modelfit5
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")


## BOXPLOTS DIVIDED BY FACTOR

library(ggplot2)

# Specify the factor as a column name
factor_var <- var_dv  # replace with your actual factor column name
vars <- vars_ivsmodel

# Loop for generating and saving boxplots
for (var in vars) {
  var_name_modified <- gsub("\\.", " ", var)
  var_name_modified <- gsub("_", "", var_name_modified)
  factor_var_modified <- gsub("\\.", " ", factor_var)
  
  # Create boxplot for each variable using ggplot2
  p <- ggplot(df_fibrooutlier, aes_string(x = factor_var, y = var)) +
    geom_boxplot() +
    stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
    stat_summary(fun = mean, geom = "text", aes(label = round(..y.., 2)), vjust = -1) +
    labs(title = paste(var_name_modified, "by", factor_var_modified),
         x = factor_var_modified, y = var_name_modified)
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = paste("boxplot_", gsub("\\.", " ", var), ".png", sep = ""), plot = p, width = 10, height = 6)
}

# New model


library(dplyr)

# Assuming `data` is your data frame and `BMI` is the column to categorize
df <- df %>%
  mutate(
    BMI_category = case_when(
      BMI < 18.5 ~ "Underweight",
      BMI >= 18.5 & BMI < 25 ~ "Healthy",
      BMI >= 25 & BMI < 30 ~ "Overweight",
      BMI >= 30 ~ "Obesity",
      TRUE ~ "Unknown" # Optional: For any unexpected cases
    )
  )

df_freq_obesity <- create_frequency_tables(df,"BMI_category")

df <- df %>%
  mutate(
    BMI_categoryext = case_when(
      BMI < 18.5 ~ "Underweight",
      BMI >= 18.5 & BMI < 25 ~ "Healthy",
      BMI >= 25 & BMI < 30 ~ "Overweight",
      BMI >= 30 & BMI < 35 ~ "Class 1 Obesity",
      BMI >= 35 & BMI < 40 ~ "Class 2 Obesity",
      BMI >= 40 ~ "Class 3 Obesity",
      TRUE ~ "Unknown" # Optional: For any unexpected cases
    )
  )

df_freq_obesity <- create_frequency_tables(df,"BMI_categoryext")

#data for new analysis

df_filt<- df %>%
  filter(!BMI_categoryext %in% c("Healthy", "Unknown"))

model_results8 <- multinomial_regression_vgam(df_filt, dependent_var, "BMI_categoryext")

df_modelcoef8 <- model_results8$Coefficients
df_modelfit8 <- model_results8$FitStatistics

# Example usage
data_list <- list(
  
  "Frequency Table - Obesity" = df_freq_obesity,
  "Model Coef" = df_modelcoef8,
  "Model Fit" = df_modelfit8
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_Obesity.xlsx")
