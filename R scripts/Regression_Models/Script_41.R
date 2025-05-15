setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/michaellutoo")

library(openxlsx)
df <- read.csv("Framing effect- Autism_April 11_ 2024_15.38.csv")

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

df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    x <- as.character(x) # Ensuring the column is treated as character
    x <- gsub("^\\s*$", NA, x) # Replacing only spaces or empty strings with NA
  }
  return(x)
}))

vars_GA <- c("QID41_1"     ,         
               "QID41_2"      ,         "QID41_3"    ,           "QID41_4"   ,            "QID41_5"     ,          "QID41_6"  ,            
               "QID41_7"       ,        "QID41_8"     ,          "QID41_9"    ,           "QID41_10"     ,         "QID41_11" ,            
               "QID41_12"       ,       "QID41_13"     ,         "QID41_14")

vars_ATAB <- c("QID42_1"     ,          "QID42_2"  ,            
                 "QID42_3")

vars_qualitycontact <- c("Quality.of.Contact_1" , "Quality.of.Contact_2" ,
                           "Quality.of.Contact_3"  ,"Quality.of.Contact_4"  ,"Quality.of.Contact_5" , "Quality.of.Contact_6")

var_group <- "FL_14_DO"

vars_cat <- c("QID13","Quantity.of.Contact")

var_policychoice <- "QID35"

vars_attitudesnews <- c("QID37"	,"QID51_1",	"QID51_2"	,"QID51_3",	"QID51_4"	,"QID51_5")


# Reverse scores

reverse_scores_multiple_columns <- function(df, column_names) {
  # Ensure all column names exist in the dataframe
  missing_columns <- setdiff(column_names, names(df))
  if (length(missing_columns) > 0) {
    stop("The following column names do not exist in the dataframe: ", paste(missing_columns, collapse=", "))
  }
  
  # Iterate over each column name, calculate reversed scores, and add as a new column
  for (column_name in column_names) {
    new_column_name <- paste0(column_name, "_rev")
    df[[new_column_name]] <- 6 - df[[column_name]]
  }
  
  return(df)
}

df <- reverse_scores_multiple_columns(df, c("QID41_6",
                                            "QID41_12",
                                            "QID42_1",
                                            "QID42_2",
                                            "QID42_3"))

## RELIABILITY ANALYSIS

# Scale Construction

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SE_of_the_Mean = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Added for item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val <- alpha_results$total$raw_alpha
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"] # Get ITC for the item
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = mean(item_data, na.rm = TRUE),
        SE_of_the_Mean = sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data))),
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
        Alpha = NA
      ))
    }
    
    # Calculate the mean score for the scale and add as a new column
    scale_mean <- rowMeans(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_mean
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = mean(scale_mean, na.rm = TRUE),
      SE_of_the_Mean = sd(scale_mean, na.rm = TRUE) / sqrt(sum(!is.na(scale_mean))),
      StDev = sd(scale_mean, na.rm = TRUE),
      ITC = item_itc,
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}


vars_GA_rev <- c("QID41_1"     ,         
             "QID41_2"      ,         "QID41_3"    ,           "QID41_4"   ,            "QID41_5"     ,          "QID41_6_rev"  ,            
             "QID41_7"       ,        "QID41_8"     ,          "QID41_9"    ,           "QID41_10"     ,         "QID41_11" ,            
             "QID41_12_rev"       ,       "QID41_13"     ,         "QID41_14")

vars_ATAB_rev <- c("QID42_1_rev"     ,          "QID42_2_rev"  ,            
               "QID42_3_rev")


scales <- list(
  "GA" = vars_GA_rev,
  "ATAB" = vars_ATAB_rev,
  "Quality_of_Contact" = vars_qualitycontact,
  "Strenght_of_Frame" = vars_attitudesnews)

alpha_results <- reliability_analysis(df, scales)

df <- alpha_results$data_with_scales
df_reliability <- alpha_results$statistics

## CHI SQUARE

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

row_variables <- var_policychoice
column_variable <- var_group

df_chisquare <- chi_square_analysis_multiple(df, row_variables, column_variable)

# Frequency Table

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


df_freq <- create_frequency_tables(df, c("QID13", var_policychoice, var_group, vars_cat))


# Descriptives by Group

## DESCRIPTIVE TABLES

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

df_descriptive_stats <- calculate_descriptive_stats(df, c("GA"   ,                                        "ATAB"  ,                                      
                                                    "Quality_of_Contact","Strenght_of_Frame", "QID16"))

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

df_descriptives_bygroup <- calculate_descriptive_stats_bygroups(df, c("GA"   ,                                        "ATAB"  ,                                      
                                                                      "Quality_of_Contact", "Strenght_of_Frame"), var_group)


# GLMs


# Outlier and Normality inspection

library(e1071) 

calculate_stats <- function(data, variables) {
  results <- data.frame(Variable = character(),
                        Skewness = numeric(),
                        Kurtosis = numeric(),
                        Shapiro_Wilk_F = numeric(),
                        Shapiro_Wilk_p_value = numeric(),
                        KS_Statistic = numeric(),
                        KS_p_value = numeric(),
                        stringsAsFactors = FALSE)
  
  for (var in variables) {
    if (var %in% names(data)) {
      # Calculate skewness and kurtosis
      skew <- skewness(data[[var]], na.rm = TRUE)
      kurt <- kurtosis(data[[var]], na.rm = TRUE)
      
      # Perform Shapiro-Wilk test
      shapiro_test <- shapiro.test(data[[var]])
      
      # Perform Kolmogorov-Smirnov test
      ks_test <- ks.test(data[[var]], "pnorm", mean = mean(data[[var]], na.rm = TRUE), sd = sd(data[[var]], na.rm = TRUE))
      
      # Add results to the dataframe
      results <- rbind(results, c(var, skew, kurt, shapiro_test$statistic, shapiro_test$p.value, ks_test$statistic, ks_test$p.value))
    } else {
      warning(paste("Variable", var, "not found in the data. Skipping."))
    }
  }
  
  colnames(results) <- c("Variable", "Skewness", "Kurtosis", "Shapiro_Wilk_F", "Shapiro_Wilk_p_value", "KS_Statistic", "KS_p_value")
  return(results)
}

# Example usage with your data
df_normality_results <- calculate_stats(df, c("GA"   ,                                        "ATAB"  ,                                      
                                              "Quality_of_Contact", "Strenght_of_Frame"))

library(dplyr)

calculate_z_scores <- function(data, vars, id_var) {
  # Calculate z-scores for each variable
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~scale(.) %>% as.vector, .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  return(z_score_data)
}

id_var <- "ResponseId"
z_scores_df <- calculate_z_scores(df, c("GA"   ,                                        "ATAB"  ,                                      
                                        "Quality_of_Contact", "Strenght_of_Frame"), id_var)

# Run GLM

library(broom)
library(dplyr)

# Function to run GLM and return APA formatted output
run_glm <- function(formula, data) {
  # Fit the linear model
  model <- lm(formula, data = data)
  
  # Tidy the model to get coefficients and relevant statistics
  tidy_model <- tidy(model) %>%
    select(term, estimate, std.error, statistic, p.value)
  
  # Extract model fit statistics
  model_summary <- summary(model)
  r_squared <- model_summary$r.squared
  adj_r_squared <- model_summary$adj.r.squared
  f_statistic <- model_summary$fstatistic[1]  # The F-statistic value
  p_value_f <- pf(f_statistic, model_summary$df[1], model_summary$df[2], lower.tail = FALSE)  # p-value for the F-statistic
  residual_se <- sigma(model)  # Residual Standard Error
  
  # Create a data frame for model fit statistics
  fit_stats <- data.frame(
    R_Squared = r_squared,
    Adjusted_R_Squared = adj_r_squared,
    F_Statistic = f_statistic,
    P_Value_F = p_value_f,
    Residual_Standard_Error = residual_se
  )
  
  # Return a list containing both the tidy model and fit statistics
  return(list(Coefficients = tidy_model, Fit_Statistics = fit_stats))
}


# Assuming 'data' is your dataframe and 'group' is your categorical variable
df$FL_14_DO <- factor(df$FL_14_DO)  # Convert 'group' to a factor if it's not already
df$FL_14_DO <- relevel(df$FL_14_DO, ref = "GroupC")  # Set 'Control' as the reference category

# Now creating a model
formula_GA <- 'GA ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact' #group - age - gender
formula_GA_mod <- 'GA ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact + FL_14_DO:Quantity.of.Contact + FL_14_DO:Quality_of_Contact'

model_GA <- run_glm(formula_GA, df)
df_modelcoef_GA <- model_GA$Coefficients
df_modelfit_GA <- model_GA$Fit_Statistics

model_GA_mod <- run_glm(formula_GA_mod, df)
df_modelcoef_GA_mod <- model_GA_mod$Coefficients
df_modelfit_GA_mod <- model_GA_mod$Fit_Statistics

formula_ATAB <- 'ATAB ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact' #group - age - gender
formula_ATAB_mod <- 'ATAB ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact + FL_14_DO:Quantity.of.Contact + FL_14_DO:Quality_of_Contact'

model_ATAB <- run_glm(formula_ATAB, df)
df_modelcoef_ATAB <- model_ATAB$Coefficients
df_modelfit_ATAB <- model_ATAB$Fit_Statistics

model_ATAB_mod <- run_glm(formula_ATAB_mod, df)
df_modelcoef_ATAB_mod <- model_ATAB_mod$Coefficients
df_modelfit_ATAB_mod <- model_ATAB_mod$Fit_Statistics


# Now creating a model
formula_GA <- 'GA ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact' #group - age - gender
formula_GA_mod <- 'GA ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact + FL_14_DO:Quantity.of.Contact + FL_14_DO:Quality_of_Contact'

model_GA <- run_glm(formula_GA, df)
df_modelcoef_GA <- model_GA$Coefficients
df_modelfit_GA <- model_GA$Fit_Statistics

model_GA_mod <- run_glm(formula_GA_mod, df)
df_modelcoef_GA_mod <- model_GA_mod$Coefficients
df_modelfit_GA_mod <- model_GA_mod$Fit_Statistics

formula_ATAB <- 'ATAB ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact' #group - age - gender
formula_ATAB_mod <- 'ATAB ~ FL_14_DO + QID16 + QID13 + Quantity.of.Contact + Quality_of_Contact + FL_14_DO:Quantity.of.Contact + FL_14_DO:Quality_of_Contact'

model_ATAB <- run_glm(formula_ATAB, df)
df_modelcoef_ATAB <- model_ATAB$Coefficients
df_modelfit_ATAB <- model_ATAB$Fit_Statistics

model_ATAB_mod <- run_glm(formula_ATAB_mod, df)
df_modelcoef_ATAB_mod <- model_ATAB_mod$Coefficients
df_modelfit_ATAB_mod <- model_ATAB_mod$Fit_Statistics


# Now creating a modstrframeel
formula_GA_modstrframe <- 'GA ~ FL_14_DO + QID16 + QID13 + Strenght_of_Frame + FL_14_DO:Strenght_of_Frame'

modstrframeel_GA_modstrframe <- run_glm(formula_GA_modstrframe, df)
df_modstrframeelcoef_GA_modstrframe <- modstrframeel_GA_modstrframe$Coefficients
df_modstrframeelfit_GA_modstrframe <- modstrframeel_GA_modstrframe$Fit_Statistics

formula_ATAB_modstrframe <- 'ATAB ~ FL_14_DO + QID16 + QID13 + Strenght_of_Frame + FL_14_DO:Strenght_of_Frame'

modstrframeel_ATAB_modstrframe <- run_glm(formula_ATAB_modstrframe, df)
df_modstrframeelcoef_ATAB_modstrframe <- modstrframeel_ATAB_modstrframe$Coefficients
df_modstrframeelfit_ATAB_modstrframe <- modstrframeel_ATAB_modstrframe$Fit_Statistics


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
  "Z-scores" = z_scores_df,
  "Frequency Table" = df_freq, 
  "Chi-Square Results" = df_chisquare, 
  "Descriptive Stats" = df_descriptive_stats,
  "Descriptive by Group" = df_descriptives_bygroup,
  "Reliability" = df_reliability,
  "Model GA" = df_modelcoef_GA,
  "Model GA Int" = df_modelcoef_GA_mod,
  "Model GA Int Frame" = df_modstrframeelcoef_GA_modstrframe,
  "Model ATAB" = df_modelcoef_GA,
  "Model ATAB Int" = df_modelcoef_GA_mod,
  "Model ATAB Int Frame" = df_modstrframeelcoef_ATAB_modstrframe
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")

library(tidyr)
library(ggplot2)

# Function to create dot-and-whisker plots
create_mean_sd_plot <- function(data, variables, factor_column) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    group_by(!!sym(factor_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      .groups = 'drop'
    )
  
  # Create the plot with mean points and error bars for SD
  p <- ggplot(long_data, aes(x = !!sym(factor_column), y = Mean)) +
    geom_point(aes(color = !!sym(factor_column)), size = 3) +
    geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD, color = !!sym(factor_column)), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y") +
    ylim(1, 5) +  # Setting the y-axis limits
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mean and Standard Deviation by Group", 
      x = "",  # Rename the x-axis title
      y = "Value",
      color = "Group"  # Rename the legend title
    ) +
    scale_color_discrete(name = "Group")  # Rename the legend title
  
  # Print the plot
  return(p)
}
plot <- create_mean_sd_plot(df, c("GA", "ATAB", "Quality_of_Contact", "Strenght_of_Frame"), var_group)
print(plot)
