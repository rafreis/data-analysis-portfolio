setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/ocin1234")

library(haven)
df <- read_sav("Data_Recoded.sav")

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


library(car)
library(multcomp)
library(nparcomp)


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
df_normality_results <- calculate_stats(df, "COMP1_Scale")

print(df_normality_results)


library(dplyr)

# Define the variables where 0 should be replaced by NA
vars_to_replace <- c("v_35", "v_37", "v_124", "v_132", "v_233", "v_243", "v_188", "v_189", 
                     "v_167", "v_146", "v_145", "v_98", "v_191", "v_237", "v_239", 
                     "v_19", "v_20", "v_21", "v_22")

# Remove duplicates to ensure efficiency
unique_vars <- unique(vars_to_replace)

# Replace 0 with NA for the specified variables
df <- df %>%
  mutate(across(all_of(unique_vars), ~na_if(.x, 0)))

# ONE WAY ANOVA


# Updated function to perform one-way ANOVA for a single DV against multiple IVs
perform_one_way_anova <- function(data, dv, ivs) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  
  # Iterate over independent variables (factors)
  for (factor in ivs) {
    # One-way ANOVA
    anova_model <- aov(reformulate(factor, response = dv), data = data)
    model_summary <- summary(anova_model)
    
    # Extract relevant statistics for the factor
    if (!is.null(model_summary[[1]])) {
      model_results <- data.frame(
        DV = dv,
        IV = factor,
        Sum_Sq = model_summary[[1]]$"Sum Sq"[1],  # Extract Sum Sq for the factor
        Mean_Sq = model_summary[[1]]$"Mean Sq"[1],  # Extract Mean Sq for the factor
        Df = model_summary[[1]]$Df[1],  # Extract Df for the factor
        FValue = model_summary[[1]]$"F value"[1],  # Extract F value for the factor
        pValue = model_summary[[1]]$"Pr(>F)"[1]  # Extract p-value for the factor
      )
      
      anova_results <- rbind(anova_results, model_results)
    }
  }
  
  return(anova_results)
}

# List of independent variables (factors) to test against the DV "COMP1_Scale"
ivs <- c("v_193", "v_194", "v_195", "v_28", "v_29", "v_30", "v_31", "v_33", 
         "v_35", "v_37", "v_124", "v_132", "v_233", "v_243", "v_188", "v_189", 
         "v_167", "v_146", "v_145", "v_98", "v_191", "v_237", "v_239", "v_19", 
         "v_20", "v_21", "v_22")


# Perform one-way ANOVA for "COMP1_Scale" against each independent variable
one_way_anova_results <- perform_one_way_anova(df, "COMP1_Scale", ivs)

# View the results
print(one_way_anova_results)

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
  "Normality results" = df_normality_results,
  "ANOVA Results" = one_way_anova_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

