# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/hy_lau")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Survey Results.xlsx")

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


# Apply listwise deletion based on the specified conditions
df_clean <- df[!(df$SEX_A %in% c(7, 8, 9) | 
                   df$AGEP_A %in% c(97, 98, 99) | 
                   df$RACEALLP_A %in% c(7, 9) | 
                   df$PHSTAT_A %in% c(7, 8, 9) | 
                   df$LSATIS4_A %in% c(7, 8, 9) | 
                   df$EMPWRKFT1_A %in% c(7, 8, 9)), ]

# View the cleaned dataframe
head(df_clean)

# Load dplyr package
library(dplyr)

# Recode RACEALLP_A: All codes other than 1, 2, 3 as 'Other'
df_clean <- df_clean %>%
  mutate(RACEALLP_A = ifelse(RACEALLP_A %in% c(1, 2, 3), RACEALLP_A, "Other"))



# Label variables


# Load dplyr package
library(dplyr)

# Recode SEX_A
df_clean <- df_clean %>%
  mutate(SEX_A = recode(SEX_A, `1` = "Male", `2` = "Female"))

# Recode RACEALLP_A (other values already recoded to "Other")
df_clean <- df_clean %>%
  mutate(RACEALLP_A = recode(RACEALLP_A, `1` = "White only", 
                             `2` = "Black/African American only", 
                             `3` = "Asian only", 
                             .default = "Other"))

# Recode PHSTAT_A (reverse coding 1-5 to 5-1)
df_clean <- df_clean %>%
  mutate(PHSTAT_A = recode(PHSTAT_A, `1` = 5, `2` = 4, `3` = 3, `4` = 2, `5` = 1))

# Recode LSATIS4_A
df_clean <- df_clean %>%
  mutate(LSATIS4_A = recode(LSATIS4_A, `1` = "Very satisfied", 
                            `2` = "Satisfied", 
                            `3` = "Dissatisfied", 
                            `4` = "Very dissatisfied"))

# Recode PA18_05R_A
df_clean <- df_clean %>%
  mutate(PA18_05R_A = recode(PA18_05R_A, `1` = "Meets neither criteria", 
                             `2` = "Meets strength only", 
                             `3` = "Meets aerobic only", 
                             `4` = "Meets both criteria", 
                             `8` = "Not Ascertained"))

# Recode EMPWRKFT1_A
df_clean <- df_clean %>%
  mutate(EMPWRKFT1_A = recode(EMPWRKFT1_A, `1` = "Yes", `2` = "No"))

# View the structure of the dataframe to confirm changes
str(df_clean)

colnames(df_clean)
vars_dummy <- c("SEX_A",  "RACEALLP_A", "PA18_05R_A" )

library(fastDummies)

df_dummy <- dummy_cols(df_clean, select_columns = vars_dummy)

names(df_clean) <- gsub(" ", "_", names(df_clean))
names(df_clean) <- gsub("\\(", "_", names(df_clean))
names(df_clean) <- gsub("\\)", "_", names(df_clean))
names(df_clean) <- gsub("\\-", "_", names(df_clean))
names(df_clean) <- gsub("/", "_", names(df_clean))
names(df_clean) <- gsub("\\\\", "_", names(df_clean)) 
names(df_clean) <- gsub("\\?", "", names(df_clean))
names(df_clean) <- gsub("\\'", "", names(df_clean))
names(df_clean) <- gsub("\\,", "_", names(df_clean))

colnames(df_dummy)

# Multinomial Regression
library(VGAM)

multinomial_regression_vgam <- function(data, dependent_var, independent_vars) {
  
  # Create the formula for the multinomial logistic regression model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit the multinomial logistic regression model using vglm
  model <- vglm(formula, family = multinomial(refLevel = "Very dissatisfied"), data = data, na.action = na.omit)
  
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

names(df_dummy) <- gsub(" ", "_", names(df_dummy))
names(df_dummy) <- gsub("\\(", "_", names(df_dummy))
names(df_dummy) <- gsub("\\)", "_", names(df_dummy))
names(df_dummy) <- gsub("\\-", "_", names(df_dummy))
names(df_dummy) <- gsub("/", "_", names(df_dummy))
names(df_dummy) <- gsub("\\\\", "_", names(df_dummy)) 
names(df_dummy) <- gsub("\\?", "", names(df_dummy))
names(df_dummy) <- gsub("\\'", "", names(df_dummy))
names(df_dummy) <- gsub("\\,", "_", names(df_dummy))

colnames(df_dummy)

dependent_variable <- "LSATIS4_A"

df_dummy <- df_dummy %>%
  mutate(LSATIS4_A = factor(LSATIS4_A, 
                            levels = c("Very dissatisfied", "Dissatisfied", "Satisfied", "Very satisfied"),
                            ordered = TRUE))

independent_variables <- c("PHSTAT_A", "AGEP_A", "SEX_A_Female", "RACEALLP_A_Asian_only", 
                           "RACEALLP_A_Black_African_American_only", "RACEALLP_A_Other", 
                           "PA18_05R_A_Meets_aerobic_only", "PA18_05R_A_Meets_both_criteria", 
                           "PA18_05R_A_Meets_strength_only", "PA18_05R_A_Not_Ascertained")
str(df_dummy)
model_results <- multinomial_regression_vgam(df_dummy, dependent_variable, independent_variables)
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
  "Model Fit" = model_results$FitStatistics,
  "Model Coefs" = model_results$Coefficients
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")


library(car)

# Fit a linear model using the same independent variables
vif_model <- lm(PHSTAT_A ~ AGEP_A + SEX_A_Female + RACEALLP_A_Asian_only + 
                  RACEALLP_A_Black_African_American_only + RACEALLP_A_Other + 
                  PA18_05R_A_Meets_aerobic_only + PA18_05R_A_Meets_both_criteria + 
                  PA18_05R_A_Meets_strength_only + PA18_05R_A_Not_Ascertained, data = df_dummy)

# Calculate VIF for the independent variables
vif_values <- vif(vif_model)

# Display the VIF values
print(vif_values)

# Example usage
data_list <- list(
  "DummyDf" = df_dummy
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "DummyDF.xlsx")
