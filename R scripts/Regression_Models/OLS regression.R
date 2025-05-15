library(openxlsx)
library(fastDummies)
library(dplyr)
library(broom)

# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/teamyaadle")

# Read the Excel file
df <- read.xlsx("Updated Clean Data rental_properties_kingston.xlsx")

# Recoding Bedrooms and Bathrooms to "4 or more" if 4 or above
df$Bedrooms <- ifelse(df$Bedrooms >= 4, "4 or more", as.character(df$Bedrooms))
df$Bathrooms <- ifelse(df$Bathrooms >= 4, "4 or more", as.character(df$Bathrooms))

# Define categorical variables with specified reference categories
categorical_vars <- c("Location", "Property_Type", "Bedrooms", "Bathrooms",
                      "Furnished_Unfirnished", "Amenities", "Condition", "Listing_Visibility")

# Create dummy variables except for reference categories
df_recoded <- df %>%
  dummy_cols(select_columns = categorical_vars)

# Outliers

library(dplyr)

calculate_z_scores <- function(data, vars, id_var) {
  # Calculate z-scores for each variable
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~scale(.) %>% as.vector, .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  return(z_score_data)
}

# Example usage of the function
id_var <- c("Price")  # This should be the name of the column you want to use as an identifier
vars_to_scale <- c("Price", "Time.on.Market")  # Replace with actual variable names you want to scale

z_scores_df <- calculate_z_scores(df, vars_to_scale, id_var)


names(df_recoded) <- gsub(" ", "_", trimws(names(df_recoded)))
names(df_recoded) <- gsub("\\s+", "_", trimws(names(df_recoded), whitespace = "[\\h\\v\\s]+"))
names(df_recoded) <- gsub("\\(", "_", names(df_recoded))
names(df_recoded) <- gsub("\\)", "_", names(df_recoded))
names(df_recoded) <- gsub("\\-", "_", names(df_recoded))
names(df_recoded) <- gsub("/", "_", names(df_recoded))
names(df_recoded) <- gsub("\\\\", "_", names(df_recoded)) 
names(df_recoded) <- gsub("\\?", "", names(df_recoded))
names(df_recoded) <- gsub("\\'", "", names(df_recoded))
names(df_recoded) <- gsub("\\,", "_", names(df_recoded))
names(df_recoded) <- gsub("\\$", "", names(df_recoded))
names(df_recoded) <- gsub("\\+", "_", names(df_recoded))

colnames(df_recoded)

# Ensure all predictors and response variables are correctly defined
ivs_model1 <- c("Location_Central"           ,                                             
                "Location_Suburban"           ,                                           
                "Bedrooms_1"                   ,                "Bedrooms_2"        ,                          
                "Bedrooms_3"                    ,               "Bedrooms_4_or_more" ,                         
                "Bathrooms_2"         ,                        
                "Bathrooms_3"                     ,             "Bathrooms_4_or_more"  ,                       
                "Furnished_Unfirnished_Furnished " ,            "Furnished_Unfirnished_Major_Appliances_Only",
                
                "Amenities_Luxury"                   ,          "Amenities_Moderate"    ,                      
                "Condition_Good"                      ,                          
                "Condition_New", "Price")

ivs_model2 <- c("Location_Central"           ,                                             
                "Location_Suburban"           ,                                           
                "Bedrooms_1"                   ,                "Bedrooms_2"        ,                          
                "Bedrooms_3"                    ,               "Bedrooms_4_or_more" ,                         
                "Bathrooms_2"         ,                        
                "Bathrooms_3"                     ,             "Bathrooms_4_or_more"  ,                       
                "Furnished_Unfirnished_Furnished " ,            "Furnished_Unfirnished_Major_Appliances_Only",
                
                "Amenities_Luxury"                   ,          "Amenities_Moderate"    ,                      
                "Condition_Good"                      ,                          
                "Condition_New")

dv_model1 <- "Time.on.Market"
dv_model2 <- "Price"

# OLS Regression

library(broom)
library(stats)

library(broom)
library(stats)

fit_ols_and_format <- function(data, predictors, response_vars, p_value_threshold = 0.20, save_plots = FALSE) {
  ols_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    full_formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Fit the initial OLS multiple regression model
    current_model <- lm(full_formula, data = data)
    
    # Perform manual backward elimination based on p-value threshold
    repeat {
      model_summary <- summary(current_model)
      p_values <- coef(summary(current_model))[, "Pr(>|t|)"]  # Extract p-values of predictors
      max_p_value <- max(p_values[p_values < 1])  # Exclude the intercept p-value if present
      
      if (max_p_value > p_value_threshold) {
        # Find the predictor with the highest p-value and remove it
        worst_predictor <- names(which.max(p_values))
        predictors <- predictors[predictors != worst_predictor]  # Update the predictor set
        # Refit the model without the worst predictor
        current_model <- update(current_model, paste("~ . -", worst_predictor))
      } else {
        break  # Stop if all p-values are below the threshold
      }
    }
    
    # Get the final model summary
    final_model_summary <- summary(current_model)
    
    # Optionally print the final model summary
    print(final_model_summary)
    
    # Extract the tidy output and assign it to ols_results
    ols_results <- broom::tidy(current_model) %>%
      mutate(ResponseVariable = response_var, 
             R_Squared = final_model_summary$r.squared, 
             F_Statistic = final_model_summary$fstatistic[1], 
             P_Value = pf(final_model_summary$fstatistic[1], 
                          final_model_summary$fstatistic[2], 
                          final_model_summary$fstatistic[3], 
                          lower.tail = FALSE))
    
    # Optionally print the tidy output
    print(ols_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(current_model) ~ fitted(current_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(current_model))
      qqline(residuals(current_model))
      dev.off()
    }
    
    # Store the results in a list
    ols_results_list[[response_var]] <- ols_results
  }
  
  return(ols_results_list)
}


ols_results_list_both_groups <- fit_ols_and_format(
  data = df_recoded,
  predictors = ivs_model1,  # Replace with actual predictor names
  response_vars = dv_model1,       # Replace with actual response variable names
  save_plots = TRUE
)

ols_results_list_both_groups2 <- fit_ols_and_format(
  data = df_recoded,
  predictors = ivs_model2,  # Replace with actual predictor names
  response_vars = dv_model2,       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_model1results <- bind_rows(ols_results_list_both_groups)
df_model2results <- bind_rows(ols_results_list_both_groups2)

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
  "Model Time on Market" = df_model1results,
  "Model Price" = df_model2results
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_StepwiseModels.xlsx")
