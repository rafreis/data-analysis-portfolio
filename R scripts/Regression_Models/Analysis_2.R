setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/beasalazar/Second order")

# New Models

ivs <- c("Relación_con_la_Agricultura_Orgánica",
                          "Sexo", "Edad_Cat", "Nivel_máximo_de_estudios", "Ingresos_mensuales_netos")

dvs_list <- c("Monto_de_Impuestos", "Polinización"      ,                
              "Hábitat_para_especies"            ,    "Descomposición_de_la_MO"           ,  
              "Educación_ambiental"            ,      "Regulación_del_clima"                ,
              "Turismo_de_naturaleza_y_rural"   ,     "Conocimiento_científico"             ,
              "Dispersión_de_semillas"           ,    "Identidad_local"                     ,
              "Control_de_plagas"                ,    "Agua_para_riego_y_consumo"           ,
              "Provisión_de_alimentos"            ,   "Conocimiento_local_tradicional")

# OLS Regression

library(broom)

fit_ols_and_format <- function(data, predictors, response_vars, save_plots = FALSE) {
  ols_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Fit the OLS multiple regression model
    lm_model <- lm(formula, data = data)
    
    # Get the summary of the model
    model_summary <- summary(lm_model)
    
    # Extract the R-squared, F-statistic, and p-value
    r_squared <- model_summary$r.squared
    f_statistic <- model_summary$fstatistic[1]
    p_value <- pf(f_statistic, model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE)
    
    # Print the summary of the model for fit statistics
    print(summary(lm_model))
    
    # Extract the tidy output and assign it to ols_results
    ols_results <- broom::tidy(lm_model) %>%
      mutate(ResponseVariable = response_var, R_Squared = r_squared, F_Statistic = f_statistic, P_Value = p_value)
    
    # Optionally print the tidy output
    print(ols_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lm_model) ~ fitted(lm_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lm_model), main = "Q-Q Plot")
      qqline(residuals(lm_model))
      dev.off()
    }
    
    # Store the results in a list
    ols_results_list[[response_var]] <- ols_results
  }
  
  return(ols_results_list)
}

ols_results_list_both_groups <- fit_ols_and_format(
  data = df_recode,
  predictors = ivs,  
  response_vars = dvs_list,       
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults_bothgroups <- bind_rows(ols_results_list_both_groups)

# Apply the same models only when willingness to pay = Yes

library(dplyr)

# Assuming your dataframe is named 'your_dataframe'
df_recode_willingtopay <- df_recode %>% 
  filter(Disposición_a_Pagar_Impuestos == "Yes")

# Now apply the function
ols_results_list <- fit_ols_and_format(
  data = df_recode_willingtopay,
  predictors = ivs,  
  response_vars = dvs_list,       
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults_onlywilling <- bind_rows(ols_results_list)

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
  "Models - Both Groups" = df_modelresults_bothgroups, 
  "Models - Only Willing to Pay" = df_modelresults_onlywilling
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

