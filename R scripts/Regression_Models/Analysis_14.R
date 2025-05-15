setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/beasalazar")

library(readxl)
df <- read_xlsx('Encuestas_Resultados.xlsx', sheet = 'Clean_Data')

# Get rid of special characters

names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))

# Load dplyr for data manipulation
library(dplyr)

# Recode the data
df_recode <- df %>%
  mutate(
    Disposición_a_Pagar_Impuestos = as.factor(ifelse(Disposición_a_Pagar_Impuestos == 1, "Yes", "No")),
    Motivos_para_no_Pagar = as.factor(case_when(
      Motivos_para_no_Pagar == 1 ~ "Economic constraints",
      Motivos_para_no_Pagar == 2 ~ "Prefer to use money on other things",
      Motivos_para_no_Pagar == 3 ~ "Prefer ecological almond grove as is",
      Motivos_para_no_Pagar == 4 ~ "Lack of trust in administration",  # Protest answer
      Motivos_para_no_Pagar == 5 ~ "Administrations should subsidize",  # Protest answer
      Motivos_para_no_Pagar == 6 ~ "Already pay enough taxes",
      TRUE ~ NA_character_ 
    )),
    Sexo = as.factor(case_when(
      Sexo == 1 ~ "Female",
      Sexo == 2 ~ "Male",
      TRUE ~ NA_character_
    )),
    Nivel_máximo_de_estudios = as.factor(case_when(
      Nivel_máximo_de_estudios == 1 ~ "None",
      Nivel_máximo_de_estudios == 2 ~ "Primary",
      Nivel_máximo_de_estudios == 3 ~ "Secondary",
      Nivel_máximo_de_estudios == 4 ~ "Baccalaureate",
      Nivel_máximo_de_estudios == 5 ~ "University education",
      Nivel_máximo_de_estudios == 6 ~ "PhD",
      TRUE ~ NA_character_
    )),
    Ingresos_mensuales_netos = as.factor(case_when(
      Ingresos_mensuales_netos == 1 ~ "<1000€",
      Ingresos_mensuales_netos == 2 ~ "1000-2000€",
      Ingresos_mensuales_netos == 3 ~ "2000-3000€",
      Ingresos_mensuales_netos == 4 ~ ">3000€",
      TRUE ~ NA_character_
    )),
    Relación_con_la_Agricultura_Orgánica = as.factor(ifelse(Relación_con_la_Agricultura_Orgánica == 1, "Yes", "No"))
  )

# View the first few rows of the recoded data
head(df_recode)

# Categorize 'Edad' into age groups
df_recode <- df_recode %>%
  mutate(
    Edad_Cat = cut(Edad,
                   breaks = c(13, 25, 35, 60, Inf),
                   labels = c("14-25", "26-35", "36-60", "61-older"),
                   right = FALSE)
  )


# FREQUENCY TABLE

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

vars = c("Relación_con_la_Agricultura_Orgánica","Disposición_a_Pagar_Impuestos", "Motivos_para_no_Pagar", "Monto_de_Impuestos",
         "Sexo", "Edad_Cat", "Nivel_máximo_de_estudios", "Ingresos_mensuales_netos")
freq_df <- create_frequency_tables(df_recode, vars)


## Chi-Square Analysis

chi_square_analysis_multiple <- function(data, row_vars, col_var) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over row variables
  for (row_var in row_vars) {
    # Ensure that the variables are factors
    data[[row_var]] <- factor(data[[row_var]])
    data[[col_var]] <- factor(data[[col_var]])
    
    # Create a crosstab with row percentages
    crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 1) * 100
    
    # Perform chi-square test
    chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]))
    
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
      
      # Extract the percentages for each level
      level_percentages <- crosstab_df[level, , drop = FALSE]
      
      # Add percentages to the dataframe
      level_df[paste0(col_var, "_No_Percentage")] <- level_percentages[,1]
      level_df[paste0(col_var, "_Yes_Percentage")] <- level_percentages[,2]
      
      level_df$Chi_Square <- chi_square_test$statistic
      level_df$P_Value <- chi_square_test$p.value
      
      # Add the result to the list
      results[[paste0(row_var, "_", level)]] <- level_df
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}

row_variables <- c("Relación_con_la_Agricultura_Orgánica",
                   "Sexo", "Edad_Cat", "Nivel_máximo_de_estudios", "Ingresos_mensuales_netos")  
column_variable <- "Disposición_a_Pagar_Impuestos"

result_df <- chi_square_analysis_multiple(df_recode, row_variables, column_variable)
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
  
  # Extract coefficients, Odds Ratios, and p-values
  coefficients <- coef(model)
  odds_ratios <- exp(coefficients)
  p_values <- summary(model)$coef[, 4]
  
  # Initialize a variable to store predictors
  predictors <- names(coefficients)
  
  # Process the names to combine variable and level with a separator
  predictors <- sapply(predictors, function(name) {
    parts <- strsplit(name, split = ":", fixed = TRUE)[[1]]
    if (length(parts) > 1) {
      return(paste(parts[1], parts[2], sep = ": "))
    } else {
      return(parts[1])
    }
  })
  
  # Create a dataframe to store results
  results_df <- data.frame(
    Predictor = predictors,
    Coefficients = coefficients,
    Odds_Ratios = odds_ratios,
    P_Values = p_values
  )
  
  # Print R-squared and overall model summary
  cat("McFadden's Pseudo R-squared: ", r_squared, "\n")
  
  # Return results dataframe
  return(results_df)
}

dependent_variable <- "Disposición_a_Pagar_Impuestos" 
independent_variables <- c("Relación_con_la_Agricultura_Orgánica",
                           "Sexo", "Edad_Cat", "Nivel_máximo_de_estudios", "Ingresos_mensuales_netos") 

df_recode$Nivel_máximo_de_estudios <- relevel(df_recode$Nivel_máximo_de_estudios, ref = "Primary")

model_results <- probit_regression(df_recode, dependent_variable, independent_variables)
print(model_results)

#BOXPLOTS

library(ggplot2)
library(tidyr)

# Select the variables for which you want to create boxplots
vars <- c("Polinización",                
          "Hábitat_para_especies"  ,              "Descomposición_de_la_MO"        ,      "Educación_ambiental"  ,               
          "Regulación_del_clima"         ,        "Turismo_de_naturaleza_y_rural"     ,  "Conocimiento_científico",             
          "Dispersión_de_semillas"          ,     "Identidad_local"              ,        "Control_de_plagas"      ,             
          "Agua_para_riego_y_consumo"         ,   "Provisión_de_alimentos"       ,        "Conocimiento_local_tradicional")

# Reshape the data to a long format
long_df <- df %>%
  select(one_of(vars)) %>%
  gather(key = "Variable", value = "Value")

# Create side-by-side boxplots for each variable
p <- ggplot(long_df, aes(x = Variable, y = Value)) +
  geom_boxplot() +
  stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
  labs(title = "Distribución de impuesto en diferentes servicios ecosistémicos",
       x = "Variable", y = "Valor") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) # Rotate x-axis labels


# Print the plot
print(p)

# Save the plot
ggsave(filename = "boxplots.png", plot = p, width = 10, height = 6)


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
  "Frequency Table" = freq_df, 
  "Chi-Square Results" = result_df, 
  "Probit Model results" = model_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")
