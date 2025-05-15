setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/cooldude20202")

library(openxlsx)
library(dplyr)
df <- read.xlsx("Data collection 2024 EDIT 102.xlsx")

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

# Chi-Square Analysis

chi_square_analysis_multiple <- function(data, row_vars, col_vars) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over column variables
  for (col_var in col_vars) {
    # Ensure the column variable is a factor
    data[[col_var]] <- factor(data[[col_var]])
    
    # Iterate over row variables
    for (row_var in row_vars) {
      # Ensure the row variable is a factor
      data[[row_var]] <- factor(data[[row_var]])
      
      # Create the contingency table
      contingency_table <- table(data[[row_var]], data[[col_var]])
      
      # Create a crosstab with percentages
      crosstab <- prop.table(contingency_table, margin = 2) * 100
      
      # Check if the crosstab is 2x2
      is_2x2_table <- all(dim(contingency_table) == 2)
      
      # Perform chi-square test with correction for 2x2 tables
      chi_square_test <- chisq.test(contingency_table, correct = !is_2x2_table)
      
      # Convert crosstab to a dataframe
      crosstab_df <- as.data.frame.matrix(crosstab)
      
      # Calculate total sample size for each level
      sample_size_per_level <- apply(contingency_table, 1, sum)
      
      # Create a dataframe for this pair of variables
      for (level in levels(data[[row_var]])) {
        level_df <- data.frame(
          "Row_Variable" = row_var,
          "Row_Level" = level,
          "Column_Variable" = col_var,
          "Sample_Size" = sample_size_per_level[level], # Add sample size for the level
          check.names = FALSE
        )
        
        if (nrow(crosstab_df) > 0 && !is.null(crosstab_df[level, , drop = FALSE])) {
          level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
        }
        
        level_df$Chi_Square <- chi_square_test$statistic
        level_df$P_Value <- chi_square_test$p.value
        
        # Add the result to the list
        results[[paste0(row_var, "_", col_var, "_", level)]] <- level_df
      }
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}




# List of values to be replaced with NA in the "Diagnosis" column - Only one case
values_to_replace <- c("Apical abscess", 
                       "PA radiolucency", 
                       "Periapical abscess", 
                       "Periapical periodontitis")

# Replace specified values with NA in the "Diagnosis" column
df$Diagnosis <- ifelse(df$Diagnosis %in% values_to_replace, NA, df$Diagnosis)

# Step 1: Trim whitespaces for the entire dataframe
df <- data.frame(lapply(df, function(x) if(is.character(x)) trimws(x) else x))

# Example usage
row_variables <- c("Tooth.type."                                ,      
                   "Number.of.appointments"                    ,                              
                   "Diagnosis"                                ,         "Obturation.within.2mm.of.radiographic.apex."   ,   
                   "Presence.of.pre.operative.lesion."         ,        "Presence.of.voids."                           ,    
                   "Location.of.voids"                          ,       "Coronal.restoration.prior.to.root.canal.treatment",
                   "Coronal.restoration.post.root.canal.treatment" ,    "Irrigant.used."       ,                            
                   "Dental.dam.used.", "Crown.placed.Y_N"   ,                              
                   "OUTCOME._.STILL.PRESENT._.XLA._.RERCT")
column_variables <- c("Presence.after.1.year.Y_N.and.state.if.symptoms"  ,
                      "Presence.after.2.year.Y_N.and.state.if.symptoms" ,  "Presence.after.3.year.Y_N.and.state.if.symptoms" , 
                      "Presence.after.4.year.Y_N.and.state.if.symptoms")

# Replace "U" values with NA for the specified columns
for (col in column_variables) {
  df[[col]] <- ifelse(df[[col]] == "U", NA, df[[col]])
}

replace_u <- ("OUTCOME._.STILL.PRESENT._.XLA._.RERCT")

for (col in replace_u) {
  df[[col]] <- ifelse(df[[col]] == "U", NA, df[[col]])
}

result_df <- chi_square_analysis_multiple(df, row_variables, column_variables)


# Replace levels of variables with low representativity

replacecolumn <- ("Location.of.voids")

for (col in replacecolumn) {
  df[[col]] <- ifelse(df[[col]] == "apical", NA, df[[col]])
}

for (col in replacecolumn) {
  df[[col]] <- ifelse(df[[col]] == "mid", NA, df[[col]])
}

replacecolumn <- ("Diagnosis")

for (col in replacecolumn) {
  df[[col]] <- ifelse(df[[col]] == "Acute apical periodontitis", NA, df[[col]])
}

result_df2 <- chi_square_analysis_multiple(df, row_variables, column_variables)

# Visualizations

library(ggplot2)
library(ggrepel)

result_df2_filtered <- result_df2[!grepl("OUTCOME._.STILL.PRESENT._.XLA._.RERCT", result_df2$Row_Variable), ]

# Column charts

# Get unique row variables to create a plot for each
row_vars <- unique(result_df2_filtered$Row_Variable)

# Assuming 'result_df2_filtered' is your data frame and 'Column_Variable' is the name of your column
result_df2$Column_Variable <- factor(result_df2$Column_Variable,
                                              levels = c("Presence.after.1.year.Y_N.and.state.if.symptoms"  ,
                                                         "Presence.after.2.year.Y_N.and.state.if.symptoms" ,  "Presence.after.3.year.Y_N.and.state.if.symptoms" , 
                                                         "Presence.after.4.year.Y_N.and.state.if.symptoms"),
                                              labels = c("1 year", "2 years", "3 years", "4 years"))

# Loop through each row variable to create a plot
for (row_var in row_vars) {
  # Filter the results for the current row variable
  filtered_data <- result_df2 %>%
    filter(Row_Variable == row_var) %>%
    arrange(Column_Variable)  
  
  # Create the plot with side-by-side bars
  p <- ggplot(filtered_data, aes(x = Column_Variable, y = Y, fill = Row_Level)) +
    geom_col(position = position_dodge()) +  # Use position_dodge to place bars side by side
    geom_text(aes(label = sprintf("%.1f", Y), y = Y + 3), position = position_dodge(width = 0.9), vjust = 0, size = 3) + 
    theme_minimal() +
    labs(title = paste("Percentage of Present Teeth for", row_var),
         x = "Follow-up",
         y = "Percentage of Present Teeth") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  # Save the plot
  ggsave(filename = paste0("column_chart_", row_var, ".png"), plot = p, width = 10, height = 6, dpi = 300)
  
  # Print the plot
  print(p)
}

# Assuming 'result_df2_filtered' is your data frame and 'Column_Variable' is the name of your column
result_df2_filtered$Column_Variable <- factor(result_df2_filtered$Column_Variable,
                                              levels = c("Presence.after.1.year.Y_N.and.state.if.symptoms"  ,
                                                         "Presence.after.2.year.Y_N.and.state.if.symptoms" ,  "Presence.after.3.year.Y_N.and.state.if.symptoms" , 
                                                         "Presence.after.4.year.Y_N.and.state.if.symptoms"),
                                              labels = c("1 year", "2 years", "3 years", "4 years"))

#P value chart

library(dplyr)
result_df2_filtered <- result_df2_filtered %>%
  group_by(Column_Variable, Row_Variable) %>%
  sample_n(size = 1) %>%
  ungroup() 

# Get the unique column variables to create a plot for each
column_vars <- unique(result_df2_filtered$Column_Variable)

# Loop through each column variable to create a plot
for (col_var in column_vars) {
  # Filter the results for the current column variable
  filtered_data <- result_df2_filtered[result_df2_filtered$Column_Variable == col_var, ]
  
  # Create the plot
  p <- ggplot(filtered_data, aes(x = Chi_Square, y = -log10(P_Value), label = Row_Variable)) +
    geom_point(aes(color = Row_Variable), size = 3) +
    geom_label_repel(aes(fill = Row_Variable), size = 3, 
                     max.overlaps = Inf, box.padding = 0.35, 
                     point.padding = 0.5, segment.color = 'grey50') +
    theme_minimal() +
    labs(title = paste("Chi-Square Statistic vs. -log10(P-Value) for", col_var), 
         x = "Chi-Square Statistic", 
         y = "-log10(P-Value)") +
    xlim(0, 30) +
    geom_hline(yintercept = -log10(0.05), linetype = "dashed", color = "blue") +
    annotate("text", x = 30, y = -log10(0.05) + 0.1, 
             label = "p = 0.05", hjust = 1, color = "blue", size = 4) +
    theme(legend.position = "none")
  
  # Print and save the plot
  ggsave(filename = paste0("chi_square_plot_", col_var, ".png"), plot = p, width = 10, height = 6, dpi = 300)
  print(p)
}


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
  "Results" = result_df,
  "Results-Cleaned" = result_df2
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")



# new plot

library(ggplot2)
library(dplyr)
library(tidyr)

# Calculate the percentage of 'Y' for each year
percentages_by_year <- df %>%
  summarise(
    `1 year` = sum(`Presence.after.1.year.Y_N.and.state.if.symptoms` == "Y", na.rm = TRUE) / n(),
    `2 years` = sum(`Presence.after.2.year.Y_N.and.state.if.symptoms` == "Y", na.rm = TRUE) / n(),
    `3 years` = sum(`Presence.after.3.year.Y_N.and.state.if.symptoms` == "Y", na.rm = TRUE) / n(),
    `4 years` = sum(`Presence.after.4.year.Y_N.and.state.if.symptoms` == "Y", na.rm = TRUE) / n()
  ) %>%
  pivot_longer(cols = everything(), names_to = "Year", values_to = "Percentage") %>%
  mutate(Percentage = Percentage * 100) # Convert to percentage

# Plot
p <- ggplot(percentages_by_year, aes(x = Year, y = Percentage, fill = Year)) +
  geom_bar(stat = "identity") +
  geom_text(aes(label = paste0(sprintf("%.1f", Percentage), "%")), position = position_stack(vjust = 0.5), size = 3.5) +
  theme_minimal() +
  labs(title = "Percentage of Present Teeth by Year", x = "Year", y = "Percentage (%)") +
  scale_fill_brewer(palette = "Pastel1") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

print(p)

# Save the plot
ggsave(filename = "percentage_y_by_year.png", plot = p, width = 8, height = 6, dpi = 300)


# Dichotomize variables

library(fastDummies)

colnames(df)
dummy_vars = c("Tooth.type.", "Diagnosis", "Location.of.voids")

df_recoded <- dummy_cols(df, select_columns = dummy_vars)

names(df_recoded) <- gsub(" ", "_", names(df_recoded))
names(df_recoded) <- gsub("\\(", "_", names(df_recoded))
names(df_recoded) <- gsub("\\)", "_", names(df_recoded))
names(df_recoded) <- gsub("\\-", "_", names(df_recoded))
names(df_recoded) <- gsub("/", "_", names(df_recoded))
names(df_recoded) <- gsub("\\\\", "_", names(df_recoded)) 
names(df_recoded) <- gsub("\\?", "", names(df_recoded))
names(df_recoded) <- gsub("\\'", "", names(df_recoded))
names(df_recoded) <- gsub("\\,", "_", names(df_recoded))

colnames(df_recoded)

row_variables <- c("Tooth.type._Anterior"      ,                       
                   "Tooth.type._Molar"                ,                 "Tooth.type._Premolar"  ,                                  
                   "Number.of.appointments"                    ,                              
                   "Diagnosis_Acute_periapical_periodontitis"     ,     "Diagnosis_Chronic_apical_periodontitis"  ,         
                   "Diagnosis_Chronic_periapical_periodontitis"   ,     "Diagnosis_Irreversible_pulpitis"   ,               
                   "Diagnosis_NA"                                 ,         "Obturation.within.2mm.of.radiographic.apex."   ,   
                   "Presence.of.pre.operative.lesion."         ,        "Presence.of.voids."                           ,    
                   "Location.of.voids_All"          ,                  
                   "Location.of.voids_coronal"        ,                 "Location.of.voids_no"   ,                          
                   "Location.of.voids_NA"                          ,       "Coronal.restoration.prior.to.root.canal.treatment",
                   "Coronal.restoration.post.root.canal.treatment" ,    "Irrigant.used."       ,                            
                   "Dental.dam.used.", "Crown.placed.Y_N"   ,                              
                   "OUTCOME._.STILL.PRESENT._.XLA._.RERCT")

column_variables <- c("Presence.after.1.year.Y_N.and.state.if.symptoms"  ,
                      "Presence.after.2.year.Y_N.and.state.if.symptoms" ,  "Presence.after.3.year.Y_N.and.state.if.symptoms" , 
                      "Presence.after.4.year.Y_N.and.state.if.symptoms")

# Revised code to invert percentages

chi_square_analysis_multiple <- function(data, row_vars, col_vars) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over column variables
  for (col_var in col_vars) {
    # Ensure the column variable is a factor
    data[[col_var]] <- factor(data[[col_var]])
    
    # Iterate over row variables
    for (row_var in row_vars) {
      # Ensure the row variable is a factor
      data[[row_var]] <- factor(data[[row_var]])
      
      # Create the contingency table
      contingency_table <- table(data[[row_var]], data[[col_var]])
      
      # Create a crosstab with percentages
      crosstab <- prop.table(contingency_table) * 100
      
      # Check if the crosstab is 2x2
      is_2x2_table <- all(dim(contingency_table) == 2)
      
      # Perform chi-square test with correction for 2x2 tables
      chi_square_test <- chisq.test(contingency_table, correct = !is_2x2_table)
      
      # Convert crosstab to a dataframe
      crosstab_df <- as.data.frame.matrix(crosstab)
      
      # Calculate total sample size for each level
      sample_size_per_level <- apply(contingency_table, 1, sum)
      
      # Create a dataframe for this pair of variables
      for (level in levels(data[[row_var]])) {
        level_df <- data.frame(
          "Row_Variable" = row_var,
          "Row_Level" = level,
          "Column_Variable" = col_var,
          "Sample_Size" = sample_size_per_level[level], # Add sample size for the level
          check.names = FALSE
        )
        
        if (nrow(crosstab_df) > 0 && !is.null(crosstab_df[level, , drop = FALSE])) {
          level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
        }
        
        level_df$Chi_Square <- chi_square_test$statistic
        level_df$P_Value <- chi_square_test$p.value
        
        # Add the result to the list
        results[[paste0(row_var, "_", col_var, "_", level)]] <- level_df
      }
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}



result_df_totalperc <- chi_square_analysis_multiple(df_recoded, row_variables, column_variables)

result_df_totalperc$Column_Variable <- factor(result_df_totalperc$Column_Variable,
                                     levels = c("Presence.after.1.year.Y_N.and.state.if.symptoms"  ,
                                                "Presence.after.2.year.Y_N.and.state.if.symptoms" ,  "Presence.after.3.year.Y_N.and.state.if.symptoms" , 
                                                "Presence.after.4.year.Y_N.and.state.if.symptoms"),
                                     labels = c("1 year", "2 years", "3 years", "4 years"))

result_df_totalperc <- result_df_totalperc[!grepl("OUTCOME._.STILL.PRESENT._.XLA._.RERCT", result_df_totalperc$Row_Variable), ]


chi_square_analysis_multiple <- function(data, row_vars, col_vars) {
  results <- list() # Initialize an empty list to store results
  
  # Iterate over column variables
  for (col_var in col_vars) {
    # Ensure the column variable is a factor
    data[[col_var]] <- factor(data[[col_var]])
    
    # Iterate over row variables
    for (row_var in row_vars) {
      # Ensure the row variable is a factor
      data[[row_var]] <- factor(data[[row_var]])
      
      # Create the contingency table
      contingency_table <- table(data[[row_var]], data[[col_var]])
      
      # Create a crosstab with percentages
      crosstab <- prop.table(contingency_table, margin = 2) * 100
      
      # Check if the crosstab is 2x2
      is_2x2_table <- all(dim(contingency_table) == 2)
      
      # Perform chi-square test with correction for 2x2 tables
      chi_square_test <- chisq.test(contingency_table, correct = !is_2x2_table)
      
      # Convert crosstab to a dataframe
      crosstab_df <- as.data.frame.matrix(crosstab)
      
      # Calculate total sample size for each level
      sample_size_per_level <- apply(contingency_table, 1, sum)
      
      # Create a dataframe for this pair of variables
      for (level in levels(data[[row_var]])) {
        level_df <- data.frame(
          "Row_Variable" = row_var,
          "Row_Level" = level,
          "Column_Variable" = col_var,
          "Sample_Size" = sample_size_per_level[level], # Add sample size for the level
          check.names = FALSE
        )
        
        if (nrow(crosstab_df) > 0 && !is.null(crosstab_df[level, , drop = FALSE])) {
          level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
        }
        
        level_df$Chi_Square <- chi_square_test$statistic
        level_df$P_Value <- chi_square_test$p.value
        
        # Add the result to the list
        results[[paste0(row_var, "_", col_var, "_", level)]] <- level_df
      }
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}

result_df_dummy <- chi_square_analysis_multiple(df_recoded, row_variables, column_variables)

result_df_dummy$Column_Variable <- factor(result_df_dummy$Column_Variable,
                                              levels = c("Presence.after.1.year.Y_N.and.state.if.symptoms"  ,
                                                         "Presence.after.2.year.Y_N.and.state.if.symptoms" ,  "Presence.after.3.year.Y_N.and.state.if.symptoms" , 
                                                         "Presence.after.4.year.Y_N.and.state.if.symptoms"),
                                              labels = c("1 year", "2 years", "3 years", "4 years"))

result_df_dummy <- result_df_dummy[!grepl("OUTCOME._.STILL.PRESENT._.XLA._.RERCT", result_df_dummy$Row_Variable), ]


# Plot

# Assuming result_df_rowperc is your dataframe
result_df_dummy$Row_Level <- ifelse(result_df_dummy$Row_Level == 'Y', 'Yes',
                                    ifelse(result_df_dummy$Row_Level == 'N', 'No', result_df_dummy$Row_Level))


library(ggplot2)

# Replace '.' and '_' with spaces in 'Row_Variable'
result_df_rowperc$Row_Variable <- gsub("[._]", " ", result_df_rowperc$Row_Variable)

# Ensure 'Column_Variable' is treated as a factor
result_df_rowperc$Column_Variable <- factor(result_df_rowperc$Column_Variable, levels = unique(result_df_rowperc$Column_Variable))

# Unique Row_Variables
unique_row_vars <- unique(result_df_rowperc$Row_Variable)

# Loop through each Row_Variable and create a separate plot
for (row_var in unique_row_vars) {
  # Filter data for the current Row_Variable
  filtered_data <- subset(result_df_rowperc, Row_Variable == row_var)
  
  # Generate and print the plot
  p <- ggplot(filtered_data, aes(x = Row_Level, y = Y, fill = Row_Level)) + 
    geom_col(position = "dodge") +
    geom_text(aes(label = round(Y, 1)), position = position_dodge(width = 0.9), vjust = -0.25, size = 3) +
    theme_minimal() +
    labs(x = "Row Level", y = "Value", title = paste("Column Chart for", row_var), subtitle = "With Categories in Row_Level") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  # Display the plot
  print(p)
}

# Example usage
data_list <- list(
  "Results" = result_df,
  "Results-Cleaned" = result_df2
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_v2.xlsx")



