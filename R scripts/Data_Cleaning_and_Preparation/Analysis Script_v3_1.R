setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/cooldude20202")

library(openxlsx)
library(dplyr)
df <- read.xlsx("Data collection 2024 EDIT 105.xlsx")

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
      crosstab <- prop.table(contingency_table, margin = 1) * 100
      
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
row_variables <- c("Crown.placed.Y_N"                               
                   )
column_variables <- c("Presence.after.1.year.Y_N"  ,
                      "Presence.after.2.year.Y_N" ,  "Presence.after.3.year.Y_N" , 
                      "Presence.after.4.year.Y_N")
colnames(df)
# Replace "U" values with NA for the specified columns
for (col in column_variables) {
  df[[col]] <- ifelse(df[[col]] == "U", NA, df[[col]])
}

replace_u <- c("OUTCOME._.STILL.PRESENT._.XLA._.RERCT", "Crown.placed.Y_N")

for (col in replace_u) {
  df[[col]] <- ifelse(df[[col]] == "U", NA, df[[col]])
}

result_df3 <- chi_square_analysis_multiple(df, row_variables, column_variables)

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
      crosstab <- prop.table(contingency_table, margin = NULL) * 100
      
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


result_dfsum100 <- chi_square_analysis_multiple(df, row_variables, column_variables)



# Column charts

# Get unique row variables to create a plot for each
row_vars <- unique(result_df3$Row_Variable)

result_df3$Column_Variable <- factor(result_df3$Column_Variable,
                                              levels = c("Presence.after.1.year.Y_N"  ,
                                                         "Presence.after.2.year.Y_N" ,  "Presence.after.3.year.Y_N" , 
                                                         "Presence.after.4.year.Y_N"),
                                              labels = c("1 year", "2 years", "3 years", "4 years"))

result_df3 <- result_df3[!grepl("OUTCOME._.STILL.PRESENT._.XLA._.RERCT", result_df3$Row_Variable), ]


# Plot


library(ggplot2)

# Replace '.' and '_' with spaces in 'Row_Variable'
result_df3$Row_Variable <- gsub("[._]", " ", result_df3$Row_Variable)

# Ensure 'Column_Variable' is treated as a factor
result_df3$Column_Variable <- factor(result_df3$Column_Variable, levels = unique(result_df3$Column_Variable))

result_df3$Column_Variable <- factor(result_df3$Column_Variable, levels = c("1 year", "2 years", "3 years", "4 years"))

# Unique Row_Variables
unique_row_vars <- unique(result_df3$Row_Variable)

# Loop through each Row_Variable and create a separate plot
for (row_var in unique_row_vars) {
  
  # Filter data for the current Row_Variable
  filtered_data <- result_df3 %>%
    filter(Row_Variable == row_var)
  
  # Generate the plot
  p <- ggplot(filtered_data, aes(x = Column_Variable, y = Y, fill = Row_Level)) + 
    geom_col(position = position_dodge(width = 0.75), width = 0.7) +  # Columns side by side
    geom_text(aes(label = ifelse(Y > 0, round(Y, 1), "")), 
              position = position_dodge(width = 0.75), vjust = 0.5, size = 3, check_overlap = TRUE) +  # Centered labels
    theme_minimal() +
    labs(fill = row_var,  # Set the legend title to the value of 'Row_Variable'
         x = "Year", 
         y = "Percentage", 
         title = paste("Percentage of Present Tooth by", row_var)) +
    theme(legend.position = "bottom", axis.text.x = element_text(angle = 45, hjust = 1)) +
    scale_fill_brewer(palette = "Pastel1")  # Optional: Use a color palette
  
  # Save the plot
  ggsave(filename = paste0("individual_plot_", gsub("[ ./]", "_", row_var), ".png"), plot = p, width = 8, height = 4)
  
  # Display the plot (for interactive sessions)
  print(p)
}








# Analysis by tooth type

row_variables <- c("Crown.placed.Y_N")

column_variables <- c("Presence.after.1.year.Y_N"  ,
                      "Presence.after.2.year.Y_N" ,  "Presence.after.3.year.Y_N" , 
                      "Presence.after.4.year.Y_N")

# Subset data

anterior_df <- subset(df, Tooth.type. == "Single")
posterior_df <- subset(df, Tooth.type. == "Multiple")

chi_square_analysis_multiple <- function(data, row_vars, col_vars, margin = NULL) {
  results <- list() # Initialize an empty list to store results
  
  # Validate margin input
  if (!is.null(margin) && !margin %in% c(1, 2)) {
    stop("Margin must be either 1 (row), 2 (column), or NULL (total sample).")
  }
  
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
      
      # Calculate proportions based on the specified margin
      if (is.null(margin)) {
        # If margin is NULL, calculate proportions for the grand total
        crosstab <- prop.table(contingency_table) * 100
      } else {
        # Calculate proportions for the specified margin
        crosstab <- prop.table(contingency_table, margin = margin) * 100
      }
      
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

result_df_posterior <- chi_square_analysis_multiple(posterior_df, row_variables, column_variables, margin = NULL)

result_df_posterior_2 <- chi_square_analysis_multiple(posterior_df, row_variables, column_variables, margin = 1)

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
  "Results - Row %" = result_df3,
  "Results - Total %" = result_dfsum100,
  "Results - Posterior" = result_df_posterior_2,
  "Results - POsterior % All" = result_df_posterior
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_v10.xlsx")


library(ggplot2)
library(dplyr)
library(tidyr)

# New PLOTS


# Pivot longer to have a single 'Status' column for plotting
result_df_long <- result_dfsum100 %>%
  pivot_longer(cols = c(N, Y), names_to = "Status", values_to = "Percentage") %>%
  mutate(
    Column_Variable = case_when(
      Column_Variable == "Presence.after.1.year.Y_N" ~ "1 Year",
      Column_Variable == "Presence.after.2.year.Y_N" ~ "2 Years",
      Column_Variable == "Presence.after.3.year.Y_N" ~ "3 Years",
      Column_Variable == "Presence.after.4.year.Y_N" ~ "4 Years",
      TRUE ~ Column_Variable
    ),
    Row_Variable = ifelse(Row_Variable == "Crown.placed.Y_N", "Crown Placed", Row_Variable),
    Status = ifelse(Status == "Y", "Survided", "Did Not Survive"),
    Row_Level = ifelse(Row_Level == "Y", "Yes", "No")
  )

# Plot
ggplot(result_df_long, aes(x = Column_Variable, y = Percentage, fill = Status)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.75)) +
  geom_text(aes(label = paste0(round(Percentage, 1), "%")), 
            position = position_dodge(width = 0.75), vjust = 0.0, size = 3.5) +
  facet_wrap(~Row_Variable + Row_Level, scales = "free_x", ncol = 1) +
  labs(title = "Tooth Survival Analysis", x = "Time After Procedure", y = "Percentage", fill = "Tooth Status") +
  scale_fill_brewer(palette = "Pastel1") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))








library(ggplot2)
library(dplyr)
library(tidyr)

# Assuming 'result_dfsum100' is already loaded with the proper structure

# Filter the data for the 4-year follow-up
four_year_data <- result_dfsum100 %>%
  filter(grepl("4.year", Column_Variable)) %>%
  pivot_longer(cols = c(N, Y), names_to = "Status", values_to = "Percentage") %>%
  mutate(
    Row_Variable = ifelse(Row_Variable == "Crown.placed.Y_N", "Crown Placed", "No Crown Placed"),
    Status = ifelse(Status == "Y", "Survived", "Did Not Survive"),
    Row_Level = ifelse(Row_Level == "Y", "Yes", "No")
  )

# Plot for the 4-year follow-up
p <- ggplot(four_year_data, aes(x = Row_Level, y = Percentage, fill = Status)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.75)) +
  geom_text(aes(label = paste0(round(Percentage, 1), "%")), 
            position = position_dodge(width = 0.75), vjust = 1.5, size = 3.5) +
  labs(title = "4-Year Tooth Survival Analysis", x = "Crown Placement", y = "Percentage", fill = "Tooth Status") +
  scale_fill_brewer(palette = "Pastel1") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Print the plot
print(p)