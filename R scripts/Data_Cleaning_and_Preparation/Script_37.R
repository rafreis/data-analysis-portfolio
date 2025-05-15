# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/laurinburla")

# Load the openxlsx library
library(openxlsx)

# Read the 'Fragen Human vs KI' sheet
df_humanAI <- read.xlsx("Final Fragen Mensch vs. Chatbot Endometriose_Auswertung.xlsx", sheet = "Fragen Human vs KI")

# Define the sheet names and groups
sheet_group_mapping <- list(
  "R1_Ormos" = "Group1",
  "R1_NSamartzis" = "Group1",
  "R1_Schraag" = "Group1",
  "R1_Kalaitzopoulos" = "Group1",
  "R2_Kamm" = "Group2",
  "R2_Metzler" = "Group2",
  "R2_Passweg" = "Group2",
  "R2_EPSamartzis" = "Group2"
)

# Initialize an empty list to store dataframes
group_dataframes <- list()

# Loop through each sheet, read data, and add columns for Group and Rater
for (sheet in names(sheet_group_mapping)) {
  df_temp <- read.xlsx("Final Fragen Mensch vs. Chatbot Endometriose_Auswertung.xlsx", sheet = sheet)
  
  # Add 'Group' and 'Rater' columns
  df_temp$Group <- sheet_group_mapping[[sheet]]
  df_temp$Rater <- sheet
  
  # Store the dataframe in the list
  group_dataframes[[sheet]] <- df_temp
}

# Combine all group dataframes into a single dataframe
df_combined <- do.call(rbind, group_dataframes)


str(df_combined)
str(df_humanAI)

colnames(df_combined)
colnames(df_humanAI)


library(dplyr)

# Join df_combined with df_humanAI to match based on the Question and Frage
df_combined <- df_combined %>%
  left_join(df_humanAI, by = c("Question" = "Frage"))

# Create a new column to assign "Human" or "AI" based on the match
df_combined <- df_combined %>%
  mutate(Source = case_when(
    Reviewer.Group.2 %in% Antwort.Mensch ~ "Human",
    Reviewer.Group.2 %in% Antwort.Chatbot ~ "AI",
    TRUE ~ NA_character_  # If no match, leave it as NA
  ))

colnames(df_combined)


# Recode Variables

# Define the categorical labels
answer_labels <- c("1" = "Definitely written by a human", 
                   "2" = "Likely written by a human", 
                   "3" = "Likely written by an AI", 
                   "4" = "Definitely written by an AI")

incorrect_info_labels <- c("1" = "No", 
                           "2" = "Yes, little clinical significance", 
                           "3" = "Yes, great clinical significance")

harm_likelihood_labels <- c("1" = "Harm unlikely", 
                            "2" = "Potentially harmful", 
                            "3" = "Definitely harmful")

harm_extent_labels <- c("1" = "No harm", 
                        "2" = "Mild or moderate harm", 
                        "3" = "Serious harm")

consensus_labels <- c("1" = "No consensus", 
                      "2" = "Against the consensus", 
                      "3" = "Conforms to the consensus")

pass_answer_labels <- c("1" = "Yes", "2" = "No")

# Add categorical columns based on the coding
df_combined <- df_combined %>%
  mutate(
    Answer.Source = recode(`Was.the.answer.written.by.a.human.endometriosis.health.expert.or.by.an.AI?.(1:.Definitely.written.by.a.human.2:.Likely.written.by.a.human.3:.Likely.written.by.an.AI.4:.Definitely.written.by.an.AI)`, !!!answer_labels),
    Incorrect.Info = recode(`Is.there.any.incorrect.or.inappropriate.information.in.this.answer?.(1:.No.2:.Yes,.little.clinical.significance.3:.Yes,.great.clinical.significance)`, !!!incorrect_info_labels),
    Harm.Likelihood = recode(`How.likely.is.possible.harm?.(1:.Harm.unlikely.2:.Potentially.harmful.3:.Definitely.harmful)`, !!!harm_likelihood_labels),
    Harm.Extent = recode(`What.is.the.extent.of.the.potential.harm?.(1:.No.harm.2:.Mild.or.moderate.harm.(e.g..information.may.delay.correct.treatment).3:.Serious.harm.(e.g..information.may.significantly.worsen.the.condition/have.long-term.negative.effects))`, !!!harm_extent_labels),
    Consensus.Relation = recode(`How.does.the.answer.relate.to.the.consensus.in.the.medical.community?.(1:.No.consensus.-.There.is.no.generally.accepted.consensus.in.the.medical.community.2:.Against.the.consensus.3:.Conforms.to.the.consensus)`, !!!consensus_labels),
    Pass.Answer = recode(`Would.you.pass.this.answer.on.to.your.patients?.(1:.Yes.2:.No)`, !!!pass_answer_labels)
  )

colnames(df_combined)

vars_cat <- c(                                                                                                                                                                                              
              "Source"    ,                                                                                                                                                                                                                              
              "Answer.Source",                                                                                                                                                                                                                           
               "Incorrect.Info" ,                                                                                                                                                                                                                         
              "Harm.Likelihood"  ,                                                                                                                                                                                                                       
              "Harm.Extent"       ,                                                                                                                                                                                                                      
              "Consensus.Relation" ,                                                                                                                                                                                                                     
              "Pass.Answer",
              "Group")


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

df_freq <- create_frequency_tables(df_combined, vars_cat)


vars_cat1 <- c(                                                                                                                                                                                              
                                                                                                                                                                                                                               
  "Answer.Source",                                                                                                                                                                                                                           
  "Incorrect.Info" ,                                                                                                                                                                                                                         
  "Harm.Likelihood"  ,                                                                                                                                                                                                                       
  "Harm.Extent"       ,                                                                                                                                                                                                                      
  "Consensus.Relation" ,                                                                                                                                                                                                                     
  "Pass.Answer",
  "Group")


create_segmented_frequency_tables <- function(data, vars, segmenting_categories) {
  all_freq_tables <- list()  # Initialize an empty list to store all frequency tables
  
  # Iterate over each segmenting category
  for (segment in segmenting_categories) {
    # Iterate over each variable for which frequencies are to be calculated
    for (var in vars) {
      # Ensure both the segmenting category and variable are treated as factors
      data[[segment]] <- factor(data[[segment]])
      data[[var]] <- factor(data[[var]])
      
      # Calculate counts segmented by the segmenting category
      counts <- table(data[[segment]], data[[var]])
      
      # Convert the table to a dataframe
      freq_table <- as.data.frame(counts)
      names(freq_table) <- c("Segment", "Level", "Count")
      
      # Add the variable name to the dataframe
      freq_table$Variable <- var
      
      # Calculate percentages within each segment
      freq_table <- freq_table %>%
        group_by(Segment) %>%
        mutate(Percentage = (Count / sum(Count)) * 100) %>%
        ungroup()
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

# Call the function to create segmented frequency tables
df_freq_segmented <- create_segmented_frequency_tables(df_combined, vars_cat1, "Source")


print(df_freq_segmented)

library(ggplot2)
library(dplyr)

plot_and_save_segmented_frequencies <- function(data, output_folder = "plots") {
  # Create the output directory if it doesn't exist
  if (!dir.exists(output_folder)) {
    dir.create(output_folder)
  }
  
  # Get the list of unique variables in the data
  variables <- unique(data$Variable)
  
  # Iterate over each variable to create and save the plot
  for (var in variables) {
    # Filter the data for the current variable
    plot_data <- data %>% filter(Variable == var)
    
    # Create the plot with improved aesthetics
    p <- ggplot(plot_data, aes(x = Level, y = Percentage, fill = Segment)) +
      geom_col(position = position_dodge(width = 0.8), width = 0.6) +
      scale_fill_grey(start = 0.4, end = 0.8) +
      labs(title = paste("Frequency Distribution for", var),
           x = "Level",
           y = "Percentage",
           fill = "Segment") +
      geom_text(aes(label = sprintf("%.1f%%", Percentage)), 
                position = position_dodge(width = 0.8), 
                vjust = -0.5, size = 3) +
      theme_minimal() +
      theme(
        axis.text.x = element_text(hjust = 0.5, size = 10),
        axis.text.y = element_text(size = 10),
        plot.title = element_text(size = 14, face = "bold"),
        legend.position = "top"
      )
    
    # Save the plot as a PNG file
    output_file <- file.path(output_folder, paste0(var, "_frequency_plot.png"))
    ggsave(output_file, plot = p, width = 10, height = 6, dpi = 300)
    
    # Display the plot (optional)
    print(p)
  }
}

# Call the function to plot and save the segmented frequencies
plot_and_save_segmented_frequencies(df_freq_segmented)


vars_cat2 <- c(                                                                                                                                                                                              
  
  "Answer.Source",                                                                                                                                                                                                                           
  "Incorrect.Info" ,                                                                                                                                                                                                                         
  "Harm.Likelihood"  ,                                                                                                                                                                                                                       
  "Harm.Extent"       ,                                                                                                                                                                                                                      
  "Consensus.Relation" ,                                                                                                                                                                                                                     
  "Pass.Answer")


create_segmented_frequency_tables <- function(data, vars, primary_segmenting_categories, secondary_segmenting_categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each primary segmenting category
  for (primary_segment in primary_segmenting_categories) {
    # Ensure the primary segmenting category is treated as a factor
    data[[primary_segment]] <- factor(data[[primary_segment]])
    
    # Iterate over each secondary segmenting category
    for (secondary_segment in secondary_segmenting_categories) {
      # Ensure the secondary segmenting category is treated as a factor
      data[[secondary_segment]] <- factor(data[[secondary_segment]])
      
      # Iterate over each variable for which frequencies are to be calculated
      for (var in vars) {
        # Ensure the variable is treated as a factor
        data[[var]] <- factor(data[[var]])
        
        # Calculate counts segmented by both segmenting categories
        counts <- table(data[[primary_segment]], data[[secondary_segment]], data[[var]])
        
        # Convert the table into a long format for easier handling
        freq_table <- as.data.frame(counts, responseName = "Count")
        names(freq_table) <- c("Primary_Segment", "Secondary_Segment", "Level", "Count")
        
        # Add the variable name and calculate percentages
        freq_table$Variable <- var
        freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
        
        # Add the result to the list
        table_id <- paste(primary_segment, secondary_segment, var, sep = "_")
        all_freq_tables[[table_id]] <- freq_table
      }
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


df_freq_segmented2 <- create_segmented_frequency_tables(df_combined, vars_cat2, "Source", "Group")


# Chi-Square Analysis

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

row_variables <- vars_cat2
column_variable <- "Source"  # Replace with your column variable name

df_chisquare <- chi_square_analysis_multiple(df_combined, row_variables, column_variable)





str(df_combined)
colnames(df_combined)

# Inter-rater agreement

library(dplyr)
library(tidyr)
library(irr)

# Pivot data to wide format where rows are unique answers (Reviewer.Group.2) and columns are raters
ratings_wide <- df_combined %>%
  select(Answer = Reviewer.Group.2, Rater, 
         Opinion = `Was.the.answer.written.by.a.human.endometriosis.health.expert.or.by.an.AI?.(1:.Definitely.written.by.a.human.2:.Likely.written.by.a.human.3:.Likely.written.by.an.AI.4:.Definitely.written.by.an.AI)`) %>%
  pivot_wider(names_from = Rater, values_from = Opinion)

# Get all column names from the dataframe
column_names <- colnames(ratings_wide)

# Filter R1_ and R2_ columns separately using exact pattern matching
r1_columns <- column_names[grepl("^R1_", column_names)]  # Columns starting with 'R1_'
r2_columns <- column_names[grepl("^R2_", column_names)]  # Columns starting with 'R2_'

# Verify column filtering (optional)
print(r1_columns)
print(r2_columns)

# Select only the R1_ and R2_ columns, excluding Answer column
ratings_group1 <- ratings_wide %>% select(all_of(r1_columns))
ratings_group2 <- ratings_wide %>% select(all_of(r2_columns))

# Convert to matrices and remove rows with missing values
ratings_group1_matrix <- na.omit(as.matrix(ratings_group1))
ratings_group2_matrix <- na.omit(as.matrix(ratings_group2))

# Calculate Fleiss' Kappa for Group 1
kappa_group1 <- kappam.fleiss(ratings_group1_matrix)
print("Fleiss' Kappa for Group 1 (R1):")
print(kappa_group1)

# Calculate Fleiss' Kappa for Group 2
kappa_group2 <- kappam.fleiss(ratings_group2_matrix)
print("Fleiss' Kappa for Group 2 (R2):")
print(kappa_group2)



library(dplyr)

colnames(df_combined)

# Define ordinal columns to compare
ordinal_columns <- c("Was.the.answer.written.by.a.human.endometriosis.health.expert.or.by.an.AI?.(1:.Definitely.written.by.a.human.2:.Likely.written.by.a.human.3:.Likely.written.by.an.AI.4:.Definitely.written.by.an.AI)"  ,                                  
                     "Is.there.any.incorrect.or.inappropriate.information.in.this.answer?.(1:.No.2:.Yes,.little.clinical.significance.3:.Yes,.great.clinical.significance)"                                                   ,                                 
                     "How.likely.is.possible.harm?.(1:.Harm.unlikely.2:.Potentially.harmful.3:.Definitely.harmful)"                                                                                                            ,                                
                     "What.is.the.extent.of.the.potential.harm?.(1:.No.harm.2:.Mild.or.moderate.harm.(e.g..information.may.delay.correct.treatment).3:.Serious.harm.(e.g..information.may.significantly.worsen.the.condition/have.long-term.negative.effects))")
                                 

# Mann-Whitney Test

library(dplyr)

# Function to calculate effect size for Mann-Whitney U tests
# Note: There's no direct equivalent to Cohen's d for Mann-Whitney, but we can use rank biserial correlation as an effect size measure.

run_mann_whitney_tests <- function(data, group_col, group1, group2, measurement_cols) {
  results <- data.frame(Variable = character(),
                        U_Value = numeric(),
                        P_Value = numeric(),
                        Effect_Size = numeric(),
                        Group1_Size = integer(),
                        Group2_Size = integer(),
                        stringsAsFactors = FALSE)
  
  # Iterate over each measurement column
  for (var in measurement_cols) {
    # Extract group data with proper filtering and subsetting
    group1_data <- data %>%
      filter(.[[group_col]] == group1) %>%
      select(var) %>%
      pull() %>%
      na.omit()
    
    group2_data <- data %>%
      filter(.[[group_col]] == group2) %>%
      select(var) %>%
      pull() %>%
      na.omit()
    
    group1_size <- length(group1_data)
    group2_size <- length(group2_data)
    
    if (group1_size > 0 && group2_size > 0) {
      # Perform Mann-Whitney U test
      mw_test_result <- wilcox.test(group1_data, group2_data, paired = FALSE, exact = FALSE)
      
      # Calculate effect size (rank biserial correlation) for Mann-Whitney U test
      effect_size <- (mw_test_result$statistic - (group1_size * group2_size / 2)) / sqrt(group1_size * group2_size * (group1_size + group2_size + 1) / 12)
      
      results <- rbind(results, data.frame(
        Variable = var,
        U_Value = mw_test_result$statistic,
        P_Value = mw_test_result$p.value,
        Effect_Size = effect_size,
        Group1_Size = group1_size,
        Group2_Size = group2_size
      ))
    }
  }
  
  return(results)
}

df_mannwhitney <- run_mann_whitney_tests(df_combined, "Source", "Human", "AI", ordinal_columns)


calculate_means_sds_medians_by_factors <- function(data, variables, factors) {
  # Create an empty list to store intermediate results
  results_list <- list()
  
  # Iterate over each variable
  for (var in variables) {
    # Create a temporary data frame to store results for this variable
    temp_results <- data.frame(Variable = var)
    
    # Iterate over each factor
    for (factor in factors) {
      # Aggregate data by factor: Mean, SD, Median
      agg_data <- aggregate(data[[var]], 
                            by = list(data[[factor]]), 
                            FUN = function(x) c(Mean = mean(x, na.rm = TRUE),
                                                SD = sd(x, na.rm = TRUE),
                                                Median = median(x, na.rm = TRUE)))
      
      # Create columns for each level of the factor
      for (level in unique(data[[factor]])) {
        level_agg_data <- agg_data[agg_data[, 1] == level, ]
        
        if (nrow(level_agg_data) > 0) {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- level_agg_data$x[1, "Mean"]
          temp_results[[paste0(factor, "_", level, "_SD")]] <- level_agg_data$x[1, "SD"]
          temp_results[[paste0(factor, "_", level, "_Median")]] <- level_agg_data$x[1, "Median"]
        } else {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- NA
          temp_results[[paste0(factor, "_", level, "_SD")]] <- NA
          temp_results[[paste0(factor, "_", level, "_Median")]] <- NA
        }
      }
    }
    
    # Add the results for this variable to the list
    results_list[[var]] <- temp_results
  }
  
  # Combine all the results into a single dataframe
  descriptive_stats_bygroup <- do.call(rbind, results_list)
  return(descriptive_stats_bygroup)
}

# Example usage
df_descriptive_stats_bygroup <- calculate_means_sds_medians_by_factors(df_combined, ordinal_columns, "Source")



calculate_correlation_matrix <- function(data, variables, method = "pearson") {
  # Ensure the method is valid
  if (!method %in% c("pearson", "spearman")) {
    stop("Method must be either 'pearson' or 'spearman'")
  }
  
  # Subset the data for the specified variables
  data_subset <- data[variables]
  
  # Calculate the correlation matrix
  corr_matrix <- cor(data_subset, method = method, use = "complete.obs")
  
  # Initialize a matrix to hold formatted correlation values
  corr_matrix_formatted <- matrix(nrow = nrow(corr_matrix), ncol = ncol(corr_matrix))
  rownames(corr_matrix_formatted) <- colnames(corr_matrix)
  colnames(corr_matrix_formatted) <- colnames(corr_matrix)
  
  # Calculate p-values and apply formatting with stars
  for (i in 1:nrow(corr_matrix)) {
    for (j in 1:ncol(corr_matrix)) {
      if (i == j) {  # Diagonal elements (correlation with itself)
        corr_matrix_formatted[i, j] <- format(round(corr_matrix[i, j], 3), nsmall = 3)
      } else {
        p_value <- cor.test(data_subset[[i]], data_subset[[j]], method = method)$p.value
        stars <- ifelse(p_value < 0.01, "***", 
                        ifelse(p_value < 0.05, "**", 
                               ifelse(p_value < 0.1, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}

df_correlation <- calculate_correlation_matrix(df_combined, ordinal_columns,"spearman")

# Create a new column 'Opinion' by merging responses
df_combined <- df_combined %>%
  mutate(Opinion = case_when(
    `Was.the.answer.written.by.a.human.endometriosis.health.expert.or.by.an.AI?.(1:.Definitely.written.by.a.human.2:.Likely.written.by.a.human.3:.Likely.written.by.an.AI.4:.Definitely.written.by.an.AI)` %in% c(1, 2) ~ "Human",
    `Was.the.answer.written.by.a.human.endometriosis.health.expert.or.by.an.AI?.(1:.Definitely.written.by.a.human.2:.Likely.written.by.a.human.3:.Likely.written.by.an.AI.4:.Definitely.written.by.an.AI)` %in% c(3, 4) ~ "AI",
    TRUE ~ NA_character_
  ))

# Convert 'Opinion' to a factor
df_combined$Opinion <- factor(df_combined$Opinion, levels = c("Human", "AI"))


# Add a new column 'Correctness' based on whether Source matches Opinion
df_combined <- df_combined %>%
  mutate(Correctness = if_else(Source == Opinion, "Correct", "Incorrect", missing = "Incorrect"))

# Convert 'Correctness' to a factor for easy analysis
df_combined$Correctness <- factor(df_combined$Correctness, levels = c("Correct", "Incorrect"))


# Chi Square for Correctedness

# Split data into separate dataframes for Human and AI sources
df_human <- df_combined %>% filter(Source == "Human")
df_ai <- df_combined %>% filter(Source == "AI")

# Define your row variables (categorical variables you want to test)
row_variables <- vars_cat2  # Ensure vars_cat2 is defined with the relevant column names
column_variable <- "Correctness"  # The new column indicating if the rater was correct


# Run chi-square analysis for Human source
df_chisquare_human <- chi_square_analysis_multiple(df_human, row_variables, column_variable)
print("Chi-Square Results for Human Source:")
print(df_chisquare_human)

# Run chi-square analysis for AI source
df_chisquare_ai <- chi_square_analysis_multiple(df_ai, row_variables, column_variable)
print("Chi-Square Results for AI Source:")
print(df_chisquare_ai)



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
  "Final Data" = df_combined,
  "Frequency Table" = df_freq, 
  "Frequency Table 2" = df_freq_segmented2,
  "Chi Square" = df_chisquare,
  "Correlations" = df_correlation,
  "Descriptives" = df_descriptive_stats_bygroup,
  "Mann Whitney" = df_mannwhitney,
  "ChiSq - AI" = df_chisquare_ai,
  "ChiSq - Human" = df_chisquare_human
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")
