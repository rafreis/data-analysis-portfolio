setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/sylvain")

library(readxl)
df <- read_xlsx('titre-DocumentStatistics (2) Excel.xlsx')
df_demographics <- read_xlsx('NUM - Statut Démographique_January 23_ 2023_12.28.xlsx')

# Assuming 'ID' is the common column to join on
merged_df <- merge(df, df_demographics, by = "Participant", all.x = TRUE)

# Get rid of special characters

names(merged_df) <- gsub(" ", "_", names(merged_df))
names(merged_df) <- gsub("\\(", "_", names(merged_df))
names(merged_df) <- gsub("\\)", "_", names(merged_df))
names(merged_df) <- gsub("\\-", "_", names(merged_df))
names(merged_df) <- gsub("/", "_", names(merged_df))
names(merged_df) <- gsub("\\\\", "_", names(merged_df)) 
names(merged_df) <- gsub("\\?", "", names(merged_df))
names(merged_df) <- gsub("\\'", "", names(merged_df))
names(merged_df) <- gsub("\\,", "_", names(merged_df))

# Generalized Mixed Model

library(lme4)
library(broom.mixed)
library(dplyr)

# Function to run GLMM
run_glmm <- function(data, outcome_vars, group_var, time_var, participant_var = NULL, covariates = NULL) {
  results_list <- list()
  
  for (outcome_var in outcome_vars) {
    # Construct the formula
    covariate_part <- if (!is.null(covariates) && length(covariates) > 0) paste("+", paste(covariates, collapse = " + ")) else ""
    random_effect_part <- if (!is.null(participant_var)) paste("+ (1|", participant_var, ")") else ""
    formula_str <- paste(outcome_var, "~", group_var, "*", time_var, covariate_part, random_effect_part)
    formula <- as.formula(formula_str)
    
    # Try fitting the model, print error if it occurs
    try({
      # Choose the appropriate model function based on the presence of random effects
      if (!is.null(participant_var)) {
        model <- glmer(formula, data = data, family = binomial(link = "logit"))
      } else {
        model <- glm(formula, data = data, family = binomial(link = "logit"))
      }
      
      # Get tidy results
      results <- broom.mixed::tidy(model)
      results$outcome_var <- outcome_var
      results_list[[outcome_var]] <- results
    }, silent = FALSE)  # Print error message if occurs
  }
  
  return(bind_rows(results_list))
}

# Define your outcome variables
outcome_vars <- c("Situation", "Task", "Action", "Result")

# Inspect Age percentiles

age_percentiles <- quantile(merged_df$How_old_are_you, probs = c(0.5, 0.80))

# Print the percentiles
print(age_percentiles)

categorize_age <- function(data, age_variable, cutoffs) {
  # Create a new variable for categorized age
  data$categorized_age <- cut(
    data[[age_variable]],
    breaks = c(-Inf, cutoffs, Inf), 
    labels = c("19-21", "21-27", "28_or_older"),
    include.lowest = TRUE
  )
  
  return(data)
}

age_cutoffs <- c(21, 27)
merged_df <- categorize_age(merged_df, "How_old_are_you", age_cutoffs)

# Create Dummy Variables
library(fastDummies)

vars_recode <- c("Gender_identity","In_which_faculty_do_you_study","categorized_age")

df_recoded <- dummy_cols(merged_df, select_columns = vars_recode)

names(df_recoded) <- gsub(" ", "_", names(df_recoded))
names(df_recoded) <- gsub("\\(", "_", names(df_recoded))
names(df_recoded) <- gsub("\\)", "_", names(df_recoded))
names(df_recoded) <- gsub("\\-", "_", names(df_recoded))
names(df_recoded) <- gsub("/", "_", names(df_recoded))
names(df_recoded) <- gsub("\\\\", "_", names(df_recoded)) 
names(df_recoded) <- gsub("\\?", "", names(df_recoded))
names(df_recoded) <- gsub("\\'", "", names(df_recoded))
names(df_recoded) <- gsub("\\,", "_", names(df_recoded))

# Sample Characterization

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

vars <- c("Gender_identity", "In_which_faculty_do_you_study", "categorized_age")

df_freqanalysis <- df_recoded %>%
  group_by(Participant) %>%
  sample_n(1)

freq_df <- create_frequency_tables(df_freqanalysis, vars)

# Split dataframes for each Question

subsets <- split(df_recoded, df_recoded$Question_type)
df_teamwork <- subsets$Teamwork
df_organization <- subsets$Organization
df_stress <- subsets$Stress

# Run the GLMM function
results_df_teamwork <- run_glmm(df_teamwork, outcome_vars, "Group", "Session", participant_var = NULL, covariates = NULL)
results_df_organization <- run_glmm(df_organization, outcome_vars, "Group", "Session", participant_var = NULL, covariates = NULL)
results_df_stress <- run_glmm(df_stress, outcome_vars, "Group", "Session", participant_var = NULL, covariates = NULL)


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
        
        # Add the variable name
        freq_table$Variable <- var
        
        # Calculate percentages relative to the total size of each secondary segment category
        for (seg in unique(freq_table$Secondary_Segment)) {
          seg_rows <- freq_table$Secondary_Segment == seg
          freq_table$Percentage[seg_rows] <- (freq_table$Count[seg_rows] / sum(freq_table$Count[seg_rows])) * 100
        }
        
        # Add the result to the list
        table_id <- paste(primary_segment, secondary_segment, var, sep = "_")
        all_freq_tables[[table_id]] <- freq_table
      }
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


segmented_freq_df_teamwork <- create_segmented_frequency_tables(df_teamwork, outcome_vars, "Session", "Group")
segmented_freq_df_organization <- create_segmented_frequency_tables(df_organization, outcome_vars, "Session", "Group")
segmented_freq_df_stress <- create_segmented_frequency_tables(df_stress, outcome_vars, "Session", "Group")


#PLots

library(ggplot2)

create_pre_post_comparison_plot <- function(data, primary_segment_col, secondary_segment_col, level_col, variable_col, percentage_col, legend_title = "Legend", plot_title = "Plot") {
  # Filter for the specified level
  data_filtered <- subset(data, data[[level_col]] == 1)
  
  # Create the plot
  ggplot(data_filtered, aes_string(x = primary_segment_col, y = percentage_col, fill = secondary_segment_col)) +
    geom_bar(stat = "identity", position = position_dodge()) +
    geom_text(aes(label = sprintf("%.1f%%", Percentage)), 
              position = position_dodge(width = 0.9), 
              vjust = -0.25, 
              size = 3) +
    facet_wrap(as.formula(paste("~", variable_col))) +
    labs(title = plot_title,
         x = "Session and Group",
         y = "Percentage") +
    scale_fill_brewer(palette = "Set1", name = legend_title) +
    theme_minimal()
}

plot <- create_pre_post_comparison_plot(segmented_freq_df_teamwork, "Secondary_Segment", "Primary_Segment", "Level", "Variable", "Percentage", "Session", "Percentage Comparison of Sessions for Teamwork")

# Display the plot
print(plot)

plot_org <- create_pre_post_comparison_plot(segmented_freq_df_organization, "Secondary_Segment", "Primary_Segment", "Level", "Variable", "Percentage", "Session", "Percentage Comparison of Sessions for Organization")

# Display the plot
print(plot_org)

plot_stress <- create_pre_post_comparison_plot(segmented_freq_df_stress, "Secondary_Segment", "Primary_Segment", "Level", "Variable", "Percentage", "Session", "Percentage Comparison of Sessions for Stress")

# Display the plot
print(plot_stress)


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
  "Sample Characterization" = freq_df,
  "Descriptives - Organization" = segmented_freq_df_organization,
  "Descriptives - Stress" = segmented_freq_df_stress,
  "Descriptives - Teamwork" = segmented_freq_df_teamwork,
  "Model - Organization" = results_df_organization, 
  "Model - Stress" = results_df_stress, 
  "Model - Teamwork" = results_df_teamwork
  )

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")