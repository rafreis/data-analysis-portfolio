library(stats)
library(dplyr)

run_repeated_measures_anova <- function(data_long, measure) {
  # Filter the dataset for the specific measure and remove rows with NA in 'Value'
  data_filtered <- data_long %>%
    filter(Measure == measure) %>%
    drop_na(Value)
  
  # Ensure there are sufficient data points across time points before proceeding
  if(length(unique(data_filtered$TimePoint)) < 2) {
    return(data.frame(Measure = measure, SampleSize = NA, Df = NA, Sum_Sq = NA, Mean_Sq = NA, FValue = NA, pValue = NA, Note = "Insufficient data points across time points"))
  }
  
  # Automatically infer the order of TimePoint from the data
  unique_time_points <- unique(data_filtered$TimePoint)
  data_filtered$TimePoint <- factor(data_filtered$TimePoint, levels = unique_time_points, ordered = TRUE)
  
  # Count the number of non-NA observations for the current measure
  sample_size <- nrow(data_filtered)
  
  # Run the repeated measures ANOVA using aov
  aov_results <- tryCatch({
    aov(Value ~ TimePoint + Error(SubjectID/TimePoint), data = data_filtered)
  }, error = function(e) return(NULL))
  
  if(is.null(aov_results)) {
    return(data.frame(Measure = measure, SampleSize = sample_size, Df = NA, Sum_Sq = NA, Mean_Sq = NA, FValue = NA, pValue = NA, Note = "ANOVA failed"))
  }
  
  # Summarize the ANOVA results
  summary_df <- summary(aov_results)[[1]]
  
  # Extract the relevant statistics
  stats_df <- data.frame(
    Measure = measure,
    SampleSize = sample_size,
    Df = ifelse(!is.null(summary_df), summary_df$"Df"[1], NA),
    Sum_Sq = ifelse(!is.null(summary_df), summary_df$"Sum Sq"[1], NA),
    Mean_Sq = ifelse(!is.null(summary_df), summary_df$"Mean Sq"[1], NA),
    FValue = ifelse(!is.null(summary_df), summary_df$"F value"[1], NA),
    pValue = ifelse(!is.null(summary_df), summary_df$"Pr(>F)"[1], NA),
    Note = ifelse(!is.null(summary_df), "Analysis successful", "Insufficient data for analysis")
  )
  
  return(stats_df)
}



vars <- c("Satisfaction_HP_Access"                                ,
          "Satisfaction_HP_Waiting_Time"                          ,
          "Satisfaction_HP_Proximity"                             ,
          "Satisfaction_HP_Competence"                            ,
          "Satisfaction_HP_Empathy"                               ,
          "Satisfaction_HP_Relief"                                ,
          "Satisfaction_HP_Cost"                                  ,
          "Satisfaction_HPP_Access"                               ,
          "Satisfaction_HPP_Waiting_Time"                         ,
          "Satisfaction_HPP_Proximity"                            ,
          "Satisfaction_HPP_Competence"                           ,
          "Satisfaction_HPP_Empathy"                              ,
          "Satisfaction_HPP_Relief"                               ,
          "Satisfaction_HPP_Cost"                                 ,
          "Satisfaction_Psychologe_HPP_Access"                    ,
          "Satisfaction_Psychologe_HPP_Waiting_Time"              ,
          "Satisfaction_Psychologe_HPP_Proximity"                 ,
          "Satisfaction_Psychologe_HPP_Competence"                ,
          "Satisfaction_Psychologe_HPP_Empathy"                   ,
          "Satisfaction_Psychologe_HPP_Relief"                    ,
          "Satisfaction_Psychologe_HPP_Cost"                      ,
          "Satisfaction_Psychologischer_PT_Access"                ,
          "Satisfaction_Psychologischer_PT_Waiting_Time"          ,
          "Satisfaction_Psychologischer_PT_Proximity"             ,
          "Satisfaction_Psychologischer_PT_Competence"            ,
          "Satisfaction_Psychologischer_PT_Empathy"               ,
          "Satisfaction_Psychologischer_PT_Relief"                ,
          "Satisfaction_Psychologischer_PT_Cost"                  ,
          "Satisfaction_Aerztlicher_PT_Access"                    ,
          "Satisfaction_Aerztlicher_PT_Waiting_Time"              ,
          "Satisfaction_Aerztlicher_PT_Proximity"                 ,
          "Satisfaction_Aerztlicher_PT_Competence"                ,
          "Satisfaction_Aerztlicher_PT_Empathy"                   ,
          "Satisfaction_Aerztlicher_PT_Relief"                    ,
          "Satisfaction_Aerztlicher_PT_Cost"                      ,
          "Satisfaction_HP_Overall"                               ,
          "Satisfaction_HPP_Overall"                              ,
          "Satisfaction_Psychologe_HPP_Overall"                   ,
          "Satisfaction_Psychologischer_PT_Overall"               ,
          "Satisfaction_Aerztlicher_PT_Overall"                   ,
          "Professional_Group_HP_Access"                          ,
          "Professional_Group_Psychologischer_PT_Access"          ,
          "Professional_Group_Aerztlicher_PT_Access"              ,
          "Professional_Group_HP_Waiting_Time"                    ,
          "Professional_Group_Psychologischer_PT_Waiting_Time"    ,
          "Professional_Group_Aerztlicher_PT_Waiting_Time"        ,
          "Professional_Group_HP_Proximity"                       ,
          "Professional_Group_Psychologischer_PT_Proximity"       ,
          "Professional_Group_Aerztlicher_PT_Proximity"           ,
          "Professional_Group_HP_Qualified"                       ,
          "Professional_Group_Psychologischer_PT_Qualified"       ,
          "Professional_Group_Aerztlicher_PT_Qualified"           ,
          "Professional_Group_HP_Competence"                      ,
          "Professional_Group_Psychologischer_PT_Competence"      ,
          "Professional_Group_Aerztlicher_PT_Competence"          ,
          "Professional_Group_HP_Empathy"                         ,
          "Professional_Group_Psychologischer_PT_Empathy"         ,
          "Professional_Group_Aerztlicher_PT_Empathy"             ,
          "Professional_Group_HP_Relief"                          ,
          "Professional_Group_Psychologischer_PT_Relief"          ,
          "Professional_Group_Aerztlicher_PT_Relief"              ,
          "Professional_Group_HP_Cost"                            ,
          "Professional_Group_Psychologischer_PT_Cost"            ,
          "Professional_Group_Aerztlicher_PT_Cost")
  

# Extract specific measures by dynamically splitting the variables based on common measure names
measure_types <- unique(gsub(".*_(.*)", "\\1", vars))  # Extract unique measure types from variable names

# Initialize lists to hold variables for each measure type
measure_vars_lists <- setNames(vector("list", length(measure_types)), measure_types)

# Fill lists with variables corresponding to each measure type
for (measure in measure_types) {
  measure_vars_lists[[measure]] <- grep(paste0("_", measure, "$"), vars, value = TRUE)
}

library(stats)
library(stats)
library(dplyr)

# Function to run ANOVA directly on each list of variables
run_direct_anova <- function(df, var_list) {
  # Filter df to include only the columns in var_list
  response_var <- df[, var_list, drop = FALSE]  # Ensure 'drop = FALSE' to keep df structure
  
  # Remove rows with any NA values across the specified columns
  complete_cases <- response_var[complete.cases(response_var), ]
  
  # Stack the complete cases data
  stacked_data <- stack(complete_cases, na.rm = FALSE)  # Stack data including NA values, though there should be none
  
  # Calculate N as the number of unique non-NA entries in the original data
  estimated_participants = nrow(complete_cases)
  
  if (nrow(stacked_data) > 0) {  # Check that there are rows left after removing incomplete cases
    # Run ANOVA only on non-NA data (which should be all data at this point)
    anova_result <- aov(values ~ ind, data = stacked_data)
    summary_aov <- summary(anova_result)
    return(list(Summary=summary_aov, N=estimated_participants))  # Return the count of complete cases
  } else {
    return(list(Summary="No valid data available after removing incomplete cases", N=0))
  }
}



# Function to extract the ANOVA summary and include the sample size
extract_anova_summary <- function(anova_summary, N) {
  # First, check if anova_summary$Summary is not a list (indicating that the aov function may have failed)
  if (!is.list(anova_summary$Summary)) {
    return(data.frame(
      Term = NA,
      Df = NA,
      SumSq = NA,
      MeanSq = NA,
      FValue = NA,
      PValue = NA,
      Note = "ANOVA summary is not a list, likely due to failed computation.",
      SampleSize = N
    ))
  }
  
  # Now check if the first element of the list contains the expected ANOVA table
  if (is.null(anova_summary$Summary[[1]]) || !("Df" %in% names(anova_summary$Summary[[1]]))) {
    return(data.frame(
      Term = NA,
      Df = NA,
      SumSq = NA,
      MeanSq = NA,
      FValue = NA,
      PValue = NA,
      Note = "ANOVA table missing or incomplete in the summary.",
      SampleSize = N
    ))
  }
  
  # If the above checks pass, extract details from the summary
  anova_table <- anova_summary$Summary[[1]]
  data.frame(
    Term = rownames(anova_table),
    Df = anova_table$Df,
    SumSq = anova_table$`Sum Sq`,
    MeanSq = anova_table$`Mean Sq`,
    FValue = anova_table$`F value`,
    PValue = anova_table$`Pr(>F)`,
    Note = "Analysis successful",
    SampleSize = N  # Include the sample size in the output
  )
}


# Iterate over each measure type in your measure_vars_lists and run ANOVA
anova_results_prof <- list()
for (measure in names(professional_group_lists)) {
  var_list <- professional_group_lists[[measure]]
  
  # Ensure all columns exist in df and contain only numeric values
  if (all(var_list %in% names(df)) && all(sapply(df[, var_list, drop = FALSE], is.numeric))) {
    anova_results_prof[[measure]] <- run_direct_anova(df, var_list)
  } else {
    anova_results_prof[[measure]] <- "Invalid data: Non-numeric values or missing columns"
  }
}

# Print or process results
print(anova_results_prof)

# Iterate over each measure type in your measure_vars_lists and run ANOVA
anova_results_sat <- list()
for (measure in names(satisfaction_lists)) {
  var_list <- satisfaction_lists[[measure]]
  
  # Ensure all columns exist in df and contain only numeric values
  if (all(var_list %in% names(df)) && all(sapply(df[, var_list, drop = FALSE], is.numeric))) {
    anova_results_sat[[measure]] <- run_direct_anova(df, var_list)
  } else {
    anova_results_sat[[measure]] <- "Invalid data: Non-numeric values or missing columns"
  }
}

# Print or process results
print(anova_results_sat)

professional_results_df <- lapply(anova_results_prof, function(x) extract_anova_summary(x$Summary, x$N))
professional_results_df <- do.call(rbind, professional_results_df)
print(professional_results_df)

satisfaction_results_df <- lapply(anova_results_sat, function(x) extract_anova_summary(x$Summary, x$N))
satisfaction_results_df <- do.call(rbind, satisfaction_results_df)
print(satisfaction_results_df)

# Plots

library(ggplot2)
library(dplyr)

# Function to calculate mean, SD, SE, CI, and plot for given columns
plot_measure_results <- function(df, var_list, measure_name) {
  # Prepare the data for plotting
  data_for_plot <- df %>%
    select(all_of(var_list)) %>%
    pivot_longer(cols = everything(), names_to = "Group", values_to = "Value") %>%
    na.omit() %>%
    group_by(Group) %>%
    summarise(
      Mean = mean(Value),
      SD = sd(Value),
      N = n(),
      .groups = 'drop'
    ) %>%
    mutate(
      SE = SD / sqrt(N),
      CI = qt(0.975, df = N - 1) * SE
    )
  
  # Plotting
  ggplot(data_for_plot, aes(x = Group, y = Mean, fill = Group)) +
    geom_bar(stat = "identity", position = position_dodge(), width = 0.7) +
    geom_errorbar(aes(ymin = Mean - CI, ymax = Mean + CI), width = 0.2, position = position_dodge(0.7)) +
    labs(title = paste("Mean Scores with 95% CI for", measure_name), x = "Group", y = "Mean Score") +
    theme_minimal()
}

# Example of plotting for one measure
# Assuming 'Satisfaction_HP_Access' and other related columns are in df
plot_access <- plot_measure_results(df, professional_group_lists$Access, "Professional Group Access")
plot_access




library(moments)

# Function to calculate descriptive statistics including 95% CI for the mean
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(),  # Standard Error of the Mean
    SD = numeric(),   # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    CI_Lower = numeric(),  # Lower bound of the 95% CI for the mean
    CI_Upper = numeric(),  # Upper bound of the 95% CI for the mean
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    non_na_data <- na.omit(variable_data)  # Remove NA values
    n <- length(non_na_data)  # Sample size
    mean_val <- mean(non_na_data, na.rm = TRUE)
    sem_val <- sd(non_na_data, na.rm = TRUE) / sqrt(n)
    sd_val <- sd(non_na_data, na.rm = TRUE)
    skewness_val <- skewness(non_na_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(non_na_data, na.rm = TRUE)
    
    # Calculate the 95% CI for the mean
    t_critical <- qt(0.975, df = n - 1)  # 0.975 for 95% CI, two-tailed
    margin_error <- t_critical * sem_val
    ci_lower <- mean_val - margin_error
    ci_upper <- mean_val + margin_error
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val,
      CI_Lower = ci_lower,
      CI_Upper = ci_upper
    ))
  }
  
  return(results)
}

# Assuming 'df' is your dataframe and 'vars' are the columns you want to analyze
df_descriptive_stats <- calculate_descriptive_stats(df, vars)











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
  "Descriptives" = df_descriptive_stats,
  "Professional Group" = professional_results_df,
  "Satisfaction" = satisfaction_results_df
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables3.xlsx")




library(ggplot2)
library(dplyr)

# Function to plot and save results with corrected group names and updated aesthetics
plot_and_save_measure_results <- function(df, var_list, measure_name, file_name) {
  # Prepare the data for plotting by ensuring the Group column is correctly formatted
  data_for_plot <- df %>%
    select(all_of(var_list)) %>%
    pivot_longer(cols = everything(), names_to = "Group", values_to = "Value") %>%
    mutate(Group = gsub("Professional_Group_HP_", "HP", Group),
           Group = gsub("Professional_Group_Psychologischer_PT_", "Psychol. PT", Group),
           Group = gsub("Professional_Group_Aerztlicher_PT_", "Aerztlicher PT", Group),
           Group = gsub("_Access", "", Group),  # Remove '_Access' suffix from group names
           Group = gsub("_Waiting_Time", "", Group),  # Similarly remove other suffixes if necessary
           Group = gsub("_Proximity", "", Group),
           Group = gsub("_Competence", "", Group),
           Group = gsub("_Empathy", "", Group),
           Group = gsub("_Relief", "", Group),
           Group = gsub("_Cost", "", Group)) %>%
    drop_na(Value) %>%
    group_by(Group) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      N = n(),
      .groups = 'drop'
    ) %>%
    mutate(
      SE = SD / sqrt(N),
      CI = qt(0.975, df = N - 1) * SE
    )
  
  # Check if data_for_plot is empty
  if (nrow(data_for_plot) == 0) {
    stop("No data available for plotting. Check group transformations.")
  }
  
  # Plotting
  p <- ggplot(data_for_plot, aes(x = Group, y = Mean, fill = Group)) +
    geom_bar(stat = "identity", position = position_dodge(), width = 0.7) +
    geom_errorbar(aes(ymin = Mean - CI, ymax = Mean + CI), width = 0.2, position = position_dodge(0.7)) +
    geom_text(aes(label = sprintf("%.2f", Mean)), vjust = -1.5, position = position_dodge(0.7)) +
    labs(title = paste("Mean Scores with 95% CI for", measure_name), x = "Profession", y = "Mean Score") +
    theme_minimal() +
    scale_fill_discrete(name = "Profession")
  
  # Save the plot
  ggsave(file_name, plot = p, width = 10, height = 6)
  
  return(p)
}

# Save plots for all measures in professional_group_lists and satisfaction_lists
save_all_plots <- function(lists, df, prefix) {
  lapply(names(lists), function(measure) {
    plot_name <- paste0(prefix, measure, ".png")
    plot_and_save_measure_results(df, lists[[measure]], paste(prefix, measure), plot_name)
  })
}

# Run the function for both professional and satisfaction measures
save_all_plots(professional_group_lists, df, "Professional_Group_")
save_all_plots(satisfaction_lists, df, "Satisfaction_")
