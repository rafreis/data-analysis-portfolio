# Load the libraries
library(readxl)
library(tidyverse)

# Define the file path
file_path <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/franksalas/Data Collection PhD.xlsx"

# List all sheet names
sheet_names <- excel_sheets(file_path)

# Function to process numeric columns with comma-separated values (pre/post)
split_pre_post <- function(col) {
  # Split the column based on commas
  values <- str_split(col, ",")
  
  # If values have pre and post, split them accordingly
  pre_post <- map_dfr(values, ~{
    if (length(.) == 2) {
      tibble(pre = as.numeric(.[1]), post = as.numeric(.[2]))
    } else {
      tibble(pre = NA, post = as.numeric(.[1])) # Handles single value case
    }
  })
  
  return(pre_post)
}

# Function to read each sheet and process pre/post columns
process_sheet <- function(sheet_name) {
  df <- read_excel(file_path, sheet = sheet_name)
  
  # Select numeric columns that need splitting (ignoring 'ID' and other non-numeric columns)
  df_processed <- df %>%
    mutate(across(-ID, split_pre_post)) %>%
    unnest(cols = everything(), names_sep = "_") %>%
    mutate(Condition = sheet_name)
  
  return(df_processed)
}

# Apply the function to all sheets and combine into one dataframe
combined_df <- map_df(sheet_names, process_sheet)

# Rename the 'Workload_post' column to 'Workload' and remove 'Workload_pre'
combined_df <- combined_df %>%
  rename(Workload = Workload_post, Athlete_BW = Athlete_BW_post) %>%
  select(-Workload_pre, -Athlete_BW_pre)

str(combined_df)



# Load necessary libraries
library(tidyverse)
library(afex)
library(ggplot2)
library(psych)

# List of variable pairs (pre and post measures)
variable_pairs <- list(
  Jump_height = c("Jump_height_pre", "Jump_height_post"),
  Countermovement_depth = c("Countermovement_depth_pre", "Countermovement_depth_post"),
  Peak_braking_force = c("Peak_braking_force_pre", "Peak_braking_force_post"),
  Peak_propulsive_force = c("Peak_propulsive_force_pre", "Peak_propulsive_force_post"),
  Braking_phase = c("Braking_phase_pre", "Braking_phase_post"),
  Propulsive_phase = c("Propulsive_phase_pre", "Propulsive_phase_post"),
  Flight_time = c("Flight_time_pre", "Flight_time_post"),
  Time_to_take_off = c("Time_to_take_off_pre", "Time_to_take_off_post"),
  Take_off_velocity = c("Take_off_velocity_pre", "Take_off_velocity_post"),
  Peak_landing_force = c("Peak_landing_force_pre", "Peak_landing_force_post"),
  RSI = c("RSI_pre", "RSI_post"),
  mRSI = c("mRSI_pre", "mRSI_post")
)


# Load necessary libraries
library(tidyverse)
library(stats)
library(afex)

# Prepare long format data
combined_long_df <- combined_df %>%
  pivot_longer(
    cols = c(Jump_height_pre, Jump_height_post,
             Countermovement_depth_pre, Countermovement_depth_post,
             Peak_braking_force_pre, Peak_braking_force_post,
             Peak_propulsive_force_pre, Peak_propulsive_force_post,
             Braking_phase_pre, Braking_phase_post,
             Propulsive_phase_pre, Propulsive_phase_post,
             Flight_time_pre, Flight_time_post,
             Time_to_take_off_pre, Time_to_take_off_post,
             Take_off_velocity_pre, Take_off_velocity_post,
             Peak_landing_force_pre, Peak_landing_force_post,
             RSI_pre, RSI_post,
             mRSI_pre, mRSI_post),
    names_to = c("Measure", "Time"),
    names_pattern = "(.*)_(.*)$"
  ) %>%
  mutate(
    Time = factor(Time, levels = c("pre", "post")),
    Condition = factor(Condition),
    SubjectID = factor(ID)
  ) %>%
  rename(Value = value) %>%
  drop_na(Value) # Remove rows with NA in 'Value'

# List of measure names
measure_names <- c("Jump_height", "Countermovement_depth", "Peak_braking_force", 
                   "Peak_propulsive_force", "Braking_phase", "Propulsive_phase",
                   "Flight_time", "Time_to_take_off", "Take_off_velocity",
                   "Peak_landing_force", "RSI", "mRSI")




# Outlier Evaluation

library(dplyr)

calculate_z_scores <- function(data, vars, id_var, leader_col, condition_col, time_col, z_threshold = 4) {
  z_score_data <- data %>%
    group_by(!!sym(leader_col)) %>%
    mutate(across(all_of(vars), ~ (.-mean(.))/sd(.), .names = "z_{.col}")) %>%
    ungroup() %>%
    select(!!sym(id_var), !!sym(leader_col), !!sym(condition_col), !!sym(time_col), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  return(z_score_data)
}

# Assuming 'combined_long_df' has columns 'Condition' and 'Time' you want to preserve
condition_col <- "Condition"
time_col <- "Time"

# Calling the modified function
df_zscores <- calculate_z_scores(combined_long_df, c("Value"), "SubjectID", "Measure", condition_col, time_col)


# Define the outliers in a dataframe
outliers <- data.frame(
  SubjectID = c("#7", "#6", "#6", "#5", "#11"),
  Measure = c("Peak_braking_force", "Take_off_velocity", "Take_off_velocity", "Take_off_velocity", "Time_to_take_off"),
  Condition = c("Week 6 - 130%BW", "Week 6 - 130%BW", "Week 4 - 90%BW", "Week 4 - 90%BW", "Week 3 - 70%BW"),
  Time = c("pre", "post", "post", "post", "post"),
  stringsAsFactors = FALSE  # Ensure that strings are not converted to factors
)

# Filter out the outliers from the combined_long_df
combined_long_df_nooutliers <- combined_long_df %>%
  anti_join(outliers, by = c("SubjectID", "Measure", "Condition", "Time"))



## Segregando resultados por uma coluna de fator

library(e1071)

calculate_stats_by_group <- function(data, variables, group_var) {
  if (!group_var %in% names(data)) {
    stop(paste("Group variable", group_var, "not found in the data."))
  }
  
  results <- data.frame(Group = character(),
                        Variable = character(),
                        Skewness = numeric(),
                        Kurtosis = numeric(),
                        Shapiro_Wilk_F = numeric(),
                        Shapiro_Wilk_p_value = numeric(),
                        stringsAsFactors = FALSE)
  
  for (level in unique(data[[group_var]])) {
    data_subset <- data[data[[group_var]] == level, ]
    
    for (var in variables) {
      if (var %in% names(data)) {
        # Extract the variable data, excluding NA values
        var_data <- na.omit(data_subset[[var]])
        
        # Calculate skewness and kurtosis
        skew <- skewness(var_data, na.rm = TRUE)
        kurt <- kurtosis(var_data, na.rm = TRUE)
        
        # Perform Shapiro-Wilk test only if there are enough non-NA data points
        if (length(var_data) > 3) {
          shapiro_test <- shapiro.test(var_data)
          shapiro_F <- shapiro_test$statistic
          shapiro_p <- shapiro_test$p.value
        } else {
          shapiro_F <- NA
          shapiro_p <- NA
        }
        
        # Add results to the dataframe
        results <- rbind(results, c(level, var, skew, kurt, shapiro_F, shapiro_p))
      } else {
        warning(paste("Variable", var, "not found in the data. Skipping."))
      }
    }
  }
  
  colnames(results) <- c("Group", "Variable", "Skewness", "Kurtosis", "Shapiro_Wilk_F", "Shapiro_Wilk_p_value")
  return(results)
}



# Example usage with your data
df_normality_results <- calculate_stats_by_group(combined_long_df_nooutliers, "Value", "Measure")


library(nlme)  # Load the nlme package for handling more complex models

perform_two_way_rm_anova_per_measure <- function(data, factors, subject) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  unique_measures <- unique(data$Measure)  # Get unique measures
  
  # Set Week 1 - 30%BW as the reference category
  data$Condition <- factor(data$Condition, levels = c("Week1- 30%BW", "Week 2 - 50%BW", "Week 3 - 70%BW", 
                                                      "Week 4 - 90%BW", "Week 5 - 110%BW", "Week 6 - 130%BW", 
                                                      "Week 7 - MAXISO"))
  
  # Iterate over each unique measure
  for (measure in unique_measures) {
    measure_data <- subset(data, Measure == measure)  # Filter data for the current measure
    
    # Fit the model with repeated measures
    rm_anova_model <- lme(Value ~ Condition * Time, random = ~1 | SubjectID, 
                          data = measure_data, method = "REML")
    
    # Summary of the model
    model_summary <- summary(rm_anova_model)
    
    # Extract relevant statistics from the model for fixed effects
    model_results <- data.frame(
      Measure = measure,
      Effect = rownames(model_summary$tTable),
      Estimate = model_summary$tTable[, "Value"],
      Std_Error = model_summary$tTable[, "Std.Error"],
      DF = model_summary$tTable[, "DF"],
      tValue = model_summary$tTable[, "t-value"],
      pValue = model_summary$tTable[, "p-value"]
    )
    
    # Combine results into the results dataframe
    anova_results <- rbind(anova_results, model_results)
  }
  
  return(anova_results)
}

# Perform two-way repeated measures ANOVA
factors <- c("Condition", "Time")  # Factors in your study
subject <- "SubjectID"  # Column that identifies each subject

two_way_rm_anova_results <- perform_two_way_rm_anova_per_measure(combined_long_df_nooutliers, factors, subject)


library(ggplot2)
library(dplyr)
library(tidyr)
library(forcats)  # For factor manipulation

# Function to generate interaction plots for each measure
generate_interaction_plots <- function(df, measure_col, id_var, condition_col, time_col, output_dir) {
  # Correct and order the levels of Condition
  df[[condition_col]] <- factor(df[[condition_col]],
                                levels = c("Week1- 30%BW", "Week 2 - 50%BW", "Week 3 - 70%BW",
                                           "Week 4 - 90%BW", "Week 5 - 110%BW", "Week 6 - 130%BW",
                                           "Week 7 - MAXISO"))
  
  # Loop through each unique measure
  unique_measures <- unique(df[[measure_col]])
  
  for (measure in unique_measures) {
    # Filter data for the current measure
    measure_data <- df %>% 
      filter(!!sym(measure_col) == measure) %>%
      group_by(!!sym(condition_col), !!sym(time_col)) %>%
      summarise(mean_value = mean(Value, na.rm = TRUE), 
                se = sd(Value, na.rm = TRUE) / sqrt(n()), 
                .groups = 'drop')
    
    # Generate the plot
    p <- ggplot(measure_data, aes(x = !!sym(condition_col), y = mean_value, group = !!sym(time_col), color = !!sym(time_col), fill = !!sym(time_col))) +
      geom_line(position = position_dodge(0.1)) +
      geom_point(position = position_dodge(0.1), size = 3) +
      geom_errorbar(aes(ymin = mean_value - se, ymax = mean_value + se), width = 0.2, position = position_dodge(0.1)) +
      labs(title = paste("Interaction Plot for", measure), x = condition_col, y = "Mean Value") +
      theme_minimal() +
      scale_color_brewer(palette = "Set1") +
      scale_fill_brewer(palette = "Set1")
    
    # Print the plot to display
    print(p)
    
    # Save the plot
    ggsave(filename = paste0(output_dir, "/", gsub("[ -]", "_", measure), "_interaction_plot.png"), plot = p, width = 12, height = 8, bg = "white")
  }
}

# Example call to the function
output_dir <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/franksalas"  # Define your output directory
generate_interaction_plots(combined_long_df_nooutliers, "Measure", "SubjectID", "Condition", "Time", output_dir)

print(levels(combined_long_df$Condition))


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
  "Long-Shaped Data" = combined_long_df, 
  "Normality" = df_normality_results, 
  "Outliers" = df_zscores,
  "ANOVA Results" = two_way_rm_anova_results
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")
