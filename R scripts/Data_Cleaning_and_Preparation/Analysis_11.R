setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/oliversogard/Segunda Ordem")

library(openxlsx)
df <- read.xlsx("5 yr Poly vs MBG vs. hybrid.xlsx")
library(tidyr)

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

demographics_categorical <- c("Gender", "Diabetes", "Smoker", "Anes_Type","deceased", "Complications",                  
                  "Revisions", "dislocations", "infection", "LOS_Days")

demographics_numerical <- c("Age" , "BMI", "Total_Op_Time")

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


df_freq <- create_frequency_tables(df, demographics_categorical)


library(dplyr)

df <- df %>%
  select(
    -Preoperative_internal_rotation,
    -`1year_internal_rotation`,  # Use backticks for column names that contain spaces or special characters
    -`5year_internal_rotation`
  )


df$SubjectID <- seq_len(nrow(df))

measures1 <- c("Preoperative_VAS"      ,              "Preoperative_ASES"    ,              
              "Preoperative_forward_elevation"     ,      
              "Preoperative_external_rotation"  ,   "Preoperative_rotation_90_abduction" ,
              "1year_VAS"                 ,         "1year_ASES"              ,          
              "1year_forward_elevation"   ,                   
              "1year_external_rotation" ,           "1year_rotation_90_abduction"   ,   
              "5year_VAS"                 ,         "5year_ASES"       ,                 
              "5year_forward_elevation"    ,                  
              "5year_external_rotation"    ,        "5year_rotation_90_abduction")

df_filtered <- df %>%
  select(SubjectID, all_of(measures1))

# Now pivot the filtered dataframe to long format
df_long <- df_filtered %>%
  pivot_longer(
    cols = -SubjectID, # Exclude the SubjectID column from the pivot
    names_to = "Measure_Time", # This will hold the combined TimePoint_Measure column names
    values_to = "Value" # This will hold the corresponding values
  ) %>%
  separate(Measure_Time, into = c("TimePoint", "Measure"), sep = "_", extra = "merge") %>%
  mutate(TimePoint = sub("^X", "", TimePoint)) # Use this line only if there are cases where TimePoint starts with "X"

# Paired Samples T-test

library(stats)
library(dplyr)
library(tidyr)
library(ggplot2)

run_paired_t_test <- function(data_long, measure, time_points) {
  data_filtered <- data_long %>%
    filter(Measure == measure, TimePoint %in% time_points) %>%
    drop_na(Value)
  
  if(all(table(data_filtered$TimePoint) >= 2)) {
    stats_summary <- data_filtered %>%
      group_by(TimePoint) %>%
      summarise(
        Mean = mean(Value, na.rm = TRUE),
        SD = sd(Value, na.rm = TRUE),
        N = n(),
        .groups = 'drop'
      ) %>%
      pivot_wider(names_from = TimePoint, values_from = c(Mean, SD, N))
    
    data_wide <- data_filtered %>%
      spread(key = TimePoint, value = Value)
    
    if(all(time_points %in% names(data_wide))) {
      t_test_results <- tryCatch({
        t.test(data_wide[[time_points[1]]], data_wide[[time_points[2]]], paired = TRUE)
      }, error = function(e) return(NULL))
      
      if(!is.null(t_test_results)) {
        total_mean <- mean(c(data_wide[[time_points[1]]], data_wide[[time_points[2]]]), na.rm = TRUE)
        total_sd <- sd(c(data_wide[[time_points[1]]], data_wide[[time_points[2]]]), na.rm = TRUE)
        
        stats_df <- data.frame(
          Measure = measure,
          SampleSize = sum(stats_summary$N_Preoperative, stats_summary$N_5year, na.rm = TRUE),
          tValue = t_test_results$statistic,
          pValue = t_test_results$p.value,
          TotalMean = total_mean,
          TotalSD = total_sd,
          Note = "Analysis successful"
        )
        
        for(tp in time_points) {
          stats_df[paste0(tp, "_Mean")] <- stats_summary[[paste0("Mean_", tp)]]
          stats_df[paste0(tp, "_SD")] <- stats_summary[[paste0("SD_", tp)]]
        }
        
        return(stats_df)
      } else {
        return(data.frame(Measure = measure, SampleSize = NA, tValue = NA, pValue = NA, TotalMean = NA, TotalSD = NA, Note = "t-test failed"))
      }
    } else {
      return(data.frame(Measure = measure, SampleSize = NA, tValue = NA, pValue = NA, TotalMean = NA, TotalSD = NA, Note = "Required time points missing"))
    }
  } else {
    return(data.frame(Measure = measure, SampleSize = NA, tValue = NA, pValue = NA, TotalMean = NA, TotalSD = NA, Note = "Insufficient data points for one or more time points"))
  }
}

# Specify the time points to compare
time_points <- c("Preoperative", "5year")

# Apply to all measures and bind the results together
measures <- c("VAS", "ASES", "forward_elevation", "external_rotation", "rotation_90_abduction")
results_list <- lapply(measures, function(m) run_paired_t_test(df_long, m, time_points))
df_results_ttest <- do.call(rbind, results_list)

# View the results
print(df_results_ttest)


# plot

# Assuming df_long has columns: SubjectID, Measure, TimePoint, Value
df_plot_data <- df_long %>%
  filter(TimePoint %in% c("Preoperative", "5year")) %>%
  group_by(Measure, TimePoint) %>%
  summarise(
    Mean = mean(Value, na.rm = TRUE),
    SD = sd(Value, na.rm = TRUE),
    N = n(),
    .groups = 'drop'
  ) %>%
  mutate(
    SE = SD / sqrt(N),
    CI = qt(0.975, df = N-1) * SE,
    LowerCI = Mean - CI,
    UpperCI = Mean + CI
  ) %>%
  mutate(TimePoint = factor(TimePoint, levels = c("Preoperative", "5year")))

ggplot(df_plot_data, aes(x = TimePoint, y = Mean, group = Measure, color = Measure)) +
  geom_line(aes(group = Measure), position = position_dodge(0.5)) +
  geom_point(position = position_dodge(0.5), size = 3) +
  geom_errorbar(aes(ymin = LowerCI, ymax = UpperCI), width = 0.2, position = position_dodge(0.5)) +
  facet_wrap(~Measure, scales = "free_y") +
  theme_minimal() +
  labs(title = "Change in Measures from Preoperative to 5 Years", y = "Mean Value with 95% CI", x = "") +
  theme(legend.position = "none", axis.text.x = element_text(angle = 45, hjust = 1))




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
  "Long-Shaped Data" = df_long1,
  "Frequency Table" = df_freq, 
  "T-test Results" = df_results_ttest 
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")


# Separately by Group

df_filtered <- df %>%
  select(SubjectID, Specific_glenoid_MBG_poly_or_hybrid, all_of(measures1))

library(tidyr)
library(dplyr)

# Assuming df_filtered includes columns for SubjectID, Specific_glenoid_MBG_poly_or_hybrid, and various Measure_TimePoint combinations
df_long1 <- df_filtered %>%
  pivot_longer(
    cols = -c(SubjectID, Specific_glenoid_MBG_poly_or_hybrid), # Exclude both SubjectID and Specific_glenoid_MBG_poly_or_hybrid from the pivot
    names_to = "Measure_Time", # This will hold the combined TimePoint_Measure column names
    values_to = "Value" # This will hold the corresponding values
  ) %>%
  separate(Measure_Time, into = c("TimePoint", "Measure"), sep = "_", extra = "merge") %>%
  mutate(TimePoint = sub("^X", "", TimePoint)) # Use this line only if there are cases where TimePoint starts with "X"



run_paired_t_test_for_glenoids <- function(data, specific_glenoid, measure) {
  data_filtered <- data %>%
    filter(Specific_glenoid_MBG_poly_or_hybrid == specific_glenoid, Measure == measure) %>%
    drop_na(Value)
  
  # Prepare data for both time points
  preop_data <- data_filtered %>% filter(TimePoint == "Preoperative") %>% pull(Value)
  postop_data <- data_filtered %>% filter(TimePoint == "5year") %>% pull(Value)
  
  # Check if there are at least two pairs of observations
  if(length(preop_data) < 2 || length(postop_data) < 2 || length(preop_data) != length(postop_data)) {
    # Insufficient data or unmatched pairs, so skip this test
    return(NULL)
  }
  
  # Perform the paired t-test
  t_test_result <- t.test(preop_data, postop_data, paired = TRUE)
  
  # Compile and return the results
  result <- data.frame(
    SpecificGlenoid = specific_glenoid,
    Measure = measure,
    tValue = t_test_result$statistic,
    pValue = t_test_result$p.value,
    SampleSize = length(preop_data),  # Since preop_data and postop_data are of equal length
    Preoperative_Mean = mean(preop_data, na.rm = TRUE),
    Preoperative_SD = sd(preop_data, na.rm = TRUE),
    FiveYear_Mean = mean(postop_data, na.rm = TRUE),
    FiveYear_SD = sd(postop_data, na.rm = TRUE)
  )
  
  return(result)
}



# Assuming df_long1 is your data frame
specific_glenoids <- unique(df_long1$Specific_glenoid_MBG_poly_or_hybrid)
measures <- unique(df_long1$Measure)

# Initialize an empty list to store results
results_list <- list()

# Loop through specific glenoids and measures
for(glenoid in specific_glenoids) {
  for(measure in measures) {
    test_result <- run_paired_t_test_for_glenoids(df_long1, glenoid, measure)
    if(!is.null(test_result)) {
      results_list[[paste(glenoid, measure, sep = "_")]] <- test_result
    }
  }
}

# Combine all results into a single dataframe
final_results_df <- do.call(rbind, results_list)

# View the final results
print(final_results_df)




for(specific_glenoid in specific_glenoids) {
  df_plot_data_specific <- df_plot_data %>%
    filter(Specific_glenoid_MBG_poly_or_hybrid == specific_glenoid)
  
  ggplot(df_plot_data_specific, aes(x = TimePoint, y = Mean, group = Measure, color = Measure)) +
    geom_line(aes(group = Measure), position = position_dodge(0.5)) +
    geom_point(position = position_dodge(0.5), size = 3) +
    geom_errorbar(aes(ymin = LowerCI, ymax = UpperCI), width = 0.2, position = position_dodge(0.5)) +
    facet_wrap(~Measure, scales = "free_y") +
    theme_minimal() +
    labs(title = paste("Change in Measures from Preoperative to 5 Years for", specific_glenoid), y = "Mean Value with 95% CI", x = "") +
    theme(legend.position = "none", axis.text.x = element_text(angle = 45, hjust = 1))
  
  # Display or save each plot here
}

