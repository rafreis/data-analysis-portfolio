# Get unique time points
time_points <- unique(df$Timepoint)
# Define all possible columns
all_colnames <- c("CutoffAge", "AgeGroup", "Timepoint", "GroupPair", "DV", "WStatistic", "PValue",
                  "MedianGroup1_T0", "MedianGroup1_T42", "MedianGroup1_T84",
                  "MedianGroup2_T0", "MedianGroup2_T42", "MedianGroup2_T84")

# Initialize an empty data frame with these columns
results_mannwhitney_df <- data.frame(matrix(ncol = length(all_colnames), nrow = 0))
colnames(results_mannwhitney_df) <- all_colnames
                                                          
# Loop through the specified age cutoff points
for (cutoff_age in c(48, 50, 52, 54, 56, 58, 60)) {
  
  # Loop through age groups ("All Ages", "Cutoff Age or Older", and "Younger")
  for (age_group in c("All Ages", "Cutoff Age or Older", "Younger")) {
    
    if (age_group == "All Ages") {
      df_filtered <- df
    } else {
      df_filtered <- df %>% filter(ifelse(Age >= cutoff_age, "Cutoff Age or Older", "Younger") == age_group)
    }
    
    # Loop through each dependent variable
    for (dv in dvs) {
      
      # Loop through unique time points
      for (timepoint in time_points) {
        
        # Filter data based on the current timepoint
        df_filtered_timepoint <- df_filtered %>% filter(Timepoint == timepoint)
        
        # Perform Mann-Whitney test for all possible pairs of product groups
        product_groups <- unique(df_filtered_timepoint$Group.Name)
        
        for (k in 1:(length(product_groups) - 1)) {
          for (l in (k + 1):length(product_groups)) {
            
            # Get the current product group pair
            group_pair <- paste(product_groups[k], product_groups[l], sep = "_")
            
            # Filter data for each product group
            data_group1 <- df_filtered_timepoint %>% filter(Group.Name == product_groups[k])
            data_group2 <- df_filtered_timepoint %>% filter(Group.Name == product_groups[l])
            
            # Calculate medians
            median_group1 <- median(data_group1[[dv]], na.rm = TRUE)
            median_group2 <- median(data_group2[[dv]], na.rm = TRUE)
            
            # Perform the Mann-Whitney test
            result <- wilcox.test(data_group1[[dv]], data_group2[[dv]], exact = FALSE)
            
            # Inside your loop:
            result_row <- data.frame(matrix(ncol = length(all_colnames), nrow = 1))
            colnames(result_row) <- all_colnames
            
            # Populate known columns
            result_row$CutoffAge <- cutoff_age
            result_row$AgeGroup <- age_group
            result_row$Timepoint <- timepoint
            result_row$GroupPair <- group_pair
            result_row$DV <- dv
            result_row$WStatistic <- result$statistic
            result_row$PValue <- result$p.value
            
            # Dynamically populate median columns based on timepoint
            result_row[[paste("MedianGroup1_", timepoint, sep = "")]] <- median_group1
            result_row[[paste("MedianGroup2_", timepoint, sep = "")]] <- median_group2
            
            # Append to the main results data frame
            results_mannwhitney_df <- rbind(results_mannwhitney_df, result_row)
          }
        }
      }
      
      }
    }
}

# Load necessary libraries
library(tidyverse)
library(writexl)

# Reshape the dataframe
reshaped_mannwhitneydf <- results_mannwhitney_df %>%
  pivot_wider(names_from = Timepoint, values_from = c(WStatistic, PValue))

# Export the reshaped dataframe to an Excel file
write_xlsx(reshaped_mannwhitneydf, "Mann_Whitney_Results.xlsx")

## MEDIAN DIFFERENCES

library(dplyr)
library(tidyr)
library(writexl)

# Initialize a new dataframe to store results
results_diff_mannwhitney_df <- data.frame()

# Loop through the specified age cutoff points
age_cutoffs <- c("All Ages", 48, 50, 52, 54, 56, 58, 60)

for (cutoff_age in age_cutoffs) {
  
  # Loop through age groups ("All Ages", "Cutoff Age or Older", and "Younger")
  for (age_group in c("All Ages", "Cutoff Age or Older", "Younger")) {
    
    # Filter data based on age group and cutoff, or include all ages
    df_filtered <- if (age_group == "All Ages" || cutoff_age == "All Ages") {
      df
    } else {
      df %>% filter(ifelse(Age >= as.numeric(cutoff_age), "Cutoff Age or Older", "Younger") == age_group)
    }
    
    # Loop through each dependent variable
    for (dv in dvs) {
      
      # Calculate differences for each subject and for each group
      df_diff <- df_filtered %>%
        select(Subject, Timepoint, Group.Name, all_of(dv)) %>%
        spread(Timepoint, all_of(dv)) %>%
        mutate(diff_T0_T42 = `T42` - `T0`,
               diff_T42_T84 = `T84` - `T42`,
               diff_T0_T84 = `T84` - `T0`)
      
      # Loop through the timepoint differences
      for (time_diff in c("diff_T0_T42", "diff_T42_T84", "diff_T0_T84")) {
        
        # Loop through unique pairs of product groups
        product_groups <- unique(df_filtered$Group.Name)
        for (k in 1:(length(product_groups) - 1)) {
          for (l in (k + 1):length(product_groups)) {
            
            # Get the current product group pair
            group_pair <- paste(product_groups[k], product_groups[l], sep = "_")
            
            # Extract the differences for the current product group pair
            data_group1_diff <- df_diff %>% filter(Group.Name == product_groups[k]) %>% pull(all_of(time_diff))
            data_group2_diff <- df_diff %>% filter(Group.Name == product_groups[l]) %>% pull(all_of(time_diff))
            
            # Run Mann-Whitney test on these differences
            result <- wilcox.test(data_group1_diff, data_group2_diff, exact = FALSE)
            
            # Store results
            result_row <- data.frame(
              CutoffAge = cutoff_age,
              AgeGroup = age_group,
              TimepointDifference = time_diff,
              DV = dv,
              GroupPair = group_pair,
              WStatistic = result$statistic,
              PValue = result$p.value
            )
            results_diff_mannwhitney_df <- rbind(results_diff_mannwhitney_df, result_row)
          }
        }
      }
    }
  }
}

library(tidyr)

# Pivot the dataframe to make time differences columns
pivoted_results_df <- results_diff_mannwhitney_df %>%
  pivot_wider(names_from = TimepointDifference, values_from = c(WStatistic, PValue))

# Export the results to an Excel file
write_xlsx(pivoted_results_df, "Mann_Whitney_Time_Differences.xlsx")
