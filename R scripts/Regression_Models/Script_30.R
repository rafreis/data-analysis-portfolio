# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/heyimazier")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("human eval_ internal_vs_competitor.xlsx")


# Get rid of special characters

names(df) <- gsub(" ", "_", trimws(names(df)))
names(df) <- gsub("\\s+", "_", trimws(names(df), whitespace = "[\\h\\v\\s]+"))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))
names(df) <- gsub("\\$", "", names(df))
names(df) <- gsub("\\+", "_", names(df))

# Trim all values

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))


# Trim all values and convert to lowercase

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
    # Convert character strings to lowercase
    x <- tolower(x)
  }
  return(x)
}))

colnames(df)


# Outlier Evaluation

library(dplyr)
library(tidyr)

calculate_z_scores <- function(data, vars, id_var, z_threshold = 3) {
  # Prepare data by handling NA values and calculating Z-scores
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~ as.numeric(.))) %>%
    mutate(across(all_of(vars), ~ replace(., is.na(.), mean(., na.rm = TRUE)))) %>%
    mutate(across(all_of(vars), ~ (.-mean(., na.rm = TRUE))/ifelse(sd(.) == 0, 1, sd(., na.rm = TRUE)), .names = "z_{.col}"))
  
  # Include original raw scores and Z-scores
  z_score_data <- z_score_data %>%
    select(!!sym(id_var), all_of(vars), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  return(z_score_data)
}

df_zscores <- calculate_z_scores(df, "Which.model.is.more.helpful_.safe_.and.honest._rating_", "Prompt")


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

vars <- c("Prompt.Category", "Complexity")
df_freq <- create_frequency_tables(df, vars)


create_segmented_frequency_tables <- function(data, vars, segmenting_categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each segmenting category
  for (segment in segmenting_categories) {
    # Iterate over each variable for which frequencies are to be calculated
    for (var in vars) {
      # Ensure both the segmenting category and variable are treated as factors
      segment_factor <- factor(data[[segment]])
      var_factor <- factor(data[[var]])
      
      # Calculate counts segmented by the segmenting category
      counts <- table(segment_factor, var_factor)
      
      # Melt the table into a long format for easier handling
      freq_table <- as.data.frame(counts)
      names(freq_table) <- c("Segment", "Level", "Count")
      
      # Add the variable name
      freq_table$Variable <- var
      
      # Calculate percentages relative to the specific segment
      freq_table <- freq_table %>%
        group_by(Segment) %>%
        mutate(Percentage = (Count / sum(Count)) * 100) %>%
        ungroup()
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  combined_freq_table <- bind_rows(all_freq_tables, .id = "Combination")
  return(combined_freq_table)
}




df_freq_segmented <- create_segmented_frequency_tables(df, "Prompt.Category", "Complexity")





library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    Median = numeric(),
    SEM = numeric(),  # Standard Error of the Mean
    SD = numeric(),   # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    median_val <- median(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      Median = median_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

# Example usage with your dataframe and list of variables
df_descriptive_stats <- calculate_descriptive_stats(df, "Which.model.is.more.helpful_.safe_.and.honest._rating_")




## TWO-WAY ANOVA

perform_levenes_test <- function(data, variables, factors) {
  # Load necessary library
  if (!require(car)) install.packages("car")
  library(car)
  
  # Initialize an empty dataframe to store results
  levene_results <- data.frame(Variable = character(), 
                               Factor = character(),
                               F_Value = numeric(), 
                               DF1 = numeric(),
                               DF2 = numeric(),
                               P_Value = numeric(),
                               stringsAsFactors = FALSE)
  
  # Perform Levene's test for each variable and factor
  for (var in variables) {
    for (factor in factors) {
      # Perform Levene's Test
      test_result <- leveneTest(reformulate(factor, response = var), data = data)
      
      # Extract the F value, DF, and p-value
      F_value <- as.numeric(test_result[1, "F value"])
      DF1 <- as.numeric(test_result[1, "Df"])
      DF2 <- as.numeric(test_result[2, "Df"])
      P_value <- as.numeric(test_result[1, "Pr(>F)"])
      
      # Append the results to the dataframe
      levene_results <- rbind(levene_results, 
                              data.frame(Variable = var, 
                                         Factor = factor,
                                         F_Value = F_value, 
                                         DF1 = DF1,
                                         DF2 = DF2,
                                         P_Value = P_value))
    }
  }
  
  return(levene_results)
}

levene_test_results <- perform_levenes_test(df, "Which.model.is.more.helpful_.safe_.and.honest._rating_", vars)



# Function to perform two-way ANOVA
perform_two_way_anova <- function(data, response_vars, factors) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  
  # Iterate over response variables
  for (var in response_vars) {
    # Two-way ANOVA
    anova_model <- aov(reformulate(paste(factors[1], factors[2], sep = "*"), response = var), data = data)
    model_summary <- summary(anova_model)
    
    # Extract relevant statistics
    model_results <- data.frame(
      Variable = var,
      Effect = rownames(model_summary[[1]]),
      Sum_Sq = model_summary[[1]][["Sum Sq"]],
      Mean_Sq = model_summary[[1]][["Mean Sq"]],
      Df = model_summary[[1]][["Df"]],
      FValue = model_summary[[1]][["F value"]],
      pValue = model_summary[[1]][["Pr(>F)"]]
    )
    
    anova_results <- rbind(anova_results, model_results)
  }
  
  return(anova_results)
}

# Perform two-way ANOVA
two_way_anova_results <- perform_two_way_anova(df, "Which.model.is.more.helpful_.safe_.and.honest._rating_", vars)


## POST-HOC TESTS

library(multcomp)
library(emmeans)

# Function to perform appropriate post-hoc test for multiple response variables and factors
perform_posthoc_tests <- function(data, response_vars, factors, equal_variance) {
  results_list <- list() # Initialize a list to store results
  
  # Iterate over each response variable
  for (response_var in response_vars) {
    # Iterate over each factor
    for (factor in factors) {
      # Fit the model
      model <- aov(reformulate(factor, response = response_var), data = data)
      
      # If equal variance is assumed, perform Tukey's post-hoc test
      if (equal_variance) {
        posthoc_result <- TukeyHSD(model)
      } else {
        # Else perform Dunnett-T3 test
        # Calculate estimated marginal means (emmeans)
        emm <- emmeans(model, specs = factor)
        
        # Perform Dunnett-T3 test
        posthoc_result <- cld(emm, adjust = "dunnett")
      }
      
      # Store the results in the list with a unique name
      result_name <- paste0("Posthoc_", response_var, "_", factor)
      results_list[[result_name]] <- posthoc_result
    }
  }
  
  return(results_list)
}


posthoc_results <- perform_posthoc_tests(df, "Which.model.is.more.helpful_.safe_.and.honest._rating_", vars, equal_variance = TRUE)

# Function to integrate post-hoc test results into a data frame
integrate_posthoc_results <- function(posthoc_results) {
  # Initialize an empty data frame to store combined results
  combined_results <- data.frame()
  
  # Iterate over each item in the post-hoc results list
  for (result_name in names(posthoc_results)) {
    # Extract the test results
    test_results <- posthoc_results[[result_name]]
    
    # Iterate over each element in the test results
    for (factor_name in names(test_results)) {
      # Retrieve the result matrix and convert it to a data frame
      result_matrix <- test_results[[factor_name]]
      result_df <- as.data.frame(result_matrix)
      
      # Add columns for the response variable and factor
      result_df$Response_Variable <- sub("Posthoc_", "", result_name)
      result_df$Factor <- factor_name
      result_df$Comparison <- rownames(result_df)
      
      # Bind the result to the combined data frame
      combined_results <- rbind(result_df, combined_results)
    }
  }
  
  # Reset row names
  rownames(combined_results) <- NULL
  
  return(combined_results)
}

posthoc_results <- integrate_posthoc_results(posthoc_results)

## INTERACTION PLOTS

library(ggplot2)

# Function to create interaction plots with means for each factor level
create_mean_interaction_plots <- function(data, response_vars, factor1, factor2) {
  plot_list <- list()
  
  for (response_var in response_vars) {
    # Calculate means for each combination of factor levels
    data_means <- aggregate(data[[response_var]], by = list(data[[factor1]], data[[factor2]]), FUN = mean, na.rm = TRUE)
    names(data_means) <- c(factor1, factor2, "Mean")
    
    # Convert factors to factor type if they are not already
    data_means[[factor1]] <- as.factor(data_means[[factor1]])
    data_means[[factor2]] <- as.factor(data_means[[factor2]])
    
    # Create the interaction plot
    plot_title <- paste("Interaction Plot:", response_var, "by", factor1, "and", factor2)
    plot <- ggplot(data_means, aes_string(x = factor1, y = "Mean", group = factor2, color = factor2)) +
      geom_line() +
      geom_point() +
      theme_minimal() +
      labs(title = plot_title,
           x = factor1,
           y = response_var) +
      theme(legend.title = element_blank(),
            axis.text.x = element_text(angle = 45, hjust = 1))  # Rotate x-axis labels by 45 degrees
    
    plot_list[[plot_title]] <- plot
  }
  
  return(plot_list)
}

# Print and save each plot
print_and_save_plots <- function(plot_list) {
  for (plot_title in names(plot_list)) {
    plot <- plot_list[[plot_title]]
    print(plot) # Print the plot
    
    # Save the plot
    file_name <- paste0(gsub("[ :]", "_", plot_title), ".png") # Replace spaces and colons with underscores in file name
    ggsave(file_name, plot, width = 8, height = 6)
  }
}


interaction_plots <- create_mean_interaction_plots(df, "Which.model.is.more.helpful_.safe_.and.honest._rating_","Prompt.Category", "Complexity")
print_and_save_plots(interaction_plots)

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
  "Frequency Table" = df_freq, 
  "Seg. Frequency Table" = df_freq_segmented, 
  "Descriptive Stats" = df_descriptive_stats, 
  "Levenes Tests" = levene_test_results,
  "Two-way ANOVA" =two_way_anova_results,
  "Post-hoc" = posthoc_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")



# Grouping and summarizing data to calculate mean and standard deviation
library(dplyr)

# Replace 'Prompt.Category' and 'Complexity' with the actual column names in your dataset
mean_sd_table <- df %>%
  group_by(Prompt.Category, Complexity) %>%
  summarise(
    Mean = mean(Which.model.is.more.helpful_.safe_.and.honest._rating_, na.rm = TRUE),
    SD = sd(Which.model.is.more.helpful_.safe_.and.honest._rating_, na.rm = TRUE),
    Count = n()
  ) %>%
  ungroup()

# Save the table to a CSV file for reporting
write.csv(mean_sd_table, "Mean_SD_Table.csv", row.names = FALSE)

library(ggplot2)

# Bar chart with 95% confidence intervals
bar_chart_ci <- ggplot(mean_sd_table, aes(x = Prompt.Category, y = Mean, fill = Complexity)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.9)) +
  geom_errorbar(
    aes(ymin = Mean - (1.96 * SD / sqrt(Count)), ymax = Mean + (1.96 * SD / sqrt(Count))),
    width = 0.2,
    position = position_dodge(width = 0.9)
  ) +
  labs(
    title = "Mean Ratings by Prompt Category and Complexity (95% CI)",
    x = "Prompt Category",
    y = "Mean Rating"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) # Rotate x-axis labels

# Print the chart
print(bar_chart_ci)

# Save the chart as a PNG file
ggsave("Mean_Ratings_Bar_Chart_95CI.png", bar_chart_ci, width = 10, height = 6)


# Enhanced heatmap with more distinct colors and labels
heatmap_plot_enhanced <- ggplot(mean_sd_table, aes(x = Prompt.Category, y = Complexity, fill = Mean)) +
  geom_tile(color = "white") +
  geom_text(aes(label = round(Mean, 2)), color = "black", size = 4) + # Add data labels
  scale_fill_gradientn(
    colors = c("blue", "yellow", "darkgreen"),
    values = scales::rescale(c(1, 4, 7)), # Rescale for more contrast
    limits = c(1, 7),
    name = "Mean Rating"
  ) +
  labs(
    title = "Enhanced Heatmap of Mean Ratings by Prompt Category and Complexity",
    x = "Prompt Category",
    y = "Complexity"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1), # Rotate x-axis labels
    panel.grid = element_blank()
  )

# Print the enhanced heatmap
print(heatmap_plot_enhanced)

# Save the enhanced heatmap as a PNG file
ggsave("Enhanced_Heatmap_Mean_Ratings.png", heatmap_plot_enhanced, width = 12, height = 6)



# Load necessary libraries
library(ggplot2)
library(dplyr)

# Load necessary libraries
library(ggplot2)
library(dplyr)

# Exclude the first row from the dataframe
posthoc_results <- posthoc_results %>%
  slice(-1)

# Order the entire dataset by differences
posthoc_results_ordered <- posthoc_results %>%
  arrange(desc(diff)) %>%
  mutate(Comparison = reorder(Comparison, diff)) # Reorder y-axis by differences

# Split the ordered dataset into two parts
posthoc_results_part1 <- posthoc_results_ordered %>%
  slice(1:(nrow(posthoc_results_ordered) / 2))

posthoc_results_part2 <- posthoc_results_ordered %>%
  slice((nrow(posthoc_results_ordered) / 2 + 1):nrow(posthoc_results_ordered))


# Heatmap for Part 1
heatmap_part1 <- ggplot(posthoc_results_part1, aes(x = Factor, y = Comparison, fill = diff)) +
  geom_tile(color = "white") +
  geom_text(aes(label = ifelse(Significant == "Yes", 
                               paste0(round(diff, 2), "*"), 
                               round(diff, 2))),
            color = "black", size = 4) +
  scale_fill_gradient2(
    low = "blue", mid = "white", high = "red", midpoint = 0,
    name = "Difference (diff)"
  ) +
  labs(
    title = "Post-Hoc Results Heatmap - Part 1 (Ordered)",
    x = "Factor",
    y = "Comparison (Ordered by Differences)"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(hjust = 1),
    axis.text.y = element_text(size = 10),
    panel.grid = element_blank()
  )

# Heatmap for Part 2
heatmap_part2 <- ggplot(posthoc_results_part2, aes(x = Factor, y = Comparison, fill = diff)) +
  geom_tile(color = "white") +
  geom_text(aes(label = ifelse(Significant == "Yes", 
                               paste0(round(diff, 2), "*"), 
                               round(diff, 2))),
            color = "black", size = 4) +
  scale_fill_gradient2(
    low = "blue", mid = "white", high = "red", midpoint = 0,
    name = "Difference (diff)"
  ) +
  labs(
    title = "Post-Hoc Results Heatmap - Part 2 (Ordered)",
    x = "Factor",
    y = "Comparison (Ordered by Differences)"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(hjust = 1),
    axis.text.y = element_text(size = 10),
    panel.grid = element_blank()
  )

# Print and save the plots
print(heatmap_part1)
ggsave("PostHoc_Results_Heatmap_Part1_Ordered.png", heatmap_part1, width = 12, height = 8)

print(heatmap_part2)
ggsave("PostHoc_Results_Heatmap_Part2_Ordered.png", heatmap_part2, width = 12, height = 8)
