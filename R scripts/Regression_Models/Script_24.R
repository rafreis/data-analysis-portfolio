# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elizabethwoodl")

# Read data from a CSV file
df <- read.csv("fully_filtered_data.csv")

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

# Define lists of relevant variables
vars_AIEDsupport <- c("Q12_2", "Q5_1","Q5_2", "Q5_3",  "Q5_4" , "Q5_5", "Q5_6","Q5_7")

vars_Regulationsupport <- c("restict_importance", "govt_ranks_1"  , "govt_ranks_2", "govt_ranks_3", "govt_ranks_4")

vars_familiarity <- c("familiarity_AI","frequency_AI")

vars_comparison <- c("state", "educ", "pid5")


# Recode variables to numeric to allow for statistical tests

library(dplyr)
# Recode Q12_1 and Q12_2
df$Q12_1 <- recode(df$Q12_1,
                   "Neutral" = 3,
                   "Somewhat support" = 4,
                   "Strongly support" = 5,
                   "Somewhat oppose" = 2,
                   "Strongly oppose" = 1,
                   "-99" = NA_real_,
                   .default = NA_real_)

df$Q12_2 <- recode(df$Q12_2,
                   "Neutral" = 3,
                   "Somewhat support" = 4,
                   "Strongly support" = 5,
                   "Somewhat oppose" = 2,
                   "Strongly oppose" = 1,
                   "-99" = NA_real_,
                   .default = NA_real_)

# Recode Q5_1 to Q5_7
q5_vars <- c("Q5_1", "Q5_2", "Q5_3", "Q5_4", "Q5_5", "Q5_6", "Q5_7")

df[q5_vars] <- lapply(df[q5_vars], function(x) recode(x,
                                                      "Somewhat positive" = 3,
                                                      "Extremely positive" = 4,
                                                      "Extremely negative" = 1,
                                                      "Somewhat negative" = 2,
                                                      "-99" = NA_real_,
                                                      .default = NA_real_))



# Recode familiarity_AI
df$familiarity_AI <- recode(df$familiarity_AI,
                            "Not familiar at all (e.g., never used or heard about AI tools like ChatGPT or Grammarly)" = 1,
                            "Slightly familiar" = 2,
                            "Moderately familiar" = 3,
                            "Very familiar" = 4,
                            "Extremely familiar" = 5,
                            "-99" = NA_real_,
                            .default = NA_real_)

# Recode frequency_AI
df$frequency_AI <- recode(df$frequency_AI,
                          "I have never used AI tools" = 1,
                          "Less than once a month" = 2,
                          "Once or twice a month" = 3,
                          "Once a week (e.g., 4-5 times a month)" = 4,
                          "Twice a week (e.g., 8-10 times a month)" = 5,
                          "Almost every day" = 6,
                          "Multiple times a day" = 7,
                          "-99" = NA_real_,
                          .default = NA_real_)

# Recode regulation support
df$restict_importance <- recode(df$restict_importance,
                                "Not at all important" = 1,
                                "Slightly important" = 2,
                                "Moderately important" = 3,
                                "Very important" = 4,
                                "Extremely important" = 5,
                                "-99" = NA_real_,
                                .default = NA_real_)

govt_vars <- c("govt_ranks_1", "govt_ranks_2", "govt_ranks_3", "govt_ranks_4")

df[govt_vars] <- lapply(df[govt_vars], function(x) recode(x,
                                                          "Somewhat Agree" = 3,
                                                          "Agree" = 4,
                                                          "Somewhat Disagree" = 2,
                                                          "Disagree" = 1,
                                                          "-99" = NA_real_,
                                                          .default = NA_real_))


# Create Frequency Tables to inspect dataset

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

# For Numeric vars
df_freq_numeric <- create_frequency_tables(df, c(vars_AIEDsupport, vars_familiarity, vars_Regulationsupport))

# For categorical vars
df_freq_cat <- create_frequency_tables(df, vars_comparison)


# Recoding States with minimal representation in the data (< 10 rows)

# List of states to recode as 'Other'
states_to_recode <- c("Alaska", "Delaware", "District of Columbia", "Hawaii", 
                      "Idaho", "Montana", "Nebraska", "New Mexico", 
                      "North Dakota", "South Dakota", "Vermont", "Wyoming", "Maine", "New Hampshire", "Rhode Island", "Utah")

# Manually recode df$state
df$state <- ifelse(df$state %in% states_to_recode, "Other", df$state)


# Rescaling variables to allow integration

# Rescale Q12_2 from 1-5 to 1-4
df$Q12_2 <- scales::rescale(df$Q12_2, to = c(1, 4), from = c(1, 5))

# Rescale frequency_AI from 1-7 to 1-5
df$frequency_AI <- scales::rescale(df$frequency_AI, to = c(1, 5), from = c(1, 7))

# Rescale restict_importance from 1-5 to 1-4
df$restict_importance <- scales::rescale(df$restict_importance, to = c(1, 4), from = c(1, 5))


# New Frequency table
# For categorical vars
df_freq_cat2 <- create_frequency_tables(df, vars_comparison)


## Reliability Analysis

# Scale Construction
library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Added for item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val <- alpha_results$total$raw_alpha
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"] # Get ITC for the item
      item_mean <- mean(item_data, na.rm = TRUE)
      item_sem <- sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data)))
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = item_mean,
        SEM = item_sem,
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
        Alpha = NA
      ))
    }
    
    # Calculate the mean score for the scale and add as a new column
    scale_mean <- rowMeans(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_mean
    scale_mean_overall <- mean(scale_mean, na.rm = TRUE)
    scale_sem <- sd(scale_mean, na.rm = TRUE) / sqrt(sum(!is.na(scale_mean)))
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = scale_mean_overall,
      SEM = scale_sem,
      StDev = sd(scale_mean, na.rm = TRUE),
      ITC = NA,  # ITC is not applicable for the total scale
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}

scales <- list(
  "AI_Ed_support" = c("Q12_2", "Q5_1","Q5_2", "Q5_3",  "Q5_4" , "Q5_5", "Q5_6","Q5_7"),
  "regulation_support" = c("restict_importance", "govt_ranks_1"  , "govt_ranks_2", "govt_ranks_3", "govt_ranks_4"),
  "familiarity_score" = c("familiarity_AI","frequency_AI")
)

alpha_results <- reliability_analysis(df, scales)

df_recoded <- alpha_results$data_with_scales
df_descriptives <- alpha_results$statistics

colnames(df_recoded)


# Normality Assessment

library(e1071) 

calculate_stats <- function(data, variables) {
  results <- data.frame(Variable = character(),
                        Skewness = numeric(),
                        Kurtosis = numeric(),
                        Shapiro_Wilk_F = numeric(),
                        Shapiro_Wilk_p_value = numeric(),
                        KS_Statistic = numeric(),
                        KS_p_value = numeric(),
                        stringsAsFactors = FALSE)
  
  for (var in variables) {
    if (var %in% names(data)) {
      # Calculate skewness and kurtosis
      skew <- skewness(data[[var]], na.rm = TRUE)
      kurt <- kurtosis(data[[var]], na.rm = TRUE)
      
      # Perform Shapiro-Wilk test
      shapiro_test <- shapiro.test(data[[var]])
      
      # Perform Kolmogorov-Smirnov test
      ks_test <- ks.test(data[[var]], "pnorm", mean = mean(data[[var]], na.rm = TRUE), sd = sd(data[[var]], na.rm = TRUE))
      
      # Add results to the dataframe
      results <- rbind(results, c(var, skew, kurt, shapiro_test$statistic, shapiro_test$p.value, ks_test$statistic, ks_test$p.value))
    } else {
      warning(paste("Variable", var, "not found in the data. Skipping."))
    }
  }
  
  colnames(results) <- c("Variable", "Skewness", "Kurtosis", "Shapiro_Wilk_F", "Shapiro_Wilk_p_value", "KS_Statistic", "KS_p_value")
  return(results)
}

# Example usage with your data
df_normality_results <- calculate_stats(df_recoded, c("AI_Ed_support" ,"regulation_support", "familiarity_score"))


# Generate Boxplots for a visual inspection of variable's distributions
library(ggplot2)
library(tidyr)
library(stringr)

create_boxplots <- function(df, vars) {
  # Ensure that vars are in the dataframe
  df <- df[, vars, drop = FALSE]
  
  # Reshape the data to a long format
  long_df <- df %>%
    gather(key = "Variable", value = "Value")
  
  # Create side-by-side boxplots for each variable
  p <- ggplot(long_df, aes(x = Variable, y = Value)) +
    geom_boxplot() +
    stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = "Boxplots", x = "Variable", y = "Value") +
    theme(axis.text.x = element_text(hjust = 1, vjust = 1)) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 10))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = "side_by_side_boxplots.png", plot = p, width = 10, height = 6)
}

create_boxplots(df_recoded, c("AI_Ed_support" ,"regulation_support", "familiarity_score"))

# Employing ANOVA to test the relationship between categorical vars (variable pairs 2 to 5 on the freelance brief doc)

# Function to perform one-way ANOVA
perform_one_way_anova <- function(data, response_vars, factors) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  
  # Iterate over response variables
  for (var in response_vars) {
    # Iterate over factors
    for (factor in factors) {
      # One-way ANOVA
      anova_model <- aov(reformulate(factor, response = var), data = data)
      model_summary <- summary(anova_model)
      
      # Extract relevant statistics
      model_results <- data.frame(
        Variable = var,
        Effect = factor,
        Sum_Sq = model_summary[[1]]$"Sum Sq"[1],  # Extract Sum Sq for the factor
        Mean_Sq = model_summary[[1]]$"Mean Sq"[1],  # Extract Mean Sq for the factor
        Df = model_summary[[1]]$Df[1],  # Extract Df for the factor
        FValue = model_summary[[1]]$"F value"[1],  # Extract F value for the factor
        pValue = model_summary[[1]]$"Pr(>F)"[1]  # Extract p-value for the factor
      )
      
      anova_results <- rbind(anova_results, model_results)
    }
  }
  
  return(anova_results)
}


# Perform ANOVA
anova_results <- perform_one_way_anova(
  data = df_recoded,
  response_vars = c("regulation_support", "AI_Ed_support", "familiarity_score"),
  factors = c("state", "educ", "pid5")
)

# Display results
print("ANOVA Results:")
print(anova_results)

# Function to calculate correlation matrix
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

# Calculate correlation
correlation_matrix <- calculate_correlation_matrix(
  data = df_recoded,
  variables = c("AI_Ed_support", "regulation_support"),
  method = "pearson"
)

print("Correlation Matrix:")
print(correlation_matrix)


# Crosstabulation to understand differences in scores across subgroups

calculate_means_and_sds_by_factors <- function(data, variables, factors) {
  # Create an empty list to store intermediate results
  results_list <- list()
  
  # Iterate over each variable
  for (var in variables) {
    # Create a temporary data frame to store results for this variable
    temp_results <- data.frame(Variable = var)
    
    # Iterate over each factor
    for (factor in factors) {
      # Aggregate data by factor
      agg_data <- aggregate(data[[var]], by = list(data[[factor]]), 
                            FUN = function(x) c(Mean = mean(x, na.rm = TRUE), SD = sd(x, na.rm = TRUE)))
      
      # Create columns for each level of the factor
      for (level in unique(data[[factor]])) {
        level_agg_data <- agg_data[agg_data[, 1] == level, ]
        
        if (nrow(level_agg_data) > 0) {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- level_agg_data$x[1, "Mean"]
          temp_results[[paste0(factor, "_", level, "_SD")]] <- level_agg_data$x[1, "SD"]
        } else {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- NA
          temp_results[[paste0(factor, "_", level, "_SD")]] <- NA
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

factors <- vars_comparison
vars = c("regulation_support", "AI_Ed_support", "familiarity_score")

descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df_recoded, vars, factors)


print(descriptive_stats_bygroup)

# Comparative Bar Plots

library(ggplot2)
library(dplyr)
library(tidyr)
library(stringr)

generate_comparative_bar_plot <- function(df, vars, factor_column, factor_levels = NULL, font_size = 12, output_filename = "comparative_bar_plot.png") {
  # Convert the specified columns to numeric
  df[vars] <- lapply(df[vars], function(x) as.numeric(as.character(x)))
  
  # Reshape data to long format, excluding the factor column
  long_df <- df %>%
    pivot_longer(cols = all_of(vars), names_to = "variable", values_to = "value") %>%
    drop_na(value)
  
  # Automatically set factor levels if not provided
  if (is.null(factor_levels)) {
    factor_levels <- sort(unique(df[[factor_column]]))
  }
  long_df[[factor_column]] <- factor(long_df[[factor_column]], levels = factor_levels)
  
  # Generate the plot
  p <- ggplot(long_df, aes(x = variable, y = value, fill = !!sym(factor_column))) +
    geom_bar(stat = "summary", fun = "mean", position = position_dodge()) +
    geom_text(aes(label = sprintf("%.3f", ..y..)), stat = "summary", fun = "mean", position = position_dodge(width = 0.9), vjust = -0.3, size = 3.5) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 20)) + # Wrap labels for long names
    scale_fill_brewer(palette = "Pastel1", name = factor_column) + # Use a color palette for the factor levels
    theme_minimal(base_size = font_size) +
    theme(axis.text.x = element_text(angle = 0, hjust = 0.5, size = font_size), # Centered X-axis text
          axis.text.y = element_text(size = font_size)) +
    labs(y = "Mean", x = "") +
    ggtitle("Comparative Bar Plot") +
    theme(plot.background = element_rect(fill = "white"),
          legend.position = "bottom")
  
  # Print the plot to display
  print(p)
  
  # Save the plot
  ggsave(filename = output_filename, plot = p, width = 10, height = 6, bg = "white")
}

# Example usage
# Assuming 'df_recoded' is your dataset
factors <- vars_comparison
vars = c("regulation_support", "AI_Ed_support", "familiarity_score")

descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df_recoded, vars, factors)

# Generate comparative bar plot example
factor_col <- "educ"
factor_levels <- NULL

generate_comparative_bar_plot(df_recoded, vars, factor_col, factor_levels, font_size = 9, output_filename = "/mnt/data/comparative_bar_plot.png")

factor_col <- "pid5"
generate_comparative_bar_plot(df_recoded, vars, factor_col, factor_levels, font_size = 9, output_filename = "/mnt/data/comparative_bar_plot.png")


# Function to generate horizontal column plots
# Factor labels on Y-axis, mean scores on X-axis, ordered alphabetically
generate_horizontal_column_plot <- function(df, var, factor_column, font_size = 12, output_filename = "horizontal_column_plot.png") {
  # Convert the variable to numeric
  df[[var]] <- as.numeric(as.character(df[[var]]))
  
  # Automatically factorize and order the factor column alphabetically
  df[[factor_column]] <- factor(df[[factor_column]], levels = sort(unique(df[[factor_column]])))
  
  # Calculate mean scores by factor levels
  summary_df <- df %>%
    group_by(!!sym(factor_column)) %>%
    summarise(mean_score = mean(!!sym(var), na.rm = TRUE))
  
  # Generate the horizontal column plot
  p <- ggplot(summary_df, aes(x = mean_score, y = !!sym(factor_column))) +
    geom_col(fill = "steelblue") +
    geom_text(aes(label = sprintf("%.2f", mean_score)), hjust = -0.2, size = 3.5) +
    theme_minimal(base_size = font_size) +
    theme(
      axis.text.y = element_text(size = font_size),
      axis.text.x = element_text(size = font_size),
      panel.grid.major.y = element_blank(),
      panel.grid.minor = element_blank()
    ) +
    labs(x = "Mean Score", y = factor_column, title = paste("Mean Scores for", var))
  
  # Adjust plot limits to accommodate text labels
  p <- p + expand_limits(x = max(summary_df$mean_score) * 1.2)
  
  # Print the plot to display
  print(p)
  
  # Save the plot
  ggsave(filename = output_filename, plot = p, width = 8, height = 6, bg = "white")
}

var <- "regulation_support"
factor_column <- "state"

# Call the function to generate the plot
generate_horizontal_column_plot(df_recoded, var, factor_column, font_size = 12, output_filename = "horizontal_column_plot.png")

var <- "AI_Ed_support"
# Call the function to generate the plot
generate_horizontal_column_plot(df_recoded, var, factor_column, font_size = 12, output_filename = "horizontal_column_plot.png")

var <-  "familiarity_score"
# Call the function to generate the plot
generate_horizontal_column_plot(df_recoded, var, factor_column, font_size = 12, output_filename = "horizontal_column_plot.png")


## Export Results to Excel

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
  "ANOVA REsults" = anova_results, 
  "Correlations" = correlation_matrix,
  "Reliability" = df_descriptives,
  "Frequency Table" = df_freq_cat,
  "Frequency Table2" = df_freq_cat2,
  "Frequency Table3" = df_freq_numeric,
  "Normality Results" = df_normality_results,
  "Descriptives by Group" = descriptive_stats_bygroup
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")