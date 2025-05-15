setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/kestrel23")

library(readxl)
df <- read_xlsx('Input_Data.xlsx')


## Convert numeric columns and string columns

vars <- c("I_feel_free_to_make_choices_with_regards_to_the_way_I_train","I_feel_pushed_to_behave_in_certain_ways","I_have_a_say_in_how_things_are_done","feel_forced_to_follow_training_decisions","I_have_the_freedom_to_make_training_decisions","I_feel_forced_to_do_training_tasks_that_I_would_not_choose_to_do",
        "I_pursue_goals_that_are_my_own","I_feel_excessive_pressure","I_feel_like_I_can_be_myself\r\n","I_must_do_what_I_am_told","AS","AF")

# Convert specified variables to numeric
for (var in vars) {
  df[[var]] <- as.numeric(as.character(df[[var]]))
}

# Get rid of spaces in column names

names(df) <- gsub(" ", "_", names(df))

## DESCRIPTIVE TABLES

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

descriptive_stats <- calculate_descriptive_stats(df, vars)

## RELIABILITY ANALYSIS

library (psych)

# Function to calculate Cronbach's Alpha with custom scale names
calculate_cronbach_alpha <- function(data, scales) {
  results <- data.frame(
    ScaleName = character(),
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  for (scale in names(scales)) {
    # Subset the data for the current set of variables
    subset_data <- data[scales[[scale]]]
    # Calculate Cronbach's Alpha
    alpha_val <- alpha(subset_data)$total$raw_alpha
    # Append the results
    results <- rbind(results, data.frame(
      ScaleName = scale,
      Alpha = alpha_val
    ))
  }
  
  return(results)
}

scales <- list("Autonomy_Satisfaction" = c("I_feel_free_to_make_choices_with_regards_to_the_way_I_train","I_have_a_say_in_how_things_are_done","I_have_the_freedom_to_make_training_decisions",
                                        "I_pursue_goals_that_are_my_own","I_feel_like_I_can_be_myself\r\n"),"Autonomy_Frustration" = c("I_feel_pushed_to_behave_in_certain_ways", "feel_forced_to_follow_training_decisions","I_feel_forced_to_do_training_tasks_that_I_would_not_choose_to_do","I_feel_excessive_pressure","I_must_do_what_I_am_told"))
alpha_results <- calculate_cronbach_alpha(df, scales)

## DESCRIPTIVE STATS BY FACTORS

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

factors <- c("Adulthood Group", "Coached")
calculate_means_and_sds_by_factors(df, vars, factors)
descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df, vars, factors)

## CORRELATION ANALYSIS

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
        stars <- ifelse(p_value < 0.001, "***", 
                        ifelse(p_value < 0.01, "**", 
                               ifelse(p_value < 0.05, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}


correlation_matrix <- calculate_correlation_matrix(df, vars, method = "pearson")

## TWO-WAY ANOVA

library(car)
library(multcomp)
library(nparcomp)


# Levenes Test

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

response_vars <- c("AS", "AF")
factors <- c("Adulthood_Group", "Coached")

levene_test_results <- perform_levenes_test(df, response_vars, factors)

## TWO-WAY ANOVA

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
two_way_anova_results <- perform_two_way_anova(df, response_vars, factors)


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


posthoc_results <- perform_posthoc_tests(df, response_vars, factors, equal_variance = TRUE)

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
      theme(legend.title = element_blank())
    
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


factor1 = "Adulthood_Group"
factor2 = "Coached"

interaction_plots <- create_mean_interaction_plots(df, response_vars, factor1, factor2)
print_and_save_plots(interaction_plots)

## Export Results

correlation_matrix = as.data.frame(correlation_matrix)

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
  "Descriptive Stats" = descriptive_stats, 
  "Descriptive Stats by Group" = descriptive_stats_bygroup, 
  "Reliability" = alpha_results, 
  "Correlation" = correlation_matrix, 
  "Two-way ANOVA" = two_way_anova_results, 
  "Post-hoc Analysis" = posthoc_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")
