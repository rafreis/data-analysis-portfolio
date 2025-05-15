# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/mcankaya")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Dissertation_August_20__2024_13.09_Cleaned.xlsx")

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

colnames(df)


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


df_freq_sociodem <- create_frequency_tables(df, c("GENDER", "EDUCATION"))

# Recoding Recommendation Agent
df$Recommendation_Agent <- ifelse(!is.na(df$FL_16_DO_SCENARIO1) | !is.na(df$FL_16_DO_SCENARIO2), "Human",
                                  ifelse(!is.na(df$FL_16_DO_SCENARIO3) | !is.na(df$FL_16_DO_SCENARIO4), "Brand",
                                         ifelse(!is.na(df$FL_16_DO_SCENARIO5) | !is.na(df$FL_16_DO_SCENARIO6), "AI", NA)))

# Recoding Level of Personalization
df$Level_of_Personalization <- ifelse(!is.na(df$FL_16_DO_SCENARIO1) | !is.na(df$FL_16_DO_SCENARIO3) | !is.na(df$FL_16_DO_SCENARIO5), "Low",
                                      ifelse(!is.na(df$FL_16_DO_SCENARIO2) | !is.na(df$FL_16_DO_SCENARIO4) | !is.na(df$FL_16_DO_SCENARIO6), "High", NA))

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
      
      # Add the variable name and calculate percentages
      freq_table$Variable <- var
      freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


df_freq_scenarios <- create_segmented_frequency_tables(df, "Recommendation_Agent", "Level_of_Personalization")



## RELIABILITY ANALYSIS

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
  "Sense of Uniqueness" = c("Q10_1", "Q10_2", "Q10_3", "Q10_4", "Q10_5"),
  "Perceived Vulnerability" = c("A_PV_1_5_1", "A_PV_1_5_2", "A_PV_1_5_3", "A_PV_1_5_4", "A_PV_1_5_5"),
  "Competence" = c("CMPT_1_5_1", "CMPT_1_5_2", "CMPT_1_5_3", "CMPT_1_5_4", "CMPT_1_5_5"),
  "Benevolence" = c("BNVL_1_3_1", "BNVL_1_3_2", "BNVL_1_3_3"),
  "Integrity" = c("INTG_1_3_1", "INTG_1_3_2", "INTG_1_3_3"),
  "Emotional Trust" = c("ET_1_3_1", "ET_1_3_2", "ET_1_3_3"),
  "Purchase Intention" = c("PI_1_5_1", "PI_1_5_2", "PI_1_5_3", "PI_1_5_4", "PI_1_5_5"),
  "Cognitive Trust" = c("CMPT_1_5_1", "CMPT_1_5_2", "CMPT_1_5_3", "CMPT_1_5_4", "CMPT_1_5_5", 
                        
                        "INTG_1_3_1", "INTG_1_3_2", "INTG_1_3_3")
)

# Reverse coding Q10_3
df$Q10_3 <- 8 - df$Q10_3

alpha_results <- reliability_analysis(df, scales)

df_recoded <- alpha_results$data_with_scales
df_descriptives <- alpha_results$statistics


# Create a list of specific variables
scales <- c(
  "Sense of Uniqueness",
  "Perceived Vulnerability",
  "Competence",
  "Benevolence",
  "Integrity",
  "Emotional Trust",
  "Purchase Intention",
  "Cognitive Trust"
)

# Outlier Evaluation

library(dplyr)

calculate_z_scores <- function(data, vars, id_var, z_threshold = 3) {
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~ (.-mean(.))/sd(.), .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  return(z_score_data)
}


df_zscores <- calculate_z_scores(df_recoded, scales, "ResponseId")

library(dplyr)
library(tidyr)
library(stringr)

# Function to remove values in df_recoded based on z-score table
remove_outliers_based_on_z_scores <- function(data, z_score_data, id_var) {
  # Reshape the z-score data to long format
  z_score_data_long <- z_score_data %>%
    pivot_longer(cols = starts_with("z_"), names_to = "Variable", values_to = "z_value") %>%
    pivot_longer(cols = starts_with("flag_z_"), names_to = "Flag_Variable", values_to = "Flag_Value") %>%
    filter(str_replace(Flag_Variable, "flag_", "") == Variable) %>%
    select(!!sym(id_var), Variable, Flag_Value)
  
  # Remove rows where z-scores are flagged as outliers
  for (i in 1:nrow(z_score_data_long)) {
    if (z_score_data_long$Flag_Value[i] == "Outlier") {
      col_name <- gsub("z_", "", z_score_data_long$Variable[i])
      data <- data %>%
        mutate(!!sym(col_name) := ifelse(!!sym(id_var) == z_score_data_long[[id_var]][i], NA, !!sym(col_name)))
    }
  }
  return(data)
}

df_recoded_nooutliers <- remove_outliers_based_on_z_scores(df_recoded, df_zscores, "ResponseId")



## BOXPLOTS DIVIDED BY FACTOR

library(ggplot2)

colnames(df_recoded_nooutliers)

library(ggplot2)

# Function to create and save boxplots for given variables
create_and_save_boxplots <- function(df, vars, factor_var) {
  for (var in vars) {
    # Ensure the variable names are correctly formatted
    var_sym <- sym(var)
    factor_var_sym <- sym(factor_var)
    
    # Create boxplot for each variable using ggplot2
    p <- ggplot(df, aes(x = !!factor_var_sym, y = !!var_sym)) +
      geom_boxplot() +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
      labs(title = paste("Boxplot of", var, "by", factor_var),
           x = factor_var, y = var)
    
    # Print the plot
    print(p)
    
    # Save the plot to file
    ggsave(filename = paste0("boxplot_", var, ".png"), plot = p, width = 10, height = 6)
  }
}

create_and_save_boxplots(df_recoded_nooutliers, scales, "Recommendation_Agent")
create_and_save_boxplots(df_recoded_nooutliers, scales, "Level_of_Personalization")

library(dplyr)

# Update dataset with a new combined factor column
df_recoded_nooutliers <- df_recoded_nooutliers %>%
  mutate(Combined_Factor = interaction(Recommendation_Agent, Level_of_Personalization, sep = " - "))

library(ggplot2)
library(rlang)

create_and_save_boxplots <- function(df, vars, factor_var) {
  for (var in vars) {
    # Ensure the variable names are correctly formatted as symbols
    var_sym <- sym(var)
    factor_var_sym <- sym(factor_var)
    
    # Create boxplot for each variable using ggplot2
    p <- ggplot(df, aes(x = !!factor_var_sym, y = !!var_sym)) +
      geom_boxplot() +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
      labs(title = paste("Boxplot of", var, "by", factor_var),
           x = factor_var, y = var) +
      theme(axis.text.x = element_text(hjust = 0.5)) # Improve x-axis labels readability
    
    # Print the plot
    print(p)
    
    # Save the plot to file
    ggsave(filename = paste0("boxplot_", var, "_", gsub(" ", "_", factor_var), ".png"), plot = p, width = 10, height = 6)
  }
}

# Example usage
create_and_save_boxplots(df_recoded_nooutliers, scales, "Combined_Factor")

## CORRELATION ANALYSIS

colnames(df_recoded_nooutliers)

scales <- c(
  "Sense of Uniqueness",
  "Perceived Vulnerability",
  "Competence",
  "Benevolence",
  "Integrity",
  "Emotional Trust",
  "Purchase Intention",
  "Cognitive Trust"
)

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


correlation_matrix <- calculate_correlation_matrix(df_recoded_nooutliers, scales, method = "pearson")

scales <- c(
  "Sense_of_Uniqueness",
  "Perceived_Vulnerability",
  "Competence",
  "Benevolence",
  "Integrity",
  "Emotional_Trust",
  "Purchase_Intention",
  "Cognitive_Trust"
)

library(stats)
library(broom)

perform_3x2_anova <- function(df, response_var, factor1, factor2) {
  # Ensure the variable names are correctly formatted as formulas
  formula_string <- as.formula(paste(response_var, "~", factor1, "*", factor2))
  
  # Fit the ANOVA model
  anova_model <- aov(formula_string, data = df)
  
  # Return tidy results
  tidy_results <- tidy(anova_model)
  return(tidy_results)
}

names(df_recoded_nooutliers) <- gsub(" ", "_", trimws(names(df_recoded_nooutliers)))
colnames(df_recoded_nooutliers)

df_results_mainrq1 <- perform_3x2_anova(df_recoded_nooutliers, "Purchase_Intention", "Recommendation_Agent", "Level_of_Personalization")
df_results_mainrq2 <- perform_3x2_anova(df_recoded_nooutliers, "Emotional_Trust", "Recommendation_Agent", "Level_of_Personalization")
df_results_mainrq3 <- perform_3x2_anova(df_recoded_nooutliers, "Cognitive_Trust", "Recommendation_Agent", "Level_of_Personalization")


library(ggplot2)
library(dplyr)

create_interaction_plot <- function(df, response_var, factor1, factor2, filename = "interaction_plot.png") {
  # Ensure factors are appropriately set
  df[[factor1]] <- as.factor(df[[factor1]])
  df[[factor2]] <- as.factor(df[[factor2]])
  
  # Create the interaction plot
  p <- ggplot(df, aes_string(x = factor1, y = response_var, color = factor2, group = factor2)) +
    geom_line(aes(linetype = factor2), stat = "summary", fun.y = "mean") +
    geom_point(aes(shape = factor2), stat = "summary", fun.y = "mean", size = 3) +
    labs(title = paste("Interaction of", factor1, "and", factor2, "on", response_var),
         x = factor1, y = response_var) +
    theme_minimal() +
    scale_color_brewer(palette = "Set1") +
    theme(legend.title = element_blank())
  
  # Save the plot to the working directory
  ggsave(filename = filename, plot = p, width = 8, height = 6)
  
  # Return the plot
  return(p)
}

# Plotting Purchase Intention
plot_purchase_intention <- create_interaction_plot(df_recoded_nooutliers, "Purchase_Intention", "Recommendation_Agent", "Level_of_Personalization", filename = "Interaction_PI.png")
print(plot_purchase_intention)

# Plotting Emotional Trust
plot_emotional_trust <- create_interaction_plot(df_recoded_nooutliers, "Emotional_Trust", "Recommendation_Agent", "Level_of_Personalization", filename = "Interaction_ET.png")
print(plot_emotional_trust)

# Plotting Cognitive Trust
plot_cognitive_trust <- create_interaction_plot(df_recoded_nooutliers, "Cognitive_Trust", "Recommendation_Agent", "Level_of_Personalization", filename = "Interaction_CT.png")
print(plot_cognitive_trust)



library(mediation)

# Mediation model function updated
perform_mediation_analysis <- function(df, mediator_var, outcome_var, factor1, factor2) {
  # Create interaction term
  df$interaction_term <- interaction(df[[factor1]], df[[factor2]])
  
  # Filter the dataframe to include only the specific interaction terms
  specific_interactions <- c("AI.Low", "AI.High")
  df_filtered <- df[df$interaction_term %in% specific_interactions,]
  
  # Check if the filtered dataframe has enough data
  if (nrow(df_filtered) < 20) {
    stop("Not enough data points for the specified interaction terms.")
  }
  
  # Model for the mediator
  mediator_model <- lm(as.formula(paste(mediator_var, "~ interaction_term")), data = df_filtered)
  
  # Model for the outcome including the mediator
  outcome_model <- lm(as.formula(paste(outcome_var, "~", mediator_var, "+ interaction_term")), data = df_filtered)
  
  # Perform the mediation analysis
  med_analysis <- mediate(
    mediator_model, outcome_model,
    treat = "interaction_term", 
    mediator = mediator_var,
    control.value = "AI.Low", 
    treat.value = "AI.High",
    sims = 500
  )
  
  # Return summary of the mediation analysis
  summary_med <- summary(med_analysis)
  return(summary_med)
}

df_recoded_nooutliers$Recommendation_Agent <- as.factor(df_recoded_nooutliers$Recommendation_Agent)
df_recoded_nooutliers$Level_of_Personalization <- as.factor(df_recoded_nooutliers$Level_of_Personalization)

df_results_mediation_pi_rq1 <- perform_mediation_analysis(df_recoded_nooutliers, "Perceived_Vulnerability", "Purchase_Intention", "Recommendation_Agent", "Level_of_Personalization")
df_results_mediation_et_rq1 <- perform_mediation_analysis(df_recoded_nooutliers, "Perceived_Vulnerability", "Emotional_Trust", "Recommendation_Agent", "Level_of_Personalization")
df_results_mediation_ct_rq1 <- perform_mediation_analysis(df_recoded_nooutliers, "Perceived_Vulnerability", "Cognitive_Trust", "Recommendation_Agent", "Level_of_Personalization")

# Function to create a DataFrame from mediation results
extract_mediation_results_to_df <- function(results) {
  # Create a dataframe from extracted results
  mediation_df <- data.frame(
    Metric = c("ACME (Average Causal Mediation Effect)", "ADE (Average Direct Effect)",
               "Total Effect", "Proportion Mediated"),
    Estimate = c(results$d0, results$z0, results$tau.coef, results$n0),
    `95% CI Lower` = c(results$d0.ci["2.5%"], results$z0.ci["2.5%"], 
                       results$tau.ci["2.5%"], results$n0.ci["2.5%"]),
    `95% CI Upper` = c(results$d0.ci["97.5%"], results$z0.ci["97.5%"],
                       results$tau.ci["97.5%"], results$n0.ci["97.5%"]),
    `p-value` = c(results$d0.p, results$z0.p, results$tau.p, results$n0.p)
  )
  
  # Return the dataframe
  return(mediation_df)
}

# Apply the function to your mediation result
df_results_mediation_pi_rq1 <- extract_mediation_results_to_df(df_results_mediation_pi_rq1)
df_results_mediation_et_rq1 <- extract_mediation_results_to_df(df_results_mediation_et_rq1)
df_results_mediation_ct_rq1 <- extract_mediation_results_to_df(df_results_mediation_ct_rq1)

library(stats)
library(broom)

# Two-way ANOVA with a moderator function
perform_3x2_anova_with_moderator <- function(df, outcome_var, factor1, factor2, moderator) {
  # Ensure factors and moderator are treated as factors
  df[[factor1]] <- as.factor(df[[factor1]])
  df[[factor2]] <- as.factor(df[[factor2]])
  df[[moderator]] <- as.factor(df[[moderator]])
  
  # Create interaction terms between the factors and the moderator
  df$interaction_f1_m <- interaction(df[[factor1]], df[[moderator]])
  df$interaction_f2_m <- interaction(df[[factor2]], df[[moderator]])
  df$three_way_interaction <- interaction(df[[factor1]], df[[factor2]], df[[moderator]])
  
  # Fit the ANOVA model including the three-way interaction
  formula_string <- paste(outcome_var, "~", factor1, "*", factor2, "+", moderator,
                          "+ interaction_f1_m + interaction_f2_m + three_way_interaction")
  anova_model <- aov(as.formula(formula_string), data = df)
  
  # Use broom to tidy the ANOVA model into a dataframe
  tidy_results <- tidy(anova_model)
  
  return(tidy_results)
}

df_results_moderation_pi_rq2 <- perform_3x2_anova_with_moderator(df_recoded_nooutliers, "Purchase_Intention", "Recommendation_Agent", "Level_of_Personalization", "Sense_of_Uniqueness")
df_results_moderation_et_rq2 <- perform_3x2_anova_with_moderator(df_recoded_nooutliers, "Emotional_Trust", "Recommendation_Agent", "Level_of_Personalization", "Sense_of_Uniqueness")
df_results_moderation_ct_rq2 <- perform_3x2_anova_with_moderator(df_recoded_nooutliers, "Cognitive_Trust", "Recommendation_Agent", "Level_of_Personalization", "Sense_of_Uniqueness")




library(stats)
library(broom)

# Function to perform 3x2 ANOVA with optional covariates
perform_3x2_anova_cov <- function(df, response_var, factor1, factor2, covariates=NULL) {
  # Base formula
  formula_string <- paste(response_var, "~", factor1, "*", factor2)
  
  # Add covariates to the formula if provided
  if (!is.null(covariates)) {
    # Ensure covariates are provided as a character vector
    if (is.character(covariates)) {
      covariate_string <- paste(covariates, collapse = " + ")
      formula_string <- paste(formula_string, "+", covariate_string)
    } else {
      stop("Covariates should be provided as a character vector.")
    }
  }
  
  # Convert string to formula
  formula_obj <- as.formula(formula_string)
  
  # Fit the ANOVA model using the specified formula
  anova_model <- aov(formula_obj, data = df)
  
  # Return tidy results
  tidy_results <- tidy(anova_model)
  return(tidy_results)
}

# Example usage with covariates
df_results_rq3 <- perform_3x2_anova_cov(df_recoded_nooutliers, "Purchase_Intention", "Recommendation_Agent", "Level_of_Personalization", c("Emotional_Trust", "Cognitive_Trust"))


# Additional Descriptives

library(dplyr)

calculate_descriptive_stats_bygroups <- function(data, desc_vars, cat_column) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each category
  for (var in desc_vars) {
    # Group data by the categorical column
    grouped_data <- data %>%
      group_by(!!sym(cat_column)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE)
        
      ) %>%
      mutate(Variable = var)
    
    # Append the results for each variable and category to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}


df_descriptives_by_condition <- calculate_descriptive_stats_bygroups(df_recoded_nooutliers, scales, "Combined_Factor")

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
 ANOVA_Purchase_Intention = df_results_mainrq1,
  ANOVA_Emotional_Trust = df_results_mainrq2,
  ANOVA_Cognitive_Trust = df_results_mainrq3,
  Mediation_PI_RQ1 = df_results_mediation_pi_rq1,
  Mediation_ET_RQ1 = df_results_mediation_et_rq1,
  Mediation_CT_RQ1 = df_results_mediation_ct_rq1,
  
  Frequency_Sociodemographics = df_freq_sociodem,
  Frequency_Scenarios = df_freq_scenarios,
  Descriptive_Statistics = df_descriptives,
  Descriptive_Stats_Condition = df_descriptives_by_condition
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_TablesNoBenevolence.xlsx")

# Example usage
data_list <- list(
 Correlation = correlation_matrix
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Correlations.xlsx")










# Manipulation Check

## DESCRIPTIVE STATS BY FACTORS - FACTORS ARE LEVELS OF A SINGLE COLUMN

calculate_means_and_sds_by_factors <- function(data, variable_names, factor_columns) {
  # Create an empty dataframe to store results
  results <- data.frame()
  
  # Create a new interaction factor that combines all factor columns
  data$interaction_factor <- interaction(data[factor_columns], drop = TRUE)
  
  # Iterate over each variable name
  for (variable_name in variable_names) {
    # Calculate means and SDs for each combination of factors for the variable
    aggregated_data <- aggregate(data[[variable_name]], 
                                 by = list(data$interaction_factor), 
                                 FUN = function(x) c(Mean = mean(x, na.rm = TRUE), SD = sd(x, na.rm = TRUE)))
    
    # Rename the columns and reshape the aggregated data
    names(aggregated_data)[1] <- "Factors"
    variable_results <- do.call(rbind, lapply(1:nrow(aggregated_data), function(i) {
      tibble(
        Factors = as.character(aggregated_data$Factors[i]),
        Variable = variable_name,
        Mean = aggregated_data$x[i, "Mean"],
        SD = aggregated_data$x[i, "SD"]
      )
    }))
    
    # Append results
    results <- bind_rows(results, variable_results)
  }
  
  # Split the interaction factor back into the original factors
  results <- results %>% 
    separate(Factors, into = factor_columns, sep = "\\.", remove = FALSE)
  
  return(results)
}

variable_name <- c("MC_1_4_1"       ,          "MC_1_4_2"   ,              "MC_1_4_3" ,               
                   "MC_1_4_4")
factor_columns <- c("Recommendation_Agent", "Level_of_Personalization")  
df_manipulationcheck_combined <- calculate_means_and_sds_by_factors(df_recoded_nooutliers, variable_name, factor_columns)


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

df_manipulationcheck <- calculate_means_and_sds_by_factors(df_recoded_nooutliers, variable_name, factor_columns)


# Validity Check

factor_analysis_loadings <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Scale = character(),
    Variable = character(),
    FactorLoading = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[,scales[[scale]]]
    
    # Remove rows with any missing values in the subset data
    subset_data <- subset_data[complete.cases(subset_data), ]
    
    # Perform factor analysis; using 1 factor as it's the most common scenario
    fa_results <- factanal(subset_data, factors = 1, rotation = "varimax")
    
    # Extract factor loadings
    loadings <- fa_results$loadings[,1]
    
    # Append results for each item in the scale
    for (item in scales[[scale]]) {
      results <- rbind(results, data.frame(
        Scale = scale,
        Variable = item,
        FactorLoading = loadings[item]
      ))
    }
    # Calculate total variance explained
    variance_explained <- sum(loadings^2) / length(loadings)
    
    # Print summary of the results - Total Variance Explained
    cat("Scale:", scale, "\n")
    cat("Total Variance Explained by Factor:", variance_explained, "\n")
  }
  
  return(results)
}

scales <- list(
  "Sense of Uniqueness" = c("Q10_1", "Q10_2", "Q10_3", "Q10_4", "Q10_5"),
  "Perceived Vulnerability" = c("A_PV_1_5_1", "A_PV_1_5_2", "A_PV_1_5_3", "A_PV_1_5_4", "A_PV_1_5_5"),
  "Competence" = c("CMPT_1_5_1", "CMPT_1_5_2", "CMPT_1_5_3", "CMPT_1_5_4", "CMPT_1_5_5"),
  "Benevolence" = c("BNVL_1_3_1", "BNVL_1_3_2", "BNVL_1_3_3"),
  "Integrity" = c("INTG_1_3_1", "INTG_1_3_2", "INTG_1_3_3"),
  "Emotional Trust" = c("ET_1_3_1", "ET_1_3_2", "ET_1_3_3"),
  "Purchase Intention" = c("PI_1_5_1", "PI_1_5_2", "PI_1_5_3", "PI_1_5_4", "PI_1_5_5"),
  "Cognitive Trust" = c("CMPT_1_5_1", "CMPT_1_5_2", "CMPT_1_5_3", "CMPT_1_5_4", "CMPT_1_5_5", 
                        
                        "INTG_1_3_1", "INTG_1_3_2", "INTG_1_3_3")
)

df_validity <- factor_analysis_loadings(df_recoded_nooutliers, scales)



# Example usage
data_list <- list(
  ManipulationCheck = df_manipulationcheck,
  ModerationAnalysisCT = df_results_moderation_ct_rq2,
  ModerationAnalysisET = df_results_moderation_et_rq2,
  ModerationAnalysisPI = df_results_moderation_pi_rq2,
  FactorAnalysisValidity = df_validity
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_New.xlsx")
