colnames(df_recoded)


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


vars_coping <- c("Have.you.ever.reported.discriminatory.content.on.social.media.platforms.to.platform.administrators." ,                                                                       
                 "Have.you.ever.unfollowed.or.blocked.individuals.or.accounts.on.social.media.due.to.the.presence.of.discriminatory.content.in.their.posts."   ,                               
                 "Do.you.believe.that.social.media.platforms.effectively.address.and.moderate.discriminatory.content."            ,                                                            
                 "Have.you.ever.participated.in.online.discussions.or.activism.related.to.combating.discriminatory.content." )


df_freq_coping <- create_frequency_tables(df_recoded, vars_coping)


vars_comforttherapy <- c("How.comfortable.do.you.feel.discussing.the.impact.of.discriminatory.content.on.social.media.with.your.peers.or.support.networks."      ,                                     
                         "How.comfortable.would.you.be.discussing.the.impact.of.social.media.and.discriminatory.content.with.a.therapist.or.mental.health.professional." ,                             
                         "To.what.extent.are.you.inclined.to.seek.support.or.engage.with.therapists.specializing.in.mental.health.and.self.esteem.for.individuals.impacted.by.discriminatory.content.")

df_freq_therapy <- create_frequency_tables(df_recoded, vars_comforttherapy)


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
      
      # Calculate the overall counts for each level
      overall_totals <- aggregate(Count ~ Level, data = freq_table, sum)
      freq_table <- merge(freq_table, overall_totals, by = "Level", suffixes = c("", "_Total"))
      freq_table$Percentage <- (freq_table$Count / freq_table$Count_Total) * 100
      
      # Drop the total count column
      freq_table$Count_Total <- NULL
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


categories <- c("Gender"   ,                                                                                                                                                                  
"Education.Level"     ,  
"categorized_Age",
"Occupation...Selected.Choice",
"On.a.whole..I.am.satisfied.with.myself.")

df_freq_segmented <- create_segmented_frequency_tables(df_recoded, categories , vars_comforttherapy)

# Recoding the Rosenberg Self-Esteem Scale
df_recoded <- df_recoded %>%
  mutate(across(starts_with("On.a.whole..I.am.satisfied.with.myself.") |
                  starts_with("At.times.I.think.I.am.no.good.at.all.") |
                  starts_with("I.feel.that.I.have.a.number.of.good.qualities.") |
                  starts_with("I.am.able.to.do.things.as.well.as.most.other.people.") |
                  starts_with("I.feel.I.do.not.have.much.to.be.proud.of.") |
                  starts_with("I.certainly.feel.useless.at.times.") |
                  starts_with("I.feel.that.I.am.a.person.of.worth.") |
                  starts_with("I.wish.I.could.have.more.respect.for.myself.") |
                  starts_with("All.in.all..I.am.inclined.to.think.that.I.am.a.failure.") |
                  starts_with("I.take.a.positive.attitude.toward.myself."),
                ~ case_when(
                  . %in% c(1, 2) ~ "Never/Seldom",
                  . == 3 ~ "Sometimes",
                  . %in% c(4, 5) ~ "Often/Almost Always"
                )))


names(df_recoded)[names(df_recoded) == "categorized_Self-Esteem_Sum"] <- "categorized_Self_Esteem_Sum"

library(tidyr)

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

variable_name <- "ExposureToDiscContent_Sum"
factor_columns <- "categorized_Self_Esteem_Sum"  
descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df_recoded, variable_name, factor_columns)

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
        stars <- ifelse(p_value < 0.01, "***", 
                        ifelse(p_value < 0.05, "**", 
                               ifelse(p_value < 0.1, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}

vars <- c("SelfEsteemSum", "ExposureToDiscContent_Sum")
correlation_matrix <- calculate_correlation_matrix(df, vars, method = "pearson")

vars <- c("ExposureToDiscContent_Sum", "Viewing.discriminatory.content.on.social.media....negatively.affects.my.mood" ,                                                                                              
          "Viewing.discriminatory.content.on.social.media....creates.a.lot.of.anxiety.and.distress.in.me"              ,                                                                
          "Viewing.discriminatory.content.on.social.media....can.affect.how.I.feel.for.the.rest.of.the.day"             ,                                                               
          "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.feel.about.myself"                 ,                                                              
          "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.think.others.in.the..real..world.think.about.me"  ,                                               
          "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.think.about.my.self.worth")


correlation_matrix2 <- calculate_correlation_matrix(df_recoded, vars, method = "spearman")



# OLS Regression

library(broom)

fit_ols_and_format <- function(data, predictors, response_vars, save_plots = FALSE) {
  ols_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Fit the OLS multiple regression model
    lm_model <- lm(formula, data = data)
    
    # Get the summary of the model
    model_summary <- summary(lm_model)
    
    # Extract the R-squared, F-statistic, and p-value
    r_squared <- model_summary$r.squared
    f_statistic <- model_summary$fstatistic[1]
    p_value <- pf(f_statistic, model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE)
    
    # Print the summary of the model for fit statistics
    print(summary(lm_model))
    
    # Extract the tidy output and assign it to ols_results
    ols_results <- broom::tidy(lm_model) %>%
      mutate(ResponseVariable = response_var, R_Squared = r_squared, F_Statistic = f_statistic, P_Value = p_value)
    
    # Optionally print the tidy output
    print(ols_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lm_model) ~ fitted(lm_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lm_model), main = "Q-Q Plot")
      qqline(residuals(lm_model))
      dev.off()
    }
    
    # Store the results in a list
    ols_results_list[[response_var]] <- ols_results
  }
  
  return(ols_results_list)
}

ivs <- c("Type.of.Discriminatory.Content...Negative.media.depictions.and.representations.of.African.Americans" ,                                                                       
                "Type.of.Discriminatory.Content...Stereotypical.depictions.and.representations.of.African.Americans"   ,                                                                      
                "Type.of.Discriminatory.Content...Rhetoric.from.public.figures.and.politicians.targeting.African.Americans",                                                                  
                "Type.of.Discriminatory.Content...Presence.of.online.hate.groups.or.figures.targeting.African.Americans"    ,                                                                 
                "Type.of.Discriminatory.Content...Images.and.videos.depicting.brutal.attacks.targeting.African.Americans")

dvs_list <- "SelfEsteemSum"

ols_results_list_both_groups <- fit_ols_and_format(
  data = df_recoded,
  predictors = ivs,  # Replace with actual predictor names
  response_vars = dvs_list,       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults_new <- bind_rows(ols_results_list_both_groups)

vars <- c("Viewing.discriminatory.content.on.social.media....negatively.affects.my.mood"                      ,                                                                         
          "Viewing.discriminatory.content.on.social.media....creates.a.lot.of.anxiety.and.distress.in.me"      ,                                                                        
          "Viewing.discriminatory.content.on.social.media....can.affect.how.I.feel.for.the.rest.of.the.day"     ,                                                                       
          "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.feel.about.myself"         ,                                                                      
          "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.think.others.in.the..real..world.think.about.me"    ,                                             
          "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.think.about.my.self.worth")

df_freq_effects <- create_frequency_tables(df_recoded, vars)


# Combine Tumblr and Redditt

# Define a function to convert usage levels to a numeric scale
convert_usage_to_numeric <- function(usage) {
  if (usage == "Non Response") return(1)
  else if (usage == "Not at all") return(2)
  else if (usage == "Less than weekly") return(3)
  else if (usage == "A few times a week") return(4)
  else if (usage == "At least once a week") return(5)
  else if (usage == "At least once a day") return(6)
  else if (usage == "Multiple times a day") return(7)
  else return(NA)
}

# Apply the function to both Tumblr and Reddit usage columns
df_recoded$Tumblr_Usage_Numeric <- sapply(df_recoded$Tumblr.Usage, convert_usage_to_numeric)
df_recoded$Reddit_Usage_Numeric <- sapply(df_recoded$Reddit.Usage, convert_usage_to_numeric)

# Combine usage by taking the highest value of the two
df_recoded$Combined_Usage_Numeric <- pmax(df_recoded$Tumblr_Usage_Numeric, df_recoded$Reddit_Usage_Numeric, na.rm = TRUE)

# Convert the combined numeric values back to categorical usage levels
convert_numeric_to_usage <- function(num) {
  if (is.na(num)) return(NA)
  else if (num == 1) return("Non Response")
  else if (num == 2) return("Not at all")
  else if (num == 3) return("Less than weekly")
  else if (num == 4) return("A few times a week")
  else if (num == 5) return("At least once a week")
  else if (num == 6) return("At least once a day")
  else if (num == 7) return("Multiple times a day")
  else return(NA)
}

df_recoded$Combined_Usage <- sapply(df_recoded$Combined_Usage_Numeric, convert_numeric_to_usage)

# Remove intermediate numeric columns if not needed
df_recoded <- df_recoded[, !names(df_recoded) %in% c("Tumblr_Usage_Numeric", "Reddit_Usage_Numeric", "Combined_Usage_Numeric")]


frequency_tumblrreditt <- create_frequency_tables(df_recoded, "Combined_Usage")


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
  "Frequency - Coping" = df_freq_coping,
  "Frequency - Therapy" = df_freq_therapy,
  "Frequency - Comfort Therapy" = df_freq_segmented,
  "Frequency - Effects" = df_freq_effects,
  "Exposure/Selfeesteem" = descriptive_stats_bygroup,
  "Correlation 1" = correlation_matrix,
  "Correlation 2" = correlation_matrix2,
  "Model" = df_modelresults_new
  
  
)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/tiffanysb91")

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables3.xlsx")