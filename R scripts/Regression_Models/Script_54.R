# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/tim_ente")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("CleanData.xlsx")

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


#Declare variables

vars_demogr_cat <- c("Gender", "Nationality", "Location", "Education", 
                     "Startup.Vertical", "B2B_B2C.Focus", "Funding", 
                     "Funding.Stage", "Founded.before")

vars_demogr_cont <- c("Age")

vars_importance_netwevent <- c("Acquisition.of.human.resources", 
                               "Acquisition.of.financial.resources", 
                               "Acquisition.of.information.on.competitors", 
                               "Acquisition.of.market.information", 
                               "Establishment.of.links.with.potential.customers", 
                               "Establishment.of.links.with.potential.partners", 
                               "Branding_Marketing.efforts.for.my.venture", 
                               "Develop.of.legitimacy._.social.status.within.peer.group")

vars_nofevents <- c("How.many.networking.events.did.you.attend.in.the.last.12.months")

vars_bigfive <- c("Neuroticism", "Extraversion", "Openness.to.Experience", 
                  "Agreeableness", "Conscientiousness")

vars_netwbehavior <- c("Network.Building..At.networking.events_.I.proactively.approach.people.I.have.not.met.before.to.expand.my.network", 
                       "Network.Maintenance..At.networking.events_.I.reconnect.with.people.in.my.existing.network.and.use.the.time.to.strengthen.these.relationships", 
                       "Network.Using..At.networking.events_.I.use.my.network.for.e.g..the.introduction.to.new.contacts.or.for.feedback.on.my.startup", 
                       "Occupation.of.brokerage.positions..At.networking.events_.I.introduce.people.who.have.no.prior.connection")


# Sample Characteristis

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


df_freq <- create_frequency_tables(df, vars_demogr_cat)

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
df_descriptive_age <- calculate_descriptive_stats(df, vars_demogr_cont)
df_descriptive_events <- calculate_descriptive_stats(df, vars_nofevents)
df_descriptive_importanceevents <- calculate_descriptive_stats(df, vars_importance_netwevent)



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
  "Netwoking_Behavior" = vars_netwbehavior
)


alpha_results <- reliability_analysis(df, scales)
df_recoded <- alpha_results$data_with_scales
df_reliability <- alpha_results$statistics

vars_netwbehavior2 <- c("Network.Building..At.networking.events_.I.proactively.approach.people.I.have.not.met.before.to.expand.my.network", 
                       "Network.Maintenance..At.networking.events_.I.reconnect.with.people.in.my.existing.network.and.use.the.time.to.strengthen.these.relationships", 
                       
                       "Occupation.of.brokerage.positions..At.networking.events_.I.introduce.people.who.have.no.prior.connection")

scales <- list(
  "Netwoking_Behavior" = vars_netwbehavior2
)

alpha_results <- reliability_analysis(df, scales)
df_recoded <- alpha_results$data_with_scales
df_reliability2 <- alpha_results$statistics


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


correlation_matrix <- calculate_correlation_matrix(df, vars_netwbehavior, method = "pearson")

colnames(df)

# RANKING OF EVENTS

# Create a vector for all X1 to X10 columns
ranking_columns <- c("X1", "X2", "X3", "X4", "X5", "X6", "X7", "X8", "X9", "X10_")

# Count how many times each event is mentioned
event_count <- as.data.frame(table(unlist(df[ranking_columns])))

# Sort by most frequent event
most_frequent_events <- event_count[order(-event_count$Freq), ]

# Top 10 most frequently mentioned events
top_10_most_frequent <- head(most_frequent_events, 10)

# Average rank of events

# Melt the data into long format for easier manipulation
library(reshape2)
df_long <- melt(df, id.vars = NULL, measure.vars = ranking_columns, variable.name = "Rank", value.name = "Event")

# Create a new column for rank weights (X1 = 1, ..., X10 = 10)
df_long$Rank_Weight <- as.numeric(sub("X", "", df_long$Rank))

# Calculate the average rank for each event
avg_rank <- aggregate(Rank_Weight ~ Event, data = df_long, FUN = mean)

# Sort by the lowest average rank (most important)
most_important_events <- avg_rank[order(avg_rank$Rank_Weight), ]

# Top 10 most important events based on average rank
top_10_most_important <- head(most_important_events, 10)

# Rename Var1 to Event
colnames(event_count) <- c("Event", "Count")
# Merge the average rank with the count of how many times each event was cited
avg_rank_with_count <- merge(avg_rank, event_count, by = "Event")

# Sort by the lowest average rank (most important)
most_important_events <- avg_rank_with_count[order(avg_rank_with_count$Rank_Weight), ]

# Top 10 most important events based on average rank
top_10_most_important_with_count <- head(most_important_events, 10)


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


# Define your predictors (IVs) and response variables (DVs)
ivs <- vars_bigfive  # Predictor variables (Big Five personality traits)
dvs_list <- vars_netwbehavior  # Dependent variables (network behavior variables)

# Run the OLS regression iteratively and generate plots
ols_results_list_both_groups <- fit_ols_and_format(
  data = df,
  predictors = ivs,         # Independent variables (Big Five traits)
  response_vars = dvs_list, # Dependent variables (network behaviors)
  save_plots = TRUE         # Save plots for each model
)

# Combine the results into a single dataframe for all DVs
df_modelresults_bothgroups <- bind_rows(ols_results_list_both_groups)

# Display the combined results
print(df_modelresults_bothgroups)



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
    labs(title = "Boxplots of Networking Importance Factors", x = "Variable", y = "Value") +
    theme(axis.text.x = element_text(angle = 90, hjust = 0.5, vjust = 1)) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 10))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = "side_by_side_boxplots.png", plot = p, width = 10, height = 6)
}

create_boxplots(df, vars_importance_netwevent)

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
  "Freq Table" = df_freq,
  "Reliability" = df_reliability,
  "Correlation" = correlation_matrix,
  "Descriptives" = df_descriptive_events,
  "Descriptives 2" = df_descriptive_importanceevents,
  "Ranked Events" = most_important_events,
  "Modelling" = df_modelresults_bothgroups
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")





# Controlled Models
# Assuming df is your dataset

# Recode Nationality to Germany or Other
df$Nationality <- ifelse(df$Nationality == "German", "Germany", "Other")

# Recode Education to merge PhD and Masters into one category
df$Education <- ifelse(df$Education %in% c("PhD", "Masters"), "Masters or PhD", df$Education)

# Recode Startup Vertical to merge all other categories into "Other"
df$Startup.Vertical <- ifelse(df$Startup.Vertical %in% c("Artificial intelligence and machine learning", "Fintech"), 
                              df$Startup.Vertical, "Other")

# Recode Funding Stage to merge Series A, Series B, and Series C into "A or more"
df$Funding.Stage <- ifelse(df$Funding.Stage %in% c("Series A", "Series B", "Series C+"), 
                           "A or more", df$Funding.Stage)

library(fastDummies)


df_recoded <- dummy_cols(df, select_columns = vars_demogr_cat)

names(df_recoded) <- gsub(" ", "_", names(df_recoded))
names(df_recoded) <- gsub("\\(", "_", names(df_recoded))
names(df_recoded) <- gsub("\\)", "_", names(df_recoded))
names(df_recoded) <- gsub("\\-", "_", names(df_recoded))
names(df_recoded) <- gsub("/", "_", names(df_recoded))
names(df_recoded) <- gsub("\\\\", "_", names(df_recoded)) 
names(df_recoded) <- gsub("\\?", "", names(df_recoded))
names(df_recoded) <- gsub("\\'", "", names(df_recoded))
names(df_recoded) <- gsub("\\,", "_", names(df_recoded))

colnames(df_recoded)

vars_controlled <- c(vars_bigfive, "Gender_Female", "Nationality_Germany", "Education_Bachelors", 
                     "Startup.Vertical_Artificial_intelligence_and_machine_learning"     ,                                                                          
                     "Startup.Vertical_Fintech",
                     "B2B_B2C.Focus_B2B",
                     "Funding_Yes",
                     "Funding.Stage_A_or_more" ,                                                                                                                    
                     "Funding.Stage_n_a"    ,                                                                                                                       
                                                                                                                                           
                     "Funding.Stage_Seed" ,
                     "Founded.before_Yes" )


# Define your predictors (IVs) and response variables (DVs)
ivs <- vars_controlled  # Predictor variables (Big Five personality traits)
dvs_list <- vars_netwbehavior  # Dependent variables (network behavior variables)

# Run the OLS regression iteratively and generate plots
ols_results_list_both_groups <- fit_ols_and_format(
  data = df_recoded,
  predictors = ivs,         # Independent variables (Big Five traits)
  response_vars = dvs_list, # Dependent variables (network behaviors)
  save_plots = TRUE         # Save plots for each model
)

# Combine the results into a single dataframe for all DVs
df_modelresults_controlled <- bind_rows(ols_results_list_both_groups)


# Example usage
data_list <- list(
  "Models - Controlled" = df_modelresults_controlled
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Controlled_Models.xlsx")