setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/richsylvester")

library(readxl)
df <- read_xlsx('Plyo _ sprint _ CON main copy-current-Dec 18 2023.xlsx')

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

# Convert to factor
df$Group <- factor(df$Group)
                     
# Reshape Dataframe

library(tidyverse)
library(reshape2)

# Rename the column
df <- df %>%
  rename(MatStatus = MatStatus_Pre)

# Drop the "MatStatus_Post" column
df <- df %>%
  select(-MatStatus_Post)

# Drop the "MatStatus_Post" column
df <- df %>%
  select(-Sport)

# Melt the data to long format while keeping the 'Group' variable
df_long <- df %>%
  pivot_longer(
    cols = -c(Code, Group, MatStatus), # Columns to exclude from reshaping
    names_to = c("Variable", "Time"), # Names for the new variable and time columns
    names_pattern = "(.+)_(?i)(pre|post)$" # Regex pattern to separate the variable names and time indicators (case-insensitive)
  )
# View the reshaped dataframe
head(df_long)
str(df)

#Convert variable
df_long$value <- as.numeric(df_long$value)

## DESCRIPTIVE STATS BY FACTORS

calculate_means_and_sds_by_factors <- function(data, variable_name, factor_columns) {
  # Create an empty dataframe to store results
  results <- data.frame()
  
  # Create a new interaction factor that combines all factor columns
  data$interaction_factor <- interaction(data[factor_columns], drop = TRUE)
  
  # Calculate means and SDs for each combination of factors for the variable
  aggregated_data <- aggregate(data[[variable_name]], 
                               by = list(data$interaction_factor), 
                               FUN = function(x) c(Mean = mean(x, na.rm = TRUE), SD = sd(x, na.rm = TRUE)))
  
  # Rename the columns and reshape the aggregated data
  names(aggregated_data)[1] <- "Factors"
  results <- do.call(rbind, lapply(1:nrow(aggregated_data), function(i) {
    tibble(
      Factors = as.character(aggregated_data$Factors[i]),
      Mean = aggregated_data$x[i, "Mean"],
      SD = aggregated_data$x[i, "SD"]
    )
  }))
  
  # Split the interaction factor back into the original factors
  results <- results %>% 
    separate(Factors, into = factor_columns, sep = "\\.", remove = FALSE)
  
  return(results)
}

variable_name <- "value"
factor_columns <- c("Variable", "Group", "Time")  
descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df_long, variable_name, factor_columns)

# Generate Pre/Post Boxplots

library(ggplot2)
library(tidyverse)

# Function to plot pre/post boxplots for each variable
plot_pre_post_boxplots <- function(data, group_col, time_col, value_col) {
  # Get the list of variables to plot
  variables <- unique(data$Variable)
  
  # Loop through each variable and create a plot
  plots <- list()
  for (variable in variables) {
    p <- data %>%
      filter(Variable == variable) %>%
      ggplot(aes(x = .data[[time_col]], y = .data[[value_col]], fill = .data[[group_col]])) +
      geom_boxplot(position = position_dodge(0.8)) +
      stat_summary(
        fun = mean, 
        geom = "point", 
        shape = 18, 
        size = 3, 
        color = "red", 
        position = position_dodge(width = 0.8), # Match the dodge position with the boxplot
        aes(group = interaction(.data[[group_col]], .data[[time_col]]))
      ) +
      labs(title = paste("Boxplot of", variable, "Pre vs Post"),
           x = "Time",
           y = "Value") +
      theme_minimal() +
      scale_x_discrete(limits = c("Pre", "Post")) # Ensure correct order of x-axis
    plots[[variable]] <- p
    
    # Save the plot
    ggsave(filename = paste("boxplot_", variable, ".png", sep = ""), plot = p, width = 10, height = 6)
  }
  
  return(plots)
}

plots <- plot_pre_post_boxplots(df_long, "Group", "Time", "value")

print(plots$SH)


## ASsess normality

df_long <- df_long[complete.cases(df_long), ]

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
        # Calculate skewness and kurtosis
        skew <- skewness(data_subset[[var]], na.rm = TRUE)
        kurt <- kurtosis(data_subset[[var]], na.rm = TRUE)
        
        # Perform Shapiro-Wilk test
        shapiro_test <- shapiro.test(data_subset[[var]])
        
        # Add results to the dataframe
        results <- rbind(results, c(level, var, skew, kurt, shapiro_test$statistic, shapiro_test$p.value))
      } else {
        warning(paste("Variable", var, "not found in the data. Skipping."))
      }
    }
  }
  
  colnames(results) <- c("Group", "Variable", "Skewness", "Kurtosis", "Shapiro_Wilk_F", "Shapiro_Wilk_p_value")
  return(results)
}

normality_results <- calculate_stats_by_group(df_long, "value", "Variable")


## Linear Mixed Effects Model

library(dplyr)

df_long <- df_long %>%
  mutate(Time = ifelse(Time == "pre", "Pre", ifelse(Time == "post", "Post", Time)))

library(lme4)
library(broom.mixed)
library(afex)

# Function to fit LMM for each unique variable in the 'Variable' column and return a single dataframe with results
fit_lmm_and_format <- function(data, variable_col, value_col, within_subject_var, between_subject_var, categorical_var, random_effects, save_plots = FALSE) {
  all_lmm_results <- list()
  
  # Ensure the categorical variable is treated as a factor
  data[[categorical_var]] <- as.factor(data[[categorical_var]])
  
  # Ensure the categorical variable is treated as a factor and set "Pre" as the reference level
  data[[within_subject_var]] <- factor(data[[within_subject_var]], levels = c("Pre", "Post"))  # Update this line if more levels exist
  
  # Get the list of unique variables to serve as dependent variables
  response_vars <- unique(data[[variable_col]])
  
  for (response_var in response_vars) {
    # Use tryCatch to handle errors and continue the loop
    tryCatch({
      # Subset the data for the current response variable
      data_subset <- data[data[[variable_col]] == response_var, ]
      
      # Construct the formula dynamically
      formula <- as.formula(paste(value_col, "~", within_subject_var, "*", between_subject_var, "+", within_subject_var, ":", categorical_var, "+", random_effects))
      
      # Fit the linear mixed-effects model
      lmer_model <- lmer(formula, data = data_subset)
      
      # Extract the tidy output and assign it to lmm_results
      lmm_results <- tidy(lmer_model, effects = "fixed")
      lmm_results$Variable <- response_var  # Add the response variable name to the results dataframe
      
      # Optionally print the tidy output
      print(lmm_results)
      
      # Generate and save residual plots
      if (save_plots) {
        plot_name_prefix <- gsub("[[:space:]]+", "_", response_var)
        jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
        plot(residuals(lmer_model) ~ fitted(lmer_model), main = paste("Residuals vs Fitted:", response_var), xlab = "Fitted values", ylab = "Residuals")
        dev.off()
        
        jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
        qqnorm(residuals(lmer_model), main = paste("Q-Q Plot:", response_var))
        qqline(residuals(lmer_model))
        dev.off()
      }
      
      # Store the results in a list
      all_lmm_results[[response_var]] <- lmm_results
    }, error = function(e) {
      message(paste("Error in variable", response_var, ": ", e$message))
    })
  }
  
  # Combine all results into a single dataframe
  combined_results <- bind_rows(all_lmm_results, .id = "Variable")
  
  return(combined_results)
}

# Example usage:
combined_lmm_results_2 <- fit_lmm_and_format(
  data = df_long,
  variable_col = "Variable",
  value_col = "value",
  within_subject_var = "Time",
  between_subject_var = "Group",
  categorical_var = "MatStatus",
  random_effects = "(1|Code)",
  save_plots = TRUE
)



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
  "LMM Results" = combined_lmm_results_2
)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/richsylvester/Order 3")

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_firstorder.xlsx")