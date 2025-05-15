setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/onkelj99")

library(openxlsx)
data <- read.xlsx("Variable Compile with new instructions (1).xlsx")

# Get rid of special characters

names(data) <- gsub(" ", "_", trimws(names(data)))
names(data) <- gsub("\\s+", "_", trimws(names(data), whitespace = "[\\h\\v\\s]+"))
names(data) <- gsub("\\(", "_", names(data))
names(data) <- gsub("\\)", "_", names(data))
names(data) <- gsub("\\-", "_", names(data))
names(data) <- gsub("/", "_", names(data))
names(data) <- gsub("\\\\", "_", names(data)) 
names(data) <- gsub("\\?", "", names(data))
names(data) <- gsub("\\'", "", names(data))
names(data) <- gsub("\\,", "_", names(data))
names(data) <- gsub("\\$", "", names(data))
names(data) <- gsub("\\+", "_", names(data))

# Trim all values

# Loop over each column in the dataframe
data <- data.frame(lapply(data, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))

library(dplyr)


# Declare Variables

data <- data %>%
  group_by(Ticker) %>%  # Group data by 'Ticker'
  mutate(has_foreign_sales = ifelse(any(Foreign.sales.to.total.sales > 0), 1, 0)) %>%
  ungroup()  # Remove grouping

# Create a named vector with the appreciation status for each year
appreciation_status <- c("2006" = "Depreciated", "2007" = "Depreciated", "2008" = "Appreciated",
                         "2009" = "Appreciation", "2010" = "Depreciation", "2011" = "Appreciation",
                         "2012" = "Depreciation", "2013" = "Depreciation", "2014" = "Appreciated",
                         "2015" = "Appreciated", "2016" = "Appreciated", "2017" = "Depreciated",
                         "2018" = "Appreciated", "2019" = "Appreciated", "2020" = "Appreciated",
                         "2021" = "Depreciated", "2022" = "Appreciated", "2023" = "Appreciation")

# Standardize the values to 'Appreciated' or 'Depreciated' (handle different phrasings)
appreciation_status <- gsub("Appreciation", "Appreciated", appreciation_status)
appreciation_status <- gsub("Depreciation", "Depreciated", appreciation_status)

# Convert the status to binary: 1 for 'Appreciated', 0 for 'Depreciated'
binary_status <- ifelse(appreciation_status == "Appreciated", 1, 0)

# Add a new column 'dollar_appreciation' to the dataframe based on the 'Year' column
data <- data %>%
  mutate(dollar_appreciation = binary_status[as.character(Year)])

var_use_derivatives <- "Use.of.foreign.currency.derivatives"
var_foreign_operations <- "has_foreign_sales"
var_dollar_appreciation <- "dollar_appreciation"
var_firmvalue_measures <- c("Market.to.book.ratio", "Market.to.Book.value.per.share", "Tobins.Q")
var_control_numeric <- c("Ln_Total_Asset","Ln_Revenue", "Long.Term.Debt_Total.Equity._._", "ROA._._" , 
                 "R.D.expense" , "Foreign.sales.to.total.sales" )

var_control_categorical <- c("Dividend" , "Business.Segment")
                             
var_industry_sector <- "Industry.Code"
var_time <- "Year"

# Convert all these variables to numeric
data <- data %>%
  mutate(across(c(var_firmvalue_measures, var_control_numeric), ~ as.numeric(as.character(.))))


# OUTLIER INSPECTION

# IQR Deviation

calculate_extreme_values_count <- function(data, vars, threshold = 1.5) {
  results <- data.frame(variable = character(), lower_extreme = numeric(), upper_extreme = numeric(), count_outliers = integer())
  data$Outlier_Flag <- integer(nrow(data))  # Initialize a column to flag rows with outliers
  
  for (var in vars) {
    var_data <- data[[var]]
    Q1 <- quantile(var_data, 0.25, na.rm = TRUE)
    Q3 <- quantile(var_data, 0.75, na.rm = TRUE)
    IQR <- Q3 - Q1
    
    # Calculate lower and upper extreme values
    lower_extreme <- Q1 - threshold * IQR
    upper_extreme <- Q3 + threshold * IQR
    
    # Identify outliers
    outlier_indices <- which(var_data < lower_extreme | var_data > upper_extreme)
    data$Outlier_Flag[outlier_indices] <- 1  # Mark rows containing outliers
    
    # Count outliers
    outliers <- length(outlier_indices)
    
    # Append to results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Lower_Extreme = lower_extreme,
      Upper_Extreme = upper_extreme,
      Count_Outliers = outliers
    ))
  }
  
  # Calculate listwise outlier count
  listwise_outliers <- sum(data$Outlier_Flag == 1, na.rm = TRUE)
  
  # Clean up by removing the outlier flag column
  data$Outlier_Flag <- NULL
  
  # Return results and the listwise outlier count
  list(results = results, listwise_count = listwise_outliers)
}

extreme_values_info <- calculate_extreme_values_count(data, c(var_firmvalue_measures, var_control_numeric), threshold = 3)
df_extreme_values <- extreme_values_info$results
cat("Total number of cases with any outliers across specified variables:", extreme_values_info$listwise_count, "\n")


delete_extreme_values <- function(data, vars, threshold = 1.5) {
  for (var in vars) {
    var_data <- data[[var]]
    Q1 <- quantile(var_data, 0.25, na.rm = TRUE)
    Q3 <- quantile(var_data, 0.75, na.rm = TRUE)
    IQR <- Q3 - Q1
    
    lower_extreme <- Q1 - threshold * IQR
    upper_extreme <- Q3 + threshold * IQR
    
    # Set extreme values to NA
    data[[var]][var_data < lower_extreme | var_data > upper_extreme] <- NA
  }
  return(data)
}

# Example usage to remove outliers
data_nooutliers <- delete_extreme_values(data, c(var_firmvalue_measures, var_control_numeric), threshold = 3)

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

df_descriptive_stats <- calculate_descriptive_stats(data_nooutliers, c(var_firmvalue_measures, var_control_numeric))

# Normality Check

library(e1071) 

calculate_stats <- function(data, variables) {
  results <- data.frame(Variable = character(),
                        Skewness = numeric(),
                        Kurtosis = numeric(),
                        
                        KS_Statistic = numeric(),
                        KS_p_value = numeric(),
                        stringsAsFactors = FALSE)
  
  for (var in variables) {
    if (var %in% names(data)) {
      # Calculate skewness and kurtosis
      skew <- skewness(data[[var]], na.rm = TRUE)
      kurt <- kurtosis(data[[var]], na.rm = TRUE)
      
     
      
      # Perform Kolmogorov-Smirnov test
      ks_test <- ks.test(data[[var]], "pnorm", mean = mean(data[[var]], na.rm = TRUE), sd = sd(data[[var]], na.rm = TRUE))
      
      # Add results to the dataframe
      results <- rbind(results, c(var, skew, kurt, ks_test$statistic, ks_test$p.value))
    } else {
      warning(paste("Variable", var, "not found in the data. Skipping."))
    }
  }
  
  colnames(results) <- c("Variable", "Skewness", "Kurtosis",  "KS_Statistic", "KS_p_value")
  return(results)
}

# Example usage with your data
df_normality_results <- calculate_stats(data_nooutliers, c(var_firmvalue_measures, var_control_numeric))

# Histograms of Numeric vars

library(ggplot2)

generate_histograms <- function(data, vars) {
  for (var in vars) {
    # Generate the histogram plot for each variable
    p <- ggplot(data, aes_string(x = var)) +
      geom_histogram(bins = 30, fill = "blue", color = "black") +
      ggtitle(paste("Histogram of", var)) +
      xlab(var) + 
      ylab("Frequency") +
      theme_minimal()
    
    # Print the plot
    print(p)
    
    # Optionally, save the plot to a file
    ggsave(filename = paste0(var, "_histogram.png"), plot = p, width = 10, height = 8)
  }
}


generate_histograms(data_nooutliers, c(var_firmvalue_measures, var_control_numeric))

# Boxplots

generate_boxplots <- function(data, vars) {
  for (var in vars) {
    # Generate the box plot for each variable
    p <- ggplot(data, aes_string(x = factor(0), y = var)) +  # factor(0) creates a single group
      geom_boxplot(fill = "lightblue", color = "black") +
      ggtitle(paste("Box Plot of", var)) +
      xlab("") +  # Remove x-axis label as it's unnecessary
      ylab(var) +
      theme_minimal()
    
    # Print the plot
    print(p)
    
    # Optionally, save the plot to a file
    ggsave(filename = paste0(var, "_boxplot.png"), plot = p, width = 10, height = 8)
  }
}

# Example usage
generate_boxplots(data_nooutliers, var_firmvalue_measures)


#Log-transformation

varstologtransform <- c("Market.to.book.ratio", "Market.to.Book.value.per.share", "Tobins.Q","Long.Term.Debt_Total.Equity._._",  
                        "R.D.expense", "Foreign.sales.to.total.sales" )

log_transform_dvs <- function(data, dvs) {
  for (dv in dvs) {
    # Creating a new column name for the log-transformed variable
    new_col_name <- paste("log", dv, sep = "_")
    
    # Find the minimum value for the variable
    min_val <- min(data[[dv]], na.rm = TRUE)
    
    # Shift the data so the minimum value becomes greater than zero (e.g., min_val + 1 if min_val is non-negative)
    shift_val <- if(min_val <= 0) abs(min_val) + 1 else 0
    
    # Applying the log transformation with the shift to handle negative and zero values
    data[[new_col_name]] <- log(data[[dv]] + shift_val)
  }
  return(data)
}

# Applying the function to your data
data_nooutliers <- log_transform_dvs(data_nooutliers, varstologtransform)

# Update vars lists

var_firmvalue_measures <- c("log_Market.to.book.ratio", "log_Market.to.Book.value.per.share", "log_Tobins.Q")
var_control_numeric <- c("Ln_Total_Asset","Ln_Revenue", "log_Long.Term.Debt_Total.Equity._._", "ROA._._" , 
                         "log_R.D.expense" , "log_Foreign.sales.to.total.sales" )

# HEDGING PREMIUM ANALYSIS

library(dplyr)
library(broom)

perform_t_tests <- function(data, dependent_vars, group_vars) {
  results <- data.frame()
  
  # Loop over each dependent variable
  for (dv in dependent_vars) {
    # Loop over each grouping variable
    for (gv in group_vars) {
      data[[gv]] <- as.factor(data[[gv]])
      
      if (length(unique(data[[gv]])) == 2) {  # Ensure there are exactly two groups
        # Conduct t-test
        test_result <- t.test(reformulate(gv, response = dv), data = data, var.equal = TRUE)
        
        # Retrieve group names
        group_levels <- levels(data[[gv]])
        
        # Prepare summary statistics and mean/SD for both groups
        summary_stats <- tidy(test_result)
        mean_sd_data <- data %>%
          group_by(!!rlang::sym(gv)) %>%
          summarise(
            Mean = mean(!!rlang::sym(dv), na.rm = TRUE),
            SD = sd(!!rlang::sym(dv), na.rm = TRUE),
            .groups = 'drop'
          )
        
        # Ensure the group order in mean_sd_data matches the levels order
        mean_sd_data <- mean_sd_data %>%
          mutate(Group = factor(!!rlang::sym(gv), levels = group_levels)) %>%
          arrange(Group) %>%
          select(-Group)
        
        # Merge summary statistics and mean/SD for both groups
        stats_row <- data.frame(
          Variable = dv,
          Group_Var = gv,
          Group1 = group_levels[1],
          Group1_Mean = mean_sd_data$Mean[1],
          Group1_SD = mean_sd_data$SD[1],
          Group2 = group_levels[2],
          Group2_Mean = mean_sd_data$Mean[2],
          Group2_SD = mean_sd_data$SD[2],
          T_Value = summary_stats$statistic,
          DF = summary_stats$parameter,
          P_Value = summary_stats$p.value
        )
        
        # Append to results dataframe
        results <- rbind(results, stats_row)
      }
    }
  }
  
  return(results)
}


# Example usage
dependent_vars <- c("Market.to.book.ratio", "Market.to.Book.value.per.share", "Tobins.Q")
group_vars <- c("Use.of.foreign.currency.derivatives", "has_foreign_sales")
df_ttest_results <- perform_t_tests(data_nooutliers, dependent_vars, group_vars)


perform_median_tests <- function(data, dependent_vars, group_vars) {
  results <- data.frame()
  
  # Loop over each dependent variable
  for (dv in dependent_vars) {
    # Loop over each grouping variable
    for (gv in group_vars) {
      data[[gv]] <- as.factor(data[[gv]])
      
      if (length(unique(data[[gv]])) == 2) {  # Ensure there are exactly two groups
        # Conduct Mann-Whitney U test
        test_result <- wilcox.test(reformulate(gv, response = dv), data = data)
        
        # Retrieve group names
        group_levels <- levels(data[[gv]])
        
        # Prepare summary statistics and median/SD for both groups
        summary_stats <- tidy(test_result)
        median_data <- data %>%
          group_by(!!rlang::sym(gv)) %>%
          summarise(
            Median = median(!!rlang::sym(dv), na.rm = TRUE),
            SD = sd(!!rlang::sym(dv), na.rm = TRUE),
            .groups = 'drop'
          )
        
        # Ensure the group order in median_data matches the levels order
        median_data <- median_data %>%
          mutate(Group = factor(!!rlang::sym(gv), levels = group_levels)) %>%
          arrange(Group) %>%
          select(-Group)
        
        # Merge summary statistics and median/SD for both groups
        stats_row <- data.frame(
          Variable = dv,
          Group_Var = gv,
          Group1 = group_levels[1],
          Group1_Median = median_data$Median[1],
          Group1_SD = median_data$SD[1],
          Group2 = group_levels[2],
          Group2_Median = median_data$Median[2],
          Group2_SD = median_data$SD[2],
          W_Statistic = summary_stats$statistic,
          P_Value = summary_stats$p.value
        )
        
        # Append to results dataframe
        results <- rbind(results, stats_row)
      }
    }
  }
  
  return(results)
}

df_median_test_results <- perform_median_tests(data_nooutliers, dependent_vars, group_vars)

# Additional cut for the t-tests

perform_tests_by_foreign_sales <- function(data, dependent_vars, foreign_sales_var, derivatives_var) {
  results <- data.frame()
  
  # Subset data based on foreign sales
  subsets <- split(data, data[[foreign_sales_var]])
  
  for (subset_name in names(subsets)) {
    subset_data <- subsets[[subset_name]]
    subset_label <- ifelse(subset_name == 1, "With Foreign Sales", "Without Foreign Sales")
    
    for (dv in dependent_vars) {
      # Ensure derivatives_var is a factor
      subset_data[[derivatives_var]] <- as.factor(subset_data[[derivatives_var]])
      
      if (length(unique(subset_data[[derivatives_var]])) == 2) {  # Ensure there are exactly two groups
        # Conduct t-test
        t_test_result <- t.test(reformulate(derivatives_var, response = dv), data = subset_data, var.equal = TRUE)
        # Conduct Mann-Whitney U test
        mw_test_result <- wilcox.test(reformulate(derivatives_var, response = dv), data = subset_data)
        
        # Retrieve group names
        group_levels <- levels(subset_data[[derivatives_var]])
        
        # Prepare summary statistics for t-test (mean and SD)
        mean_sd_data <- subset_data %>%
          group_by(!!rlang::sym(derivatives_var)) %>%
          summarise(
            Mean = mean(!!rlang::sym(dv), na.rm = TRUE),
            SD = sd(!!rlang::sym(dv), na.rm = TRUE),
            .groups = 'drop'
          )
        
        # Prepare summary statistics for Mann-Whitney test (median and SD)
        median_sd_data <- subset_data %>%
          group_by(!!rlang::sym(derivatives_var)) %>%
          summarise(
            Median = median(!!rlang::sym(dv), na.rm = TRUE),
            SD = sd(!!rlang::sym(dv), na.rm = TRUE),
            .groups = 'drop'
          )
        
        # Ensure the group order in mean_sd_data and median_sd_data matches the levels order
        mean_sd_data <- mean_sd_data %>%
          mutate(Group = factor(!!rlang::sym(derivatives_var), levels = group_levels)) %>%
          arrange(Group) %>%
          select(-Group)
        
        median_sd_data <- median_sd_data %>%
          mutate(Group = factor(!!rlang::sym(derivatives_var), levels = group_levels)) %>%
          arrange(Group) %>%
          select(-Group)
        
        # Merge summary statistics and test results for both tests
        stats_row <- data.frame(
          Subset = subset_label,
          Variable = dv,
          Group_Var = derivatives_var,
          Group1 = group_levels[1],
          Group1_Mean = mean_sd_data$Mean[1],
          Group1_SD_Mean = mean_sd_data$SD[1],
          Group1_Median = median_sd_data$Median[1],
          Group1_SD_Median = median_sd_data$SD[1],
          Group2 = group_levels[2],
          Group2_Mean = mean_sd_data$Mean[2],
          Group2_SD_Mean = mean_sd_data$SD[2],
          Group2_Median = median_sd_data$Median[2],
          Group2_SD_Median = median_sd_data$SD[2],
          T_Value = tidy(t_test_result)$statistic,
          T_DF = tidy(t_test_result)$parameter,
          T_P_Value = tidy(t_test_result)$p.value,
          W_Statistic = tidy(mw_test_result)$statistic,
          W_P_Value = tidy(mw_test_result)$p.value
        )
        
        # Append to results dataframe
        results <- rbind(results, stats_row)
      }
    }
  }
  
  return(results)
}

# Example usage
dependent_vars <- c("Market.to.book.ratio", "Market.to.Book.value.per.share", "Tobins.Q")
foreign_sales_var <- "has_foreign_sales"
derivatives_var <- "Use.of.foreign.currency.derivatives"

df_ttest_results_twoway <- perform_tests_by_foreign_sales(data_nooutliers, dependent_vars, foreign_sales_var, derivatives_var)


perform_tests_by_foreign_sales_and_appreciation <- function(data, dependent_vars, foreign_sales_var, derivatives_var, appreciation_var) {
  results <- data.frame()
  
  # Helper function to perform t-tests and Mann-Whitney U tests
  perform_tests <- function(data_subset, subset_label, dv, derivatives_var) {
    if (length(unique(data_subset[[derivatives_var]])) == 2) {
      # Conduct t-test
      t_test_result <- t.test(reformulate(derivatives_var, response = dv), data = data_subset, var.equal = TRUE)
      # Conduct Mann-Whitney U test
      mw_test_result <- wilcox.test(reformulate(derivatives_var, response = dv), data = data_subset)
      
      # Retrieve group names
      group_levels <- levels(data_subset[[derivatives_var]])
      
      # Prepare summary statistics for t-test (mean and SD)
      mean_sd_data <- data_subset %>%
        group_by(!!rlang::sym(derivatives_var)) %>%
        summarise(
          Mean = mean(!!rlang::sym(dv), na.rm = TRUE),
          SD = sd(!!rlang::sym(dv), na.rm = TRUE),
          .groups = 'drop'
        )
      
      # Prepare summary statistics for Mann-Whitney test (median and SD)
      median_sd_data <- data_subset %>%
        group_by(!!rlang::sym(derivatives_var)) %>%
        summarise(
          Median = median(!!rlang::sym(dv), na.rm = TRUE),
          SD = sd(!!rlang::sym(dv), na.rm = TRUE),
          .groups = 'drop'
        )
      
      # Ensure the group order in mean_sd_data and median_sd_data matches the levels order
      mean_sd_data <- mean_sd_data %>%
        mutate(Group = factor(!!rlang::sym(derivatives_var), levels = group_levels)) %>%
        arrange(Group) %>%
        select(-Group)
      
      median_sd_data <- median_sd_data %>%
        mutate(Group = factor(!!rlang::sym(derivatives_var), levels = group_levels)) %>%
        arrange(Group) %>%
        select(-Group)
      
      # Merge summary statistics and test results for both tests
      stats_row <- data.frame(
        Subset = subset_label,
        Variable = dv,
        Group_Var = derivatives_var,
        Group1 = group_levels[1],
        Group1_Mean = mean_sd_data$Mean[1],
        Group1_SD_Mean = mean_sd_data$SD[1],
        Group1_Median = median_sd_data$Median[1],
        Group1_SD_Median = median_sd_data$SD[1],
        Group2 = group_levels[2],
        Group2_Mean = mean_sd_data$Mean[2],
        Group2_SD_Mean = mean_sd_data$SD[2],
        Group2_Median = median_sd_data$Median[2],
        Group2_SD_Median = median_sd_data$SD[2],
        T_Value = tidy(t_test_result)$statistic,
        T_DF = tidy(t_test_result)$parameter,
        T_P_Value = tidy(t_test_result)$p.value,
        W_Statistic = tidy(mw_test_result)$statistic,
        W_P_Value = tidy(mw_test_result)$p.value
      )
      
      return(stats_row)
    } else {
      return(NULL)
    }
  }
  
  # Subset data based on foreign sales
  subsets <- split(data, data[[foreign_sales_var]])
  
  for (subset_name in names(subsets)) {
    subset_data <- subsets[[subset_name]]
    foreign_sales_label <- ifelse(subset_name == 1, "With Foreign Sales", "Without Foreign Sales")
    
    for (dv in dependent_vars) {
      # Ensure derivatives_var is a factor
      subset_data[[derivatives_var]] <- as.factor(subset_data[[derivatives_var]])
      
      if (length(unique(subset_data[[derivatives_var]])) == 2) {  # Ensure there are exactly two groups
        # Perform tests for each dollar appreciation level
        for (appreciation in unique(data[[appreciation_var]])) {
          data_appreciation_subset <- subset_data %>% filter(!!rlang::sym(appreciation_var) == appreciation)
          subset_label <- paste(foreign_sales_label, "- Dollar Appreciation", appreciation)
          
          test_result <- perform_tests(data_appreciation_subset, subset_label, dv, derivatives_var)
          if (!is.null(test_result)) {
            results <- rbind(results, test_result)
          }
        }
      }
    }
  }
  
  return(results)
}

# Example usage
dependent_vars <- c("Market.to.book.ratio", "Market.to.Book.value.per.share", "Tobins.Q")
foreign_sales_var <- "has_foreign_sales"
derivatives_var <- "Use.of.foreign.currency.derivatives"
appreciation_var <- "dollar_appreciation"

df_ttest_results_threeway <- perform_tests_by_foreign_sales_and_appreciation(data_nooutliers, dependent_vars, foreign_sales_var, derivatives_var, appreciation_var)

# LMM - Overall

# Create a new dataframe with standardized variables
data_standardized <- data_nooutliers %>%
  mutate(
    Ln_Total_Asset = scale(Ln_Total_Asset, center = TRUE, scale = TRUE),
    Ln_Revenue = scale(Ln_Revenue, center = TRUE, scale = TRUE),
    `log_Long.Term.Debt_Total.Equity._._` = scale(`log_Long.Term.Debt_Total.Equity._._`, center = TRUE, scale = TRUE),
    `ROA._._` = scale(`ROA._._`, center = TRUE, scale = TRUE),
    `R.D.expense` = scale(`R.D.expense`, center = TRUE, scale = TRUE),
    `log_Foreign.sales.to.total.sales` = scale(`log_Foreign.sales.to.total.sales`, center = TRUE, scale = TRUE)
  ) %>% ungroup() # ensure data is ungrouped if it was previously grouped


# Function to plot residuals and fits
plot_resid_fits <- function(model, plot_param) {
  residual_plot_filename <- paste0("Residuals_vs_Fitted_", plot_param, ".jpeg")
  qq_plot_filename <- paste0("QQ_Plot_", plot_param, ".jpeg")
  
  jpeg(residual_plot_filename)
  plot(residuals(model) ~ fitted(model), main = paste("Residuals vs Fitted"), xlab = "Fitted values", ylab = "Residuals")
  dev.off()
  
  jpeg(qq_plot_filename)
  qqnorm(residuals(model), main = "Q-Q Plot")
  qqline(residuals(model))
  dev.off()
}

# Function to fit LMM for the entire dataset
fit_lmm <- function(formula, data, plot_param, save_plots = FALSE) {
  if(nrow(data) == 0) {
    cat("Insufficient data.\n")
    return(NULL)
  }
  
  # Fit the linear mixed-effects model using maximum likelihood
  lmer_model <- lmer(as.formula(formula), data = data, REML = FALSE)
  
  # Summarize the model
  model_summary <- summary(lmer_model)
  cat("Model summary:\n")
  print(model_summary)
  
  # Extract results
  lmm_results <- tidy(lmer_model, effects = "fixed")
  
  # Safely retrieve random effects variances
  random_effects_variances <- tryCatch({
    as.data.frame(VarCorr(lmer_model))
  }, error = function(e) {
    cat("Failed to retrieve random effects variances: ", e$message, "\n")
    return(data.frame(term=character(), grp=character(), vcov=numeric(), sdcor=numeric()))
  })
  
  # Combine fixed and random effects results
  combined_results <- bind_rows(
    lmm_results %>% mutate(effect = "fixed"),
    random_effects_variances %>% mutate(effect = "random")
  )
  
  # Extract model fit indices
  fit_indices <- data.frame(
    AIC = AIC(lmer_model),
    BIC = BIC(lmer_model),
    LogLik = as.numeric(logLik(lmer_model)),
    N = nobs(lmer_model)
  )
  
  # Optionally save plots
  if (save_plots) {
    plot_resid_fits(lmer_model, plot_param)
  }
  
  # Return results and fit indices as a single dataframe
  combined_data <- bind_rows(
    combined_results,
    fit_indices %>% mutate(term = "Fit Indices", effect = "fit")
  )
  
  return(combined_data)
}

# Example function calls
plot_param = "MarketToBook"
df_lmm_results_markettobook <- fit_lmm("Market.to.book.ratio ~ Use.of.foreign.currency.derivatives + Ln_Total_Asset + Ln_Revenue + log_Long.Term.Debt_Total.Equity._._ + ROA._._ + log_R.D.expense + log_Foreign.sales.to.total.sales + Dividend  + (1|Year) + (1|Ticker)", data_standardized, plot_param, save_plots = TRUE)

plot_param = "MarketToBookPerShare"
df_lmm_results_markettobookpershare <- fit_lmm("Market.to.Book.value.per.share ~ Use.of.foreign.currency.derivatives + Ln_Total_Asset + Ln_Revenue + log_Long.Term.Debt_Total.Equity._._ + ROA._._ + log_R.D.expense + log_Foreign.sales.to.total.sales + Dividend + (1|Year) + (1|Ticker)", data_standardized, plot_param, save_plots = TRUE)

plot_param = "TobinsQ"
df_lmm_results_tobinsQ <- fit_lmm("Tobins.Q ~ Use.of.foreign.currency.derivatives + Ln_Total_Asset + Ln_Revenue + log_Long.Term.Debt_Total.Equity._._ + ROA._._ + log_R.D.expense + log_Foreign.sales.to.total.sales + Dividend + (1|Year) + (1|Ticker)", data_standardized, plot_param, save_plots = TRUE)

colnames(data_nooutliers)

# LMM by Year

library(lme4)
library(broom.mixed)
library(car)
library(dplyr)

# Function to fit LMM for a given year and return results along with fit indices
plot_resid_fits <- function(model, plot_param, year) {
  residual_plot_filename <- paste0("Residuals_vs_Fitted_", plot_param, "_", year, ".jpeg")
  qq_plot_filename <- paste0("QQ_Plot_", plot_param, "_", year, ".jpeg")
  
  jpeg(residual_plot_filename)
  plot(residuals(model) ~ fitted(model), main = paste("Residuals vs Fitted -", year), xlab = "Fitted values", ylab = "Residuals")
  dev.off()
  
  jpeg(qq_plot_filename)
  qqnorm(residuals(model), main = paste("Q-Q Plot -", year))
  qqline(residuals(model))
  dev.off()
}


# Function to fit LMM for a given year and return results
fit_lmm_and_format_byyear <- function(formula, data, year, save_plots = FALSE, plot_param = "") {
  if(nrow(data) <= 1 || length(unique(data$Industry.Code)) <= 1) {
    cat("Skipping model for year", year, " due to insufficient data.\n")
    return(NULL)
  }
  
  # Fit the linear mixed-effects model using maximum likelihood
  lmer_model <- lmer(as.formula(formula), data = data, REML = FALSE)
  
  # Summarize the model
  model_summary <- summary(lmer_model)
  cat("Model summary for year", year, ":\n")
  print(model_summary)
  
  # Extract results
  lmm_results <- tidy(lmer_model, effects = "fixed")
  
  # Extract random effects variances
  random_effects_variances <- tryCatch({
    as.data.frame(VarCorr(lmer_model))
  }, error = function(e) {
    cat("Failed to retrieve random effects variances for year", year, ": ", e$message, "\n")
    return(data.frame(term=character(), grp=character(), vcov=numeric(), sdcor=numeric()))
  })
  
  # Combine fixed and random effects results
  combined_results <- bind_rows(
    lmm_results %>% mutate(effect = "fixed"),
    random_effects_variances %>% mutate(effect = "random")
  ) %>% mutate(Year = year)
  
  # Extract model fit indices
  fit_indices <- data.frame(
    AIC = AIC(lmer_model),
    BIC = BIC(lmer_model),
    LogLik = as.numeric(logLik(lmer_model)),
    N = nobs(lmer_model),  # Number of observations
    Year = year
  )
  
  # Optionally save plots
  if (save_plots) {
    plot_resid_fits(lmer_model, plot_param, year)
  }
  
  # Return results and fit indices as a single dataframe
  combined_data <- bind_rows(
    combined_results,
    fit_indices %>% mutate(term = "Fit Indices", effect = "fit")
  )
  
  return(combined_data)
}


# Function to run models per year
run_lmm_per_year <- function(formula, data, dependent_var, save_plots = FALSE) {
  results <- data.frame()
  
  for (year in unique(data$Year)) {
    data_year <- data %>% filter(Year == year)
    plot_param <- paste0(dependent_var, "_", year)
    year_results <- fit_lmm_and_format_byyear(formula, data_year, year, save_plots, plot_param)
    if (!is.null(year_results)) {
      year_results <- year_results %>%
        mutate(Dependent_Var = dependent_var)
      results <- bind_rows(results, year_results)
    }
  }
  
  return(results)
}

formula_markettobook <- "Market.to.book.ratio ~ Use.of.foreign.currency.derivatives + Ln_Total_Asset + Ln_Revenue + log_Long.Term.Debt_Total.Equity._._ + ROA._._ + log_Foreign.sales.to.total.sales + Dividend + (1|Industry.Code)"
df_lmm_results_markettobook_year <- run_lmm_per_year(formula_markettobook,data_standardized, "Market.to.book.ratio", save_plots = TRUE)

formula_markettobookpershare <- "Market.to.Book.value.per.share ~ Use.of.foreign.currency.derivatives + Ln_Total_Asset + Ln_Revenue + log_Long.Term.Debt_Total.Equity._._ + ROA._._ +  log_Foreign.sales.to.total.sales + Dividend  + (1|Industry.Code)"

df_lmm_results_markettobookpershare_year <- run_lmm_per_year(formula_markettobookpershare, data_standardized, "Market.to.Book.value.per.share", save_plots = TRUE)

formula_tobinsQ <- "Tobins.Q ~ Use.of.foreign.currency.derivatives + Ln_Total_Asset + Ln_Revenue + log_Long.Term.Debt_Total.Equity._._ + ROA._._ + log_Foreign.sales.to.total.sales + Dividend  + (1|Industry.Code)"

df_lmm_results_tobinsQ_year <- run_lmm_per_year(formula_tobinsQ, data_standardized, "Tobins.Q", save_plots = TRUE)


# Frequency tables

data_ticker <- data_nooutliers %>%
  group_by(Ticker) %>%
  slice(n()) %>%
  ungroup()


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


df_freq_segmented <- create_segmented_frequency_tables(data_ticker, var_use_derivatives, var_foreign_operations)

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


df_freq_derivatives <- create_frequency_tables(data_ticker, var_use_derivatives)
df_freq_foreignops <- create_frequency_tables(data_ticker, var_foreign_operations)

# Logit Mixed Model

# Model

library(lme4)
library(broom.mixed)
library(MuMIn)
library(caret)
library(pROC)


fit_mixed_effects_logistic <- function(formula, data) {
  # Fit the mixed-effects logistic regression model
  control_settings <- glmerControl(
    optimizer = "bobyqa",       # Use the BOBYQA optimizer; it's good for mixed models
    optCtrl = list(maxfun = 1e5)  # Increase the maximum number of function evaluations
  )
  
  logistic_model_mixed <- glmer(formula, family = binomial, data = data, control = control_settings, nAGQ = 1)  # Ensure higher accuracy with nAGQ > 1 if needed
  
  # Get tidy output for the mixed-effects logistic regression model
  model_results <- tidy(logistic_model_mixed)
  
  # Calculate AIC and BIC
  aic_value <- AIC(logistic_model_mixed)
  bic_value <- BIC(logistic_model_mixed)
  
  # Calculate Log-Likelihood and Deviance
  log_likelihood <- logLik(logistic_model_mixed)
  deviance_value <- deviance(logistic_model_mixed)
  
  # Calculate the number of observations
  number_of_observations <- nobs(logistic_model_mixed)
  
  # Create a dataframe for fit indices
  fit_indices <- data.frame(
    AIC = aic_value,
    BIC = bic_value,
    Log_Likelihood = as.numeric(log_likelihood),
    Deviance = deviance_value,
    N = number_of_observations
  )
  
  # Return results and fit indices
  return(list(Model_Results = model_results, Fit_Indices = fit_indices))
}

# Example usage
formula <- Use.of.foreign.currency.derivatives ~ Ln_Total_Asset + Ln_Revenue + log_Long.Term.Debt_Total.Equity._._ + ROA._._ + R.D.expense + log_Foreign.sales.to.total.sales + Dividend + (1|Ticker) + (1|Year)
result <- fit_mixed_effects_logistic(formula, data_standardized)

df_mlogitmodel_params <- result$Model_Results
df_mlogitmodel_fit <- result$Fit_Indices

additional_vars <- c("Dollar.Index.Volatility._FX.volatility_",
                     "Cash.flow.from.operations", "Current.Asset"   ,                       
                     "Current.Liabilities"      ,               "Current.Asset.to.Current.Liabilities" ,   "EBIT", "Interest.Coverage.Ratio",
                     "Foreign.sales.to.total.sales", "Geographic.diversification."    ,         "SGA.to.Total.Asset")

library(moments)

convert_selected_to_numeric <- function(data_frame, target_columns) {
  for (col_name in target_columns) {
    if (col_name %in% names(data_frame)) {
      # Check if the column is not numeric
      if (!is.numeric(data_frame[[col_name]])) {
        # Check if the column is a factor or character
        if (is.factor(data_frame[[col_name]]) || is.character(data_frame[[col_name]])) {
          # Convert factor or character to numeric
          # This line handles factors by converting to character first, then to numeric
          numeric_values <- as.numeric(as.character(data_frame[[col_name]]))
          # Check if the conversion introduced NAs due to invalid characters
          if (any(is.na(numeric_values))) {
            warning(paste("NA introduced by coercion in column:", col_name))
          }
          data_frame[[col_name]] <- numeric_values
        } else {
          warning(paste("Column", col_name, "is not a character or factor and has not been converted."))
        }
      }
    } else {
      warning(paste("Column", col_name, "does not exist in the data frame."))
    }
  }
  return(data_frame)
}


data_standardized_numeric <- convert_selected_to_numeric(data_standardized, additional_vars)

# Now you can run your calculate_descriptive_stats function
df_descriptive_newvars <- calculate_descriptive_stats(data_standardized_numeric, additional_vars)


log_transform_dvs <- function(data, dvs) {
  for (dv in dvs) {
    # Creating a new column name for the log-transformed variable
    new_col_name <- paste("log", dv, sep = "_")
    
    # Find the minimum value for the variable
    min_val <- min(data[[dv]], na.rm = TRUE)
    
    # Shift the data so the minimum value becomes greater than zero (e.g., min_val + 1 if min_val is non-negative)
    shift_val <- if(min_val <= 0) abs(min_val) + 1 else 0
    
    # Applying the log transformation with the shift to handle negative and zero values
    data[[new_col_name]] <- log(data[[dv]] + shift_val)
  }
  return(data)
}

varstologtransform2 <- c(
                         "Cash.flow.from.operations", "Current.Asset"   ,                       
                         "Current.Liabilities"      ,               "Current.Asset.to.Current.Liabilities" ,   "EBIT", "Interest.Coverage.Ratio",
                         "Geographic.diversification."    ,         "SGA.to.Total.Asset")

# Applying the function to your data
data_standardized_numeric <- log_transform_dvs(data_standardized_numeric, varstologtransform2)




# Updated formula including new variables
formula <- Use.of.foreign.currency.derivatives ~ log_Market.to.book.ratio + Ln_Total_Asset + Ln_Revenue + 
  log_Long.Term.Debt_Total.Equity._._ + ROA._._ + R.D.expense + 
  log_Foreign.sales.to.total.sales + Dividend + log_Current.Asset.to.Current.Liabilities +
   
  `log_EBIT` + 
  `log_Interest.Coverage.Ratio` + `log_Geographic.diversification.` + 
  `log_SGA.to.Total.Asset` + `Dollar.Index.Volatility._FX.volatility_` + 
  (1|Ticker) + (1|Year)

colnames(data_standardized_numeric)
# Adding Check for Multicolliearity
library(lme4)  # for glmer
library(car)   # for vif
library(broom) # for tidy

fit_mixed_effects_logistic <- function(formula, data) {
  # Fit the mixed-effects logistic regression model
  control_settings <- glmerControl(
    optimizer = "bobyqa",       # Use the BOBYQA optimizer; it's good for mixed models
    optCtrl = list(maxfun = 1e5)  # Increase the maximum number of function evaluations
  )
  
  logistic_model_mixed <- glmer(formula, family = binomial, data = data, control = control_settings, nAGQ = 1)  # Ensure higher accuracy with nAGQ > 1 if needed
  
  # Calculate VIF for fixed effects - First, fit a linear model with the fixed effects only
  fixed_effects_formula <- as.formula(paste("scale(", deparse(formula[[2]]), ") ~ ."))
  lm_model_for_vif <- lm(fixed_effects_formula, data = data)
  
  
  # Get tidy output for the mixed-effects logistic regression model
  model_results <- tidy(logistic_model_mixed)
  
  # Calculate AIC and BIC
  aic_value <- AIC(logistic_model_mixed)
  bic_value <- BIC(logistic_model_mixed)
  
  # Calculate Log-Likelihood and Deviance
  log_likelihood <- logLik(logistic_model_mixed)
  deviance_value <- deviance(logistic_model_mixed)
  
  # Calculate the number of observations
  number_of_observations <- nobs(logistic_model_mixed)
  
  # Create a dataframe for fit indices
  fit_indices <- data.frame(
    AIC = aic_value,
    BIC = bic_value,
    Log_Likelihood = as.numeric(log_likelihood),
    Deviance = deviance_value,
    N = number_of_observations
  )
  
  # Return results and fit indices
  return(list(Model_Results = model_results, Fit_Indices = fit_indices))
}

library(dplyr)

# Create a new dataframe with standardized variables
data_standardized_numeric <- data_standardized_numeric  %>%
  mutate(
    Dollar.Index.Volatility._FX.volatility_ = scale(Dollar.Index.Volatility._FX.volatility_, center = TRUE, scale = TRUE),
    EBIT = scale(EBIT, center = TRUE, scale = TRUE),
    Current.Asset.to.Current.Liabilities = scale(Current.Asset.to.Current.Liabilities, center = TRUE, scale = TRUE),
    Interest.Coverage.Ratio = scale(Interest.Coverage.Ratio, center = TRUE, scale = TRUE),
    SGA.to.Total.Asset = scale(SGA.to.Total.Asset, center = TRUE, scale = TRUE)
  ) %>% ungroup() # ensure data is ungrouped if it was previously grouped

varstologtransform2 <- c(
  "Cash.flow.from.operations", "Current.Asset"   ,                       
  "Current.Liabilities"      ,               "Current.Asset.to.Current.Liabilities" ,   "EBIT", "Interest.Coverage.Ratio",
  "Geographic.diversification."    ,         "SGA.to.Total.Asset")

# Applying the function to your data
data_standardized_numeric <- log_transform_dvs(data_standardized_numeric, varstologtransform2)



# Fit the mixed effects logistic regression model
result_newlogit <- fit_mixed_effects_logistic(formula, data_standardized_numeric)

df_mlogitmodel_params2 <- result_newlogit$Model_Results
df_mlogitmodel_fit2 <- result_newlogit$Fit_Indices

# Example usage
data_list <- list(
  "Correlation" = correlation_matrix,
  "new Logit Model" = df_mlogitmodel_params2,
  "new model fit" = df_mlogitmodel_fit2
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_NewLogit.xlsx")


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

vars <- c("log_Market.to.book.ratio", "Market.to.Book.value.per.share", "Tobins.Q", 
          "Ln_Total_Asset", "Ln_Revenue", "log_Long.Term.Debt_Total.Equity._._", 
          "ROA._._", "R.D.expense", "log_Foreign.sales.to.total.sales", "Dividend", 
          "log_Cash.flow.from.operations", "log_Current.Asset", "log_Current.Liabilities", 
          "log_Current.Asset.to.Current.Liabilities", "log_EBIT", 
          "log_Interest.Coverage.Ratio", "log_Geographic.diversification.", 
          "log_SGA.to.Total.Asset", "Dollar.Index.Volatility._FX.volatility_", 
          "Foreign.sales.to.total.sales")

correlation_matrix <- calculate_correlation_matrix(data_standardized_numeric, vars, method = "pearson")


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
    
    # Continue only if data_to_write is not empty
    if (is.null(data_to_write) || nrow(data_to_write) == 0 || ncol(data_to_write) == 0) {
      warning(paste("Data for", sheet_name, "is empty or not available. Skipping this sheet."))
      next  # Skip this iteration if the data is empty
    }
    
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
  "Final Data" = data_nooutliers,
  "Freq Table 1" = df_freq_foreignops,
  "Freq Table 2" = df_freq_derivatives,
  "Freq Table 3" = df_freq_segmented,
  "Descriptives" = df_descriptive_stats,
  "Normality" = df_normality_results,
  "Ttest 1" = df_test_results,
  "Ttest 12" = df_test_results_foreignsales,
  "Ttest 2" = df_ttest_results,
  "Ttest 3" = df_ttest_results_twoway,
  "Ttest 4" = df_ttest_results_threeway,
  "Median Comp" = df_median_test_results,
  "Model - Market to Book" = df_lmm_results_markettobook,
  "Model - Per Share" = df_lmm_results_markettobookpershare,
  "Model - TobinsQ" = df_lmm_results_tobinsQ,
  "Models by Year 1" = df_lmm_results_markettobook_year,
  "Models by Year 2" = df_lmm_results_markettobookpershare_year,
  "Models by Year 3" = df_lmm_results_tobinsQ_year,
  "Logit Model" = df_mlogitmodel_params,
  "Logit - Fit" = df_mlogitmodel_fit
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

# Example usage
data_list <- list(
  "Outlier Evaluation" = extreme_value_counts
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_outlier.xlsx")