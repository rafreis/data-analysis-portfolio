setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/joshuag10")

library(openxlsx)
df <- read.xlsx("Public Takeovers Switzerland_BA JG_v3.xlsx", sheet = 'CleanData')

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

colnames(df)


# Boxplots
library(ggplot2)

# Assuming your data frame is named 'df'
library(ggplot2)
library(reshape2)

# Melting the data frame to long format for differences only
diff_data <- melt(df, measure.vars = c("DCF.Valuation._per.share__Diff..Price.Low.to.Offer.Price",
                                       "DCF.Valuation._per.share__Diff..Price.High.to.Offer.Price",
                                       "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
                                       "LTM.Transaction.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
                                       "LTM.Transaction.Multiple._per.share__Diff..Price.High.to.Offer.Price",
                                       "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                                       "NTM.Trading.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
                                       "NTM.Trading.Multiple._per.share__Diff..Price.High.to.Offer.Price",
                                       "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                                       "Analyst.Prices._per.share__Diff..Price.Low.to.Offer.Price",
                                       "Analyst.Prices._per.share__Diff..Price.High.to.Offer.Price",
                                       "Analyst.Prices._per.share__Diff..Price.Median.to.Offer.Price"))

diff_vars <- c("DCF.Valuation._per.share__Diff..Price.Low.to.Offer.Price",
               "DCF.Valuation._per.share__Diff..Price.High.to.Offer.Price",
               "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
               "LTM.Transaction.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
               "LTM.Transaction.Multiple._per.share__Diff..Price.High.to.Offer.Price",
               "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
               "NTM.Trading.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
               "NTM.Trading.Multiple._per.share__Diff..Price.High.to.Offer.Price",
               "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
               "Analyst.Prices._per.share__Diff..Price.Low.to.Offer.Price",
               "Analyst.Prices._per.share__Diff..Price.High.to.Offer.Price",
               "Analyst.Prices._per.share__Diff..Price.Median.to.Offer.Price")

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

df_descriptive_stats <- calculate_descriptive_stats(df, diff_vars)


# Box Plot
library(ggplot2)
library(reshape2)
library(dplyr)

# Melting the data frame to long format for differences only, with cleaner labels
diff_data <- melt(df, measure.vars = c(
  "DCF.Valuation._per.share__Diff..Price.Low.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
  "NTM.Trading.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
  "Analyst.Prices._per.share__Diff..Price.Low.to.Offer.Price",
  "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
  "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
  "Analyst.Prices._per.share__Diff..Price.Median.to.Offer.Price",
  "DCF.Valuation._per.share__Diff..Price.High.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.High.to.Offer.Price",
  "NTM.Trading.Multiple._per.share__Diff..Price.High.to.Offer.Price",
  "Analyst.Prices._per.share__Diff..Price.High.to.Offer.Price"
)) %>%
  mutate(Method = gsub(".*(DCF|LTM|NTM|Analyst).*", "\\1", variable),
         Error_Type = ifelse(grepl("Low", variable), "Error - Low Range",
                             ifelse(grepl("Median", variable), "Error - Median", "Error - High Range")),
         Error_Type = factor(Error_Type, levels = c("Error - Low Range", "Error - Median", "Error - High Range")))

# Plotting
ggplot(diff_data, aes(x = Error_Type, y = value, fill = Method)) +
  geom_boxplot() +
  labs(title = "Box Plot of Prediction Differences by Error Type and Method", x = "Error Type", y = "Difference to Offer Price") +
  scale_fill_brewer(palette = "Set1") +  # Using a distinct color palette
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.title = element_blank())

# Bar Plots

# Melting the data frame to long format for differences only
bar_data <- melt(df, measure.vars = c(
  "DCF.Valuation._per.share__Diff..Price.Low.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
  "NTM.Trading.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
  "Analyst.Prices._per.share__Diff..Price.Low.to.Offer.Price",
  "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
  "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
  "Analyst.Prices._per.share__Diff..Price.Median.to.Offer.Price",
  "DCF.Valuation._per.share__Diff..Price.High.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.High.to.Offer.Price",
  "NTM.Trading.Multiple._per.share__Diff..Price.High.to.Offer.Price",
  "Analyst.Prices._per.share__Diff..Price.High.to.Offer.Price"
)) %>%
  mutate(Method = gsub(".*(DCF|LTM|NTM|Analyst).*", "\\1", variable),
         Error_Type = ifelse(grepl("Low", variable), "Error - Low Range",
                             ifelse(grepl("Median", variable), "Error - Median", "Error - High Range")),
         Error_Type = factor(Error_Type, levels = c("Error - Low Range", "Error - Median", "Error - High Range")))

dodge <- position_dodge(width = 0.9)  # Set the dodge width

# Bar Plot showing Mean and 95% CI with data labels
ggplot(bar_data, aes(x = Error_Type, y = value, fill = Method)) +
  stat_summary(fun.data = mean_cl_normal, geom = "errorbar", width = 0.2, position = dodge) +
  stat_summary(fun = mean, geom = "bar", position = dodge, width = 0.7) +
  stat_summary(fun = mean, geom = "text", position = dodge, vjust = -0.5, aes(label = sprintf("%.1f", ..y..))) +
  labs(title = "Mean and 95% CI of Valuation Prediction Errors by Method", x = "Error Type", y = "Difference to Offer Price") +
  scale_fill_brewer(palette = "Set1") +  # Using a distinct color palette
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1),
        legend.title = element_blank())

# Paired-Samples T-test
library(dplyr)

# Function to calculate Cohen's d
run_paired_t_tests <- function(data, method1_col_low, method2_col_low, method1_col_median, method2_col_median, method1_col_high, method2_col_high) {
  results <- data.frame(
    Comparison = character(),
    T_Value = numeric(),
    P_Value = numeric(),
    Effect_Size = numeric(),
    Num_Pairs = integer(),
    stringsAsFactors = FALSE
  )
  
  # Test configurations
  test_configs <- list(
    Low = c(method1_col_low, method2_col_low),
    Median = c(method1_col_median, method2_col_median),
    High = c(method1_col_high, method2_col_high)
  )
  
  # Iterating over each configuration
  for (type in names(test_configs)) {
    var1 <- test_configs[[type]][1]
    var2 <- test_configs[[type]][2]
    
    # Filtering out NA values to ensure valid pairs
    valid_data <- data %>%
      select(all_of(var1), all_of(var2)) %>%
      na.omit()
    
    if (nrow(valid_data) > 0) {
      # Running paired t-test
      t_test_result <- t.test(valid_data[[var1]], valid_data[[var2]], paired = TRUE)
      
      # Calculating Cohen's d
      mean_diff <- mean(valid_data[[var2]]) - mean(valid_data[[var1]])
      pooled_sd <- sd(c(valid_data[[var1]], valid_data[[var2]]))
      effect_size <- mean_diff / pooled_sd
      
      results <- rbind(results, data.frame(
        Comparison = paste(var1, "vs", var2, ":", type),
        T_Value = t_test_result$statistic,
        P_Value = t_test_result$p.value,
        Effect_Size = effect_size,
        Num_Pairs = nrow(valid_data)
      ))
    }
  }
  
  return(results)
}

# Example usage with "DCF" compared to "LTM", "NTM", and "Analyst"
results_dcf_vs_ltm <- run_paired_t_tests(
  df,
  "DCF.Valuation._per.share__Diff..Price.Low.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
  "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
  "DCF.Valuation._per.share__Diff..Price.High.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Diff..Price.High.to.Offer.Price"
)
results_dcf_ntm <- run_paired_t_tests(df, "DCF.Valuation._per.share__Diff..Price.Low.to.Offer.Price",
                                      "NTM.Trading.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
                                      "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
                                      "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                                      "DCF.Valuation._per.share__Diff..Price.High.to.Offer.Price",
                                      "NTM.Trading.Multiple._per.share__Diff..Price.High.to.Offer.Price")
results_dcf_analyst <- run_paired_t_tests(df, "DCF.Valuation._per.share__Diff..Price.Low.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.Low.to.Offer.Price",
                                          "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.Median.to.Offer.Price",
                                          "DCF.Valuation._per.share__Diff..Price.High.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.High.to.Offer.Price")

results_ntm_ltm <- run_paired_t_tests(df,"LTM.Transaction.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
                                      "NTM.Trading.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
                                      "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                                      "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                                      "LTM.Transaction.Multiple._per.share__Diff..Price.High.to.Offer.Price",
                                      "NTM.Trading.Multiple._per.share__Diff..Price.High.to.Offer.Price")

results_ntm_analyst <- run_paired_t_tests(df,"Analyst.Prices._per.share__Diff..Price.Low.to.Offer.Price",
                                          "NTM.Trading.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.Median.to.Offer.Price",
                                          "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.High.to.Offer.Price",
                                          "NTM.Trading.Multiple._per.share__Diff..Price.High.to.Offer.Price")

results_ltm_analyst <- run_paired_t_tests(df, "LTM.Transaction.Multiple._per.share__Diff..Price.Low.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.Low.to.Offer.Price",
                                          "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.Median.to.Offer.Price",
                                          "LTM.Transaction.Multiple._per.share__Diff..Price.High.to.Offer.Price",
                                          "Analyst.Prices._per.share__Diff..Price.High.to.Offer.Price")


# Combine all results into a single data frame
all_results <- rbind(results_dcf_vs_ltm,
                     results_dcf_ntm,
                     results_dcf_analyst,
                     results_ntm_ltm,
                     results_ntm_analyst,
                     results_ltm_analyst)


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
  "T-test Results" = all_results,
  "Descriptives" = df_descriptive_stats
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")




# New Tests

vars_valuation <- c("DCF.Valuation._per.share__Price.Low"   ,                                
               "DCF.Valuation._per.share__Price.High"   ,                               
               "DCF.Valuation._per.share__Median"  ,
               "LTM.Transaction.Multiple._per.share__Price.Low"   ,                     
               "LTM.Transaction.Multiple._per.share__Price.High"   ,                    
               "LTM.Transaction.Multiple._per.share__Median" ,
               "NTM.Trading.Multiple._per.share__Price.Low"   ,                         
               "NTM.Trading.Multiple._per.share__Price.High"   ,                        
               "NTM.Trading.Multiple._per.share__Median" ,
               "Analyst.Prices._per.share__Price.Low"     ,                             
               "Analyst.Prices._per.share__Price.High"     ,                            
               "Analyst.Prices._per.share__Median" )

# Calculate percentage differences
for (col in diff_vars) {
  percentage_col_name <- sub("Diff..Price", "Perc_Diff", col)
  df[[percentage_col_name]] <- (df[[col]] / df$Offer.Price._CHF_) * 100
}

vars_valuation_perc <- c("DCF.Valuation._per.share__Perc_Diff.Low.to.Offer.Price"        ,        
                          "DCF.Valuation._per.share__Perc_Diff.High.to.Offer.Price"       ,        
                         "DCF.Valuation._per.share__Perc_Diff.Median.to.Offer.Price"       ,      
                         "LTM.Transaction.Multiple._per.share__Perc_Diff.Low.to.Offer.Price",     
                         "LTM.Transaction.Multiple._per.share__Perc_Diff.High.to.Offer.Price",    
                         "LTM.Transaction.Multiple._per.share__Perc_Diff.Median.to.Offer.Price" , 
                         "NTM.Trading.Multiple._per.share__Perc_Diff.Low.to.Offer.Price"  ,       
                          "NTM.Trading.Multiple._per.share__Perc_Diff.High.to.Offer.Price" ,       
                         "NTM.Trading.Multiple._per.share__Perc_Diff.Median.to.Offer.Price" ,     
                         "Analyst.Prices._per.share__Perc_Diff.Low.to.Offer.Price"           ,    
                         "Analyst.Prices._per.share__Perc_Diff.High.to.Offer.Price"           ,   
                         "Analyst.Prices._per.share__Perc_Diff.Median.to.Offer.Price")


desc_stats_perc <- calculate_descriptive_stats(df, vars_valuation_perc)

library(dplyr)

# Function to run paired t-tests for given columns against the Offer Price
run_paired_t_tests <- function(data, offer_price_col, measurement_cols) {
  results <- data.frame(
    Variable = character(),
    T_Value = numeric(),
    P_Value = numeric(),
    Effect_Size = numeric(),
    Num_Pairs = integer(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each measurement column
  for (var in measurement_cols) {
    # Prepare data, removing NA values
    valid_data <- data %>%
      select(offer_price_col, var) %>%
      na.omit()
    
    num_pairs <- nrow(valid_data)
    
    # Only perform the test if we have valid pairs
    if (num_pairs > 0) {
      t_test_result <- t.test(valid_data[[var]], valid_data[[offer_price_col]], paired = TRUE)
      # Calculating Cohen's d for effect size
      mean_diff <- mean(valid_data[[var]], na.rm = TRUE) - mean(valid_data[[offer_price_col]], na.rm = TRUE)
      pooled_sd <- sd(c(valid_data[[var]], valid_data[[offer_price_col]]), na.rm = TRUE)
      effect_size <- mean_diff / pooled_sd
      
      results <- rbind(results, data.frame(
        Variable = var,
        T_Value = t_test_result$statistic,
        P_Value = t_test_result$p.value,
        Effect_Size = effect_size,
        Num_Pairs = num_pairs
      ))
    }
  }
  
  return(results)
}

colnames(df)

# Example usage
vars_valuation <- c(
  "DCF.Valuation._per.share__Price.Low",                                   
  "DCF.Valuation._per.share__Price.High"   ,                               
  "DCF.Valuation._per.share__Median",     
  "LTM.Transaction.Multiple._per.share__Price.Low"  ,                      
  "LTM.Transaction.Multiple._per.share__Price.High",                       
  "LTM.Transaction.Multiple._per.share__Median"  ,
  "NTM.Trading.Multiple._per.share__Price.Low" ,                           
  "NTM.Trading.Multiple._per.share__Price.High" ,                          
  "NTM.Trading.Multiple._per.share__Median"     ,
  "Analyst.Prices._per.share__Price.Low"    ,                              
  "Analyst.Prices._per.share__Price.High"    ,                             
  "Analyst.Prices._per.share__Median",
  "Offer.Price._CHF_"
)

# Assuming 'Offer_Price._CHF_' is the column name in your dataset
paired_ttest_results <- run_paired_t_tests(df, "Offer.Price._CHF_", vars_valuation)

desc_stats <- calculate_descriptive_stats(df, vars_valuation)

# Example usage
data_list <- list(
  "Paired T-test Results" = paired_ttest_results,
  "Descriptives - Perc" = desc_stats_perc,
  "Descriptives - Raw" = desc_stats
  )

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_v2.xlsx")



# Histograms

library(ggplot2)
library(dplyr)
library(tidyr)

# Assuming df is already loaded and contains the necessary columns
vars_valuation_perc <- c(
  "DCF.Valuation._per.share__Perc_Diff.Low.to.Offer.Price",        
  "DCF.Valuation._per.share__Perc_Diff.High.to.Offer.Price",   
 
  
  "LTM.Transaction.Multiple._per.share__Perc_Diff.Low.to.Offer.Price",     
  "LTM.Transaction.Multiple._per.share__Perc_Diff.High.to.Offer.Price",    
  "NTM.Trading.Multiple._per.share__Perc_Diff.Low.to.Offer.Price",       
  "NTM.Trading.Multiple._per.share__Perc_Diff.High.to.Offer.Price"     
  
)
colnames(df)

# Pivot the data to long format for easier plotting
long_data <- df %>%
  pivot_longer(cols = all_of(vars_valuation_perc), names_to = "Variable", values_to = "Percentage_Error")

# Manually specify labels for the legend
legend_labels <- c(
  "DCF.Valuation._per.share__Perc_Diff.Low.to.Offer.Price" = "Low",
  "DCF.Valuation._per.share__Perc_Diff.High.to.Offer.Price" = "High",
  "LTM.Transaction.Multiple._per.share__Perc_Diff.Low.to.Offer.Price" = "Low",
  "LTM.Transaction.Multiple._per.share__Perc_Diff.High.to.Offer.Price" = "High",
  "NTM.Trading.Multiple._per.share__Perc_Diff.Low.to.Offer.Price" = "Low",
  "NTM.Trading.Multiple._per.share__Perc_Diff.High.to.Offer.Price" = "High"
)

# Function to generate histogram for each method
plot_histogram <- function(data, method) {
  filtered_data <- data %>% filter(grepl(method, Variable))
  ggplot(filtered_data, aes(x = Percentage_Error, fill = Variable)) +
    geom_histogram(bins = 30, alpha = 0.6, position = "identity") +
    scale_fill_manual(values = c("skyblue", "orange"), labels = legend_labels) +
    labs(title = paste("Distribution of Percentage Errors -", method),
         x = "Percentage Error (%)",
         y = "Frequency",
         fill = "Error Type") +
    theme_minimal() +
    theme(legend.title = element_blank(),
          legend.position = "bottom",
          legend.text = element_text(size = 8))
}

# Separate plots for each method
methods <- c("DCF", "LTM", "NTM")

# Generate and print plots for each method
plots <- list()
for (method in methods) {
  plots[[method]] <- plot_histogram(long_data, method)
}

# Display or save plots
print(plots$DCF)
print(plots$LTM)
print(plots$NTM)

library(ggplot2)
library(dplyr)



plot_histogram <- function(data, method) {
  filtered_data <- data %>%
    filter(grepl(method, Variable))  # Ensures only data relevant to the 'method' is used
  
  ggplot(filtered_data, aes(x = Percentage_Error, fill = Variable)) +
    geom_histogram(bins = 30, alpha = 0.6, position = "identity") +
    scale_fill_manual(values = c("skyblue", "orange", "lightgreen")) +  # Add a color for 'Mid' if applicable
    labs(title = paste("Distribution of Percentage Errors -", method),
         x = "Percentage Error (%)",
         y = "Frequency",
         fill = "Error Type") +
    theme_minimal() +
    theme(legend.title = element_blank(),
          legend.position = "bottom",
          legend.text = element_text(size = 8, family = "Arial"),  # Arial font
          text = element_text(family = "Arial"),  # Set all text to Arial
          panel.grid.major = element_blank(),  # Remove major grid lines
          panel.grid.minor = element_blank(),  # Remove minor grid lines
          axis.text.x = element_text(margin = margin(t = 5))) +
    scale_x_continuous(breaks = seq(-100, 100, by = 10))  # X-axis breaks at every 10%
}

# Variable names and their simple names for each method's median value
variables <- list(
  list(name = "DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price", simple = "DCF"),
  list(name = "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price", simple = "LTM"),
  list(name = "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price", simple = "NTM")
)

# Pivot the data to long format for easier plotting
long_data <- df %>%
  pivot_longer(cols = c("DCF.Valuation._per.share__Diff..Price.Median.to.Offer.Price",
                        "LTM.Transaction.Multiple._per.share__Diff..Price.Median.to.Offer.Price",
                        "NTM.Trading.Multiple._per.share__Diff..Price.Median.to.Offer.Price"), 
               names_to = "Variable", values_to = "Percentage_Error")

# Generate and print plots for each variable name and its simple name
for (var in variables) {
  plot_histogram(long_data, var$simple)
}



# Load necessary package
library(stats)

# Degrees of freedom, adjust this as necessary
df <- 41

# Confidence levels
confidence_levels <- c(0.90, 0.95, 0.975, 0.99, 0.999)

# Calculate two-tailed critical t-values
t_values <- qt(1 - (1 - confidence_levels) / 2, df)

# Print the results
data.frame(Confidence_Interval = paste0(confidence_levels * 100, "%"), Critical_t_Values = t_values)









library(ggplot2)
library(dplyr)
library(tidyr)

# Assuming 'df' is loaded with necessary columns
vars_valuation_perc <- c(
  
  "DCF.Valuation._per.share__Perc_Diff.Median.to.Offer.Price",
  "LTM.Transaction.Multiple._per.share__Perc_Diff.Median.to.Offer.Price" ,
  "NTM.Trading.Multiple._per.share__Perc_Diff.Median.to.Offer.Price"
)

long_data <- df %>%
  pivot_longer(cols = all_of(vars_valuation_perc), names_to = "Variable", values_to = "Percentage_Error") %>%
  mutate(Error_Type = case_when(
    
    
    Variable %in% c(
      "DCF.Valuation._per.share__Perc_Diff.Median.to.Offer.Price",
      "LTM.Transaction.Multiple._per.share__Perc_Diff.Median.to.Offer.Price",
      "NTM.Trading.Multiple._per.share__Perc_Diff.Median.to.Offer.Price"
    ) ~ "Mid",
    
    TRUE ~ "Other"  # Any variable names that don't match any of the specified patterns
  ))

# Color mappings for the legend
fill_colors <- c(
  
  "Mid" = "slateblue"
)


# Labels for the legend (Error Type) should be the unique names from 'Variable' after filtering
legend_labels <- c(
  "Low", "High", "Mid"
)

# Function to generate histogram for each method
plot_histogram <- function(data, method) {
  filtered_data <- data %>%
    filter(grepl(method, Variable)) %>%
    drop_na(Percentage_Error)  # Drop NA values from Percentage_Error column
  
  if (nrow(filtered_data) == 0) {
    message(paste("No data available for method:", method))
    return(NULL)  # Exit if no data to plot
  }
  
  # Define the range for the x-axis
  x_range <- range(filtered_data$Percentage_Error, na.rm = TRUE)
  x_breaks <- seq(floor(x_range[1]/10)*10, ceiling(x_range[2]/10)*10, by = 10)
  
  ggplot(filtered_data, aes(x = Percentage_Error, fill = Error_Type)) +
    geom_histogram(bins = 30, alpha = 0.7, position = "identity", binwidth = 10) +
    scale_fill_manual(values = fill_colors) +  # Ensures correct color application
    labs(title = paste("Distribution of Percentage Errors -", method),
         x = "Percentage Error (%)",
         y = "Frequency",
         fill = "Error Type") +
    theme_minimal() +
    theme(text = element_text(family = "Arial"),
          legend.title = element_blank(),
          legend.position = "bottom",
          legend.text = element_text(size = 10),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.text.x = element_text(hjust = 1, margin = margin(t = 5))) +
    scale_x_continuous(breaks = x_breaks, labels = x_breaks)  # Dynamically set breaks and labels based on data range
}


# Generate and display plots for each method
methods <- c("DCF", "LTM", "NTM")
plots <- list()
for (method in methods) {
  plots[[method]] <- plot_histogram(long_data, method)
  if (!is.null(plots[[method]])) {
    print(plots[[method]])
  }
}