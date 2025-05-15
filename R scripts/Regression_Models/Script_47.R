# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/overstockings/Order 2")

# Load the openxlsx library
library(openxlsx)
library(dplyr)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("EPA PFAS Data Summary Q4 2023 (2) (1).xlsx")


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


str(df)

names(df)[names(df) == "Maximum.Level.of.Total.PFAS._ppt_"] <- "Maximum.Level.of.Total.PFAS"


# Outliers

calculate_z_scores <- function(data, vars, id_var, z_threshold = 3) {
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~ (.-mean(.))/sd(.), .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  return(z_score_data)
}

df_zscores <- calculate_z_scores(df, "Maximum.Level.of.Total.PFAS", "Public.Water.System")
colnames(df)

remove_outliers_by_zscore <- function(data, zscore_data, id_var, z_vars, z_threshold = 3) {
  # Identify the rows in zscore_data where any of the z-scores exceed the threshold
  outlier_ids <- zscore_data %>%
    filter(across(starts_with("z_"), ~ abs(.) > z_threshold)) %>%
    pull(!!sym(id_var))
  
  # Filter the main data to exclude rows with IDs that have z-scores above the threshold
  filtered_data <- data %>%
    filter(!(!!sym(id_var) %in% outlier_ids))
  
  return(filtered_data)
}

# Usage
df_filtered <- remove_outliers_by_zscore(df, df_zscores, "Public.Water.System", "z_Maximum.Level.of.Total.PFAS", z_threshold = 3)



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


df_descriptive_stats <- calculate_descriptive_stats(df_filtered, "Maximum.Level.of.Total.PFAS")


library(dplyr)

calculate_descriptive_stats_bygroups <- function(data, desc_vars, cat_column) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category = character(),
    Variable = character(),
    N = integer(),
    Mean = numeric(),
    Median = numeric(), # Add Median column
    SEM = numeric(),     # Standard Error of the Mean
    SD = numeric(),      # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each category
  for (var in desc_vars) {
    # Group data by the categorical column
    grouped_data <- data %>%
      group_by(!!sym(cat_column)) %>%
      summarise(
        N = sum(!is.na(!!sym(var))),
        Mean = mean(!!sym(var), na.rm = TRUE),
        Median = median(!!sym(var), na.rm = TRUE), # Calculate Median
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(N),
        SD = sd(!!sym(var), na.rm = TRUE)
      ) %>%
      mutate(Variable = var)
    
    # Append the results for each variable and category to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}

# Example usage
df_descriptive_stats_PWSys <- calculate_descriptive_stats_bygroups(df_filtered, "Maximum.Level.of.Total.PFAS", "Public.Water.System")
df_descriptive_stats_City <- calculate_descriptive_stats_bygroups(df_filtered, "Maximum.Level.of.Total.PFAS", "City")
df_descriptive_stats_County <- calculate_descriptive_stats_bygroups(df_filtered, "Maximum.Level.of.Total.PFAS", "County")


## BOXPLOTS DIVIDED BY FACTOR

library(ggplot2)

# Specify the factor as a column name
factor_var <- "County"  # replace with your actual factor column name

# Loop for generating and saving boxplots
for (var in "Maximum.Level.of.Total.PFAS") {
  # Create boxplot for each variable using ggplot2
  p <- ggplot(df_filtered, aes_string(x = factor_var, y = var)) +
    geom_boxplot() +
    stat_summary(fun.y = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = paste("Boxplot of", var, "by", factor_var),
         x = factor_var, y = var)
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = paste("boxplot_", var, ".png", sep = ""), plot = p, width = 10, height = 6)
}

# Load ggplot2 for plotting
library(ggplot2)

# Create the histogram
histogram_plot <- ggplot(df, aes(x = Maximum.Level.of.Total.PFAS)) +
  geom_histogram(binwidth = 10, color = "black", fill = "blue", alpha = 0.7) +
  labs(title = "Histogram of Maximum Level of Total PFAS",
       x = "Maximum Level of Total PFAS",
       y = "Frequency") +
  theme_minimal()

# Display the plot
print(histogram_plot)

# Save the plot to the working directory
ggsave("Histogram_Maximum_Level_of_Total_PFAS.png", plot = histogram_plot, width = 8, height = 6, dpi = 300)


# Create a new dataframe with the log-transformed variable
df_log <- df %>%
  mutate(log_Maximum.Level.of.Total.PFAS = log(Maximum.Level.of.Total.PFAS))

# Plot histogram of the log-transformed data
log_histogram_plot <- ggplot(df_log, aes(x = log_Maximum.Level.of.Total.PFAS)) +
  geom_histogram(binwidth = 0.5, color = "black", fill = "blue", alpha = 0.7) +
  labs(title = "Histogram of Log-Transformed Maximum Level of Total PFAS",
       x = "Log(Maximum Level of Total PFAS)",
       y = "Frequency") +
  theme_minimal()

# Display the plot
print(log_histogram_plot)

# Save the plot to the working directory
ggsave("Log_Histogram_Maximum_Level_of_Total_PFAS.png", plot = log_histogram_plot, width = 8, height = 6, dpi = 300)



# Load necessary libraries
library(dplyr)
library(broom)

df$County <- relevel(as.factor(df$County), ref = "los angeles")

# Function to run and return tidy results of a GLM with Gamma distribution
run_glm_gamma <- function(data, response_var, predictor_var) {
  
  # Fit the GLM with Gamma distribution and log link
  glm_model <- glm(as.formula(paste(response_var, "~", predictor_var)), 
                   data = data, 
                   family = Gamma(link = "log"))
  
  # Get tidy results (coefficients and p-values)
  tidy_results <- tidy(glm_model)
  
  # Calculate and extract relevant fit statistics
  fit_stats <- list(
    AIC = AIC(glm_model),
    Deviance = deviance(glm_model),
    Residual_Deviance = glm_model$deviance
  )
  
  # Return both tidy results and fit statistics
  return(list(
    tidy_results = tidy_results,
    fit_statistics = fit_stats
  ))
}

# Example usage of the function
results <- run_glm_gamma(df, "Maximum.Level.of.Total.PFAS", "County")

# View tidy results (coefficients, standard errors, p-values, etc.)
model_results <- results$tidy_results

# View fit statistics (AIC, Deviance, Residual Deviance)
results$fit_statistics


# Load the emmeans package for pairwise comparisons
library(emmeans)

# Fit the GLM with Gamma family
glm_model <- glm(Maximum.Level.of.Total.PFAS ~ County, 
                 data = df, 
                 family = Gamma(link = "log"))

# Perform pairwise comparisons between counties
pairwise_comparisons <- emmeans(glm_model, pairwise ~ County, adjust = "bonferroni")

# Extract the pairwise comparison results into a tidy dataframe
pairwise_df <- as.data.frame(pairwise_comparisons$contrasts)

# Extract the means (emmGrid) into a tidy dataframe if needed
emmeans_df <- as.data.frame(pairwise_comparisons$emmeans)


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
  "Data" = df, 
  "Descriptive Stats" = df_descriptive_stats, 
  "Descriptive Stats City" = df_descriptive_stats_City, 
  "Descriptive Stats County" = df_descriptive_stats_County, 
  "Model results" = model_results,
  "Zscores" = df_zscores,
  "EM Means" = emmeans_df,
  "Pairwise Comparisons" = pairwise_df
   
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")



# Load necessary libraries
library(dplyr)
library(broom)
library(emmeans)

# Define the cities of interest
cities_of_interest <- c("san luis obispo", "oxnard", "el monte", "bakersfield", 
                        "orange", "santa clarita", "atascadero", "norwalk", 
                        "riverside", "san bernardino", "whittier", "willowbrook")

# Filter the dataframe for the selected cities
df_filtered_city <- df_filtered %>%
  filter(City %in% cities_of_interest)

# Set "san luis obispo" as the reference level for City
df_filtered$City <- relevel(as.factor(df_filtered$City), ref = "san luis obispo")

# Function to run and return tidy results of a GLM with Gamma distribution
run_glm_gamma <- function(data, response_var, predictor_var) {
  
  # Fit the GLM with Gamma distribution and log link
  glm_model <- glm(as.formula(paste(response_var, "~", predictor_var)), 
                   data = data, 
                   family = Gamma(link = "log"))
  
  # Get tidy results (coefficients and p-values)
  tidy_results <- tidy(glm_model)
  
  # Calculate and extract relevant fit statistics
  fit_stats <- list(
    AIC = AIC(glm_model),
    Deviance = deviance(glm_model),
    Residual_Deviance = glm_model$deviance
  )
  
  # Return both tidy results and fit statistics
  return(list(
    tidy_results = tidy_results,
    fit_statistics = fit_stats
  ))
}

# Example usage of the function
results <- run_glm_gamma(df_filtered, "Maximum.Level.of.Total.PFAS", "City")

# View tidy results (coefficients, standard errors, p-values, etc.)
model_results <- results$tidy_results

# View fit statistics (AIC, Deviance, Residual Deviance)
results$fit_statistics

# Fit the GLM with Gamma family for pairwise comparisons
glm_model <- glm(Maximum.Level.of.Total.PFAS ~ City, 
                 data = df_filtered, 
                 family = Gamma(link = "log"))

# Perform pairwise comparisons between cities
pairwise_comparisons <- emmeans(glm_model, pairwise ~ City, adjust = "bonferroni")

# Extract the pairwise comparison results into a tidy dataframe
pairwise_df <- as.data.frame(pairwise_comparisons$contrasts)

# Extract the means (emmeans) into a tidy dataframe if needed
emmeans_df <- as.data.frame(pairwise_comparisons$emmeans)

# Example usage
data_list <- list(
  
  "Model results" = model_results,
  
  "EM Means" = emmeans_df,
  "Pairwise Comparisons" = pairwise_df
  
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_Cities.xlsx")

# Specify the factor as a column name
factor_var <- "City"  # replace with your actual factor column name

# Loop for generating and saving boxplots
for (var in "Maximum.Level.of.Total.PFAS") {
  # Create boxplot for each variable using ggplot2
  p <- ggplot(df_filtered, aes_string(x = factor_var, y = var)) +
    geom_boxplot() +
    stat_summary(fun.y = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = paste("Boxplot of", var, "by", factor_var),
         x = factor_var, y = var)
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = paste("boxplotCity_", var, ".png", sep = ""), plot = p, width = 10, height = 6)
}



# Power Analysis


# Approximate R^2 for the City model
null_dev_city <- 9.7604
residual_dev_city <- 7.1878
R2_city <- 1 - (residual_dev_city / null_dev_city)
f2_city <- R2_city / (1 - R2_city)


# Approximate R^2 for the County model
null_dev_county <- 39.986
residual_dev_county <- 37.593
R2_county <- 1 - (residual_dev_county / null_dev_county)
f2_county <- R2_county / (1 - R2_county)

library(pwr)

# Set parameters
alpha <- 0.05

# Number of observations and predictors for City model
n_city <- 54  # Adjust based on actual sample size
predictors_city <- 12  # Based on the number of cities in the model

# Calculate power for City model
power_city <- pwr.f2.test(u = predictors_city, v = n_city - predictors_city - 1, f2 = f2_city, sig.level = alpha)
print(power_city)

# Number of observations and predictors for County model
n_county <- 194  # Adjust based on actual sample size
predictors_county <- 9  # Based on the number of counties in the model

# Calculate power for County model
power_county <- pwr.f2.test(u = predictors_county, v = n_county - predictors_county - 1, f2 = f2_county, sig.level = alpha)
print(power_county)


library(car)       # For multicollinearity diagnostics
library(lmtest)    # For Breusch-Pagan test
library(ggplot2)   # For residual plots

# Model Residual Check

glm_model <- glm(Maximum.Level.of.Total.PFAS ~ County, 
                 data = df_filtered, 
                 family = Gamma(link = "log"))

# 2. Normality of Residuals
# Extract residuals from the GLMM model
residuals_glm <- residuals(glm_model, type = "deviance")

# Histogram of residuals
ggplot(data = data.frame(residuals_glm), aes(x = residuals_glm)) +
  geom_histogram(bins = 30, color = "black", fill = "blue", alpha = 0.7) +
  labs(title = "Histogram of Residuals", x = "Residuals", y = "Frequency") +
  theme_minimal()

# Kolmogorov-Smirnov test for normality
ks_test <- ks.test(residuals_glm, "pnorm", mean = mean(residuals_glm), sd = sd(residuals_glm))
print("Kolmogorov-Smirnov Test for Normality:")
print(ks_test)

# 3. Homoscedasticity Checks
# Residual plot
ggplot(data = data.frame(fitted = fitted(glm_model), residuals = residuals_glm), 
       aes(x = fitted, y = residuals)) +
  geom_point(alpha = 0.7) +
  geom_smooth(method = "lm", color = "red", se = FALSE) +
  labs(title = "Residuals vs. Fitted Values", x = "Fitted Values", y = "Residuals") +
  theme_minimal()

# Q-Q Plot for Residuals
qq_plot <- ggplot(data = data.frame(residuals_glm), aes(sample = residuals_glm)) +
  stat_qq() +
  stat_qq_line() +
  labs(title = "Q-Q Plot of Residuals", x = "Theoretical Quantiles", y = "Sample Quantiles") +
  theme_minimal()
print(qq_plot)

# Breusch-Pagan test for heteroscedasticity
bp_test <- bptest(glm_model)
print("Breusch-Pagan Test for Homoscedasticity:")
print(bp_test)

# Save diagnostic plots
ggsave("Residual_Histogram.png", width = 8, height = 6, dpi = 300)
ggsave("Residuals_vs_Fitted.png", width = 8, height = 6, dpi = 300)



# Load necessary libraries
library(car)       # For multicollinearity diagnostics
library(lmtest)    # For Breusch-Pagan test
library(ggplot2)   # For plotting

# Fit GLMM Model with City as Predictor
glm_model_city <- glm(Maximum.Level.of.Total.PFAS ~ City, 
                      data = df_filtered_city, 
                      family = Gamma(link = "log"))

# 1. Normality of Residuals
# Extract residuals from the GLMM model
residuals_glm_city <- residuals(glm_model_city, type = "deviance")

# Histogram of residuals
ggplot(data = data.frame(residuals_glm_city), aes(x = residuals_glm_city)) +
  geom_histogram(bins = 30, color = "black", fill = "blue", alpha = 0.7) +
  labs(title = "Histogram of Residuals", x = "Residuals", y = "Frequency") +
  theme_minimal()

# Kolmogorov-Smirnov test for normality
ks_test_city <- ks.test(residuals_glm_city, "pnorm", 
                        mean = mean(residuals_glm_city), sd = sd(residuals_glm_city))
print("Kolmogorov-Smirnov Test for Normality (City):")
print(ks_test_city)

# 2. Homoscedasticity Checks
# Residuals vs Fitted plot
ggplot(data = data.frame(fitted = fitted(glm_model_city), residuals = residuals_glm_city), 
       aes(x = fitted, y = residuals)) +
  geom_point(alpha = 0.7) +
  geom_smooth(method = "lm", color = "red", se = FALSE) +
  labs(title = "Residuals vs. Fitted Values (City)", x = "Fitted Values", y = "Residuals") +
  theme_minimal()

# Q-Q Plot for Residuals
qq_plot_city <- ggplot(data = data.frame(residuals_glm_city), aes(sample = residuals_glm_city)) +
  stat_qq() +
  stat_qq_line() +
  labs(title = "Q-Q Plot of Residuals (City)", x = "Theoretical Quantiles", y = "Sample Quantiles") +
  theme_minimal()
print(qq_plot_city)

# 3. Breusch-Pagan test for heteroscedasticity
bp_test_city <- bptest(glm_model_city)
print("Breusch-Pagan Test for Homoscedasticity (City):")
print(bp_test_city)

# Save Diagnostic Plots for City Model
ggsave("Residual_Histogram_City.png", width = 8, height = 6, dpi = 300)
ggsave("Residuals_vs_Fitted_City.png", width = 8, height = 6, dpi = 300)
ggsave("QQ_Plot_Residuals_City.png", plot = qq_plot_city, width = 8, height = 6, dpi = 300)




library(boot)

# Define function to fit Gamma GLM and extract AIC
glm_boot <- function(data, indices) {
  d <- data[indices, ]  # Bootstrap sample
  model <- glm(Maximum.Level.of.Total.PFAS ~ County, data = d, family = Gamma(link = "log"))
  return(AIC(model))  # Return AIC for validation
}

# Bootstrap with 1000 resamples
set.seed(123)
boot_results <- boot(data = df_filtered, statistic = glm_boot, R = 1000)

# Summary of bootstrap results
print("Bootstrap Validation Results:")
print(boot_results)
plot(boot_results)  # Visualize AIC distribution


library(boot)

# Define function to fit Gamma GLM and extract AIC
glm_boot <- function(data, indices) {
  d <- data[indices, ]  # Bootstrap sample
  model <- glm(Maximum.Level.of.Total.PFAS ~ City, data = d, family = Gamma(link = "log"))
  return(AIC(model))  # Return AIC for validation
}

# Bootstrap with 1000 resamples
set.seed(123)
boot_results <- boot(data = df_filtered_city, statistic = glm_boot, R = 1000)

# Summary of bootstrap results
print("Bootstrap Validation Results:")
print(boot_results)
plot(boot_results)  # Visualize AIC distribution






bootstrap_anova_classic <- function(data, response, group, R = 1000) {
  formula <- as.formula(paste(response, "~", group))
  
  # Fit original model
  model <- lm(formula, data = data)
  anova_orig <- anova(model)
  f_value <- anova_orig$`F value`[1]
  df1 <- anova_orig$Df[1]
  df2 <- anova_orig$Df[2]
  p_value <- anova_orig$`Pr(>F)`[1]
  
  # Bootstrapped F statistic
  boot_stat <- function(data, indices) {
    d <- data[indices, ]
    mod <- tryCatch(lm(formula, data = d), error = function(e) NA)
    if (any(is.na(mod))) return(NA)
    f <- tryCatch(anova(mod)$`F value`[1], error = function(e) NA)
    return(f)
  }
  
  set.seed(123)
  boot_result <- boot(data, boot_stat, R = R)
  p_boot <- mean(boot_result$t >= f_value, na.rm = TRUE)
  
  df_out <- data.frame(
    Response = response,
    Grouping = group,
    F_value = round(f_value, 3),
    df1 = df1,
    df2 = df2,
    p = round(p_value, 4),
    p_boot = round(p_boot, 4),
    R_boot = sum(!is.na(boot_result$t))
  )
  
  return(df_out)
}


bootstrap_tukey_stratified <- function(data, response, group, R = 1000) {
  formula <- as.formula(paste(response, "~", group))
  
  # Base model
  model <- aov(formula, data = data)
  tukey <- TukeyHSD(model)
  base_df <- as.data.frame(tukey[[1]])
  base_df$Comparison <- rownames(base_df)
  rownames(base_df) <- NULL
  comparisons_list <- base_df$Comparison
  
  boot_mat <- matrix(NA, nrow = R, ncol = length(comparisons_list))
  colnames(boot_mat) <- comparisons_list
  
  # Get group labels
  groups <- unique(data[[group]])
  
  for (i in 1:R) {
    # Stratified resampling
    boot_data <- do.call(rbind, lapply(groups, function(g) {
      group_data <- data[data[[group]] == g, ]
      group_data[sample(nrow(group_data), replace = TRUE), ]
    }))
    
    mod <- tryCatch(aov(formula, data = boot_data), error = function(e) NULL)
    if (is.null(mod)) next
    
    tuk <- tryCatch(TukeyHSD(mod)[[1]], error = function(e) NULL)
    if (is.null(tuk)) next
    
    res <- as.data.frame(tuk)
    res$Comparison <- rownames(res)
    boot_diff_named <- setNames(res$diff, res$Comparison)
    common <- intersect(comparisons_list, names(boot_diff_named))
    boot_mat[i, common] <- boot_diff_named[common]
  }
  
  # Calculate bootstrapped confidence intervals
  ci_lower <- apply(boot_mat, 2, function(x) quantile(x, 0.025, na.rm = TRUE))
  ci_upper <- apply(boot_mat, 2, function(x) quantile(x, 0.975, na.rm = TRUE))
  r_used <- colSums(!is.na(boot_mat))
  
  base_df$Estimate <- round(base_df$diff, 3)
  base_df$CI_Lower <- round(ci_lower[base_df$Comparison], 3)
  base_df$CI_Upper <- round(ci_upper[base_df$Comparison], 3)
  base_df$R_boot <- r_used[base_df$Comparison]
  
  out_df <- base_df[, c("Comparison", "Estimate", "CI_Lower", "CI_Upper", "R_boot")]
  return(out_df)
}


cities_of_interest <- c("san luis obispo", "oxnard", "el monte", "bakersfield", 
                        "orange", "santa clarita")


# Filter the dataframe for the selected cities
df_filtered_city <- df_filtered %>%
  filter(City %in% cities_of_interest)

# Use with df_filtered_city or your filtered data
anova_classic_city <- bootstrap_anova_classic(df_filtered_city, "Maximum.Level.of.Total.PFAS", "City")
print(anova_classic_city)

tukey_classic_city <- bootstrap_tukey_stratified(df_filtered_city, "Maximum.Level.of.Total.PFAS", "City")
print(tukey_classic_city)

anova_classic_county <- bootstrap_anova_classic(df_filtered, "Maximum.Level.of.Total.PFAS", "County")
print(anova_classic_county)

tukey_classic_county <- bootstrap_tukey_stratified(df_filtered, "Maximum.Level.of.Total.PFAS", "County")
print(tukey_classic_county)


# Boxplot
boxplot_cities <- ggplot(df_filtered_city, aes(x = reorder(City, Maximum.Level.of.Total.PFAS, median), 
                                               y = Maximum.Level.of.Total.PFAS)) +
  geom_boxplot(outlier.shape = NA) +
  geom_jitter(width = 0.2, alpha = 0.5, color = "gray40") +
  stat_summary(fun = mean, geom = "point", shape = 18, size = 3.5, color = "red") +
  labs(title = "PFAS Levels by City (Filtered Sample ≥ 3)", 
       x = "City", 
       y = "Maximum PFAS Level (ppt)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Save plot
ggsave("Boxplot_PFAS_by_City_Filtered.png", plot = boxplot_cities, width = 12, height = 6, dpi = 300)

library(car)

levene_test_tidy <- function(data, response, group) {
  formula <- as.formula(paste(response, "~", group))
  lev_result <- leveneTest(formula, data = data, center = median)
  
  df_out <- data.frame(
    Response = response,
    Grouping = group,
    Test = "Levene's Test",
    Center = "Median",
    F_value = round(lev_result[1, "F value"], 3),
    df1 = lev_result[1, "Df"],
    df2 = lev_result[2, "Df"],
    p_value = round(lev_result[1, "Pr(>F)"], 4)
  )
  
  return(df_out)
}


levene_city <- levene_test_tidy(df_filtered_city, "Maximum.Level.of.Total.PFAS", "City")
print(levene_city)

levene_county <- levene_test_tidy(df_filtered, "Maximum.Level.of.Total.PFAS", "County")
print(levene_county)


games_howell_tidy <- function(data, response, group) {
  library(ggstatsplot)
  library(dplyr)
  
  posthoc_table <- pairwise_comparisons(
    data = data,
    x = !!rlang::sym(group),
    y = !!rlang::sym(response),
    type = "parametric",
    paired = FALSE,
    p.adjust.method = "none",
    var.equal = FALSE,
    effsize.type = "g"
  )
  
  tidy_df <- posthoc_table %>%
    mutate(
      Comparison = paste(group1, "-", group2),
      t = round(statistic, 3),
      p_value = round(p.value, 4)
    ) %>%
    select(Comparison, t, p_value, test, p.adjust.method)
  
  return(tidy_df)
}


gh_posthoc_county <-  games_howell_tidy(df_filtered, "Maximum.Level.of.Total.PFAS", "County")
print(gh_posthoc_county)


# Example usage
data_list <- list(
  "Data" = df, 
  "Descriptive Stats" = df_descriptive_stats, 
  "Descriptive Stats City" = df_descriptive_stats_City, 
  "Descriptive Stats County" = df_descriptive_stats_County,
  "Zscores" = df_zscores,
  "EM Means" = emmeans_df,
  "Pairwise GLM Comparisons" = pairwise_df,
  "posthoc City" =  tukey_classic_city,
  "posthoc County" = gh_posthoc_county
  )

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables3.xlsx")
