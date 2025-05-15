# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/bianalkahsem")

# Load the openxlsx library
library(openxlsx)
library(dplyr)

# Get the names of all sheets in the Excel file
sheets <- getSheetNames("groups data nov 2024.xlsx")

# Initialize an empty list to store data frames
dfs <- list()

# Loop over each sheet name and read the sheet into a separate data frame
for (i in seq_along(sheets)) {
  df <- read.xlsx("groups data nov 2024.xlsx", sheet = sheets[i])
  
  # Clean up column names
  names(df) <- gsub("\\s+|\\(|\\)|-|/|\\\\|\\?|'|,|\\$|\\+", "_", names(df))
  names(df) <- gsub("_+", "_", names(df)) # Replace multiple underscores with a single one
  names(df) <- gsub("_$", "", names(df)) # Remove trailing underscores
  
  # Trim and convert to lowercase for all character columns
  df <- data.frame(lapply(df, function(x) {
    if (is.character(x)) {
      x <- tolower(trimws(x))
    }
    return(x)
  }))
  
  # Store the dataframe in the list with a name
  dfs[[paste0("df", i)]] <- df
}





# Concatenate all data frames into one
df_all <- do.call(rbind, dfs)

str(df_all)
colnames(df_all)


library(plm)
library(lmtest)

data <- pdata.frame(df_all, index = c("Country", "Year"))
pooled_ols1 <- plm(GDP_BILLIONS ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "pooling")

# Breusch-Pagan test to check for random effects
plmtest(pooled_ols1, type = "bp")

pooled_ols2 <- plm(HDI ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "pooling")

# Breusch-Pagan test to check for random effects
plmtest(pooled_ols2, type = "bp")

pooled_ols3 <- plm(TRADE...GDP ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "pooling")

# Breusch-Pagan test to check for random effects
plmtest(pooled_ols3, type = "bp")

fe_model1 <- plm(GDP_BILLIONS ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "within")
re_model1 <- plm(GDP_BILLIONS ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "random")
hausman_test1 <- phtest(fe_model1, re_model1)
print(hausman_test1)
pdw_result1 <- pdwtest(fe_model1, alternative = "two.sided")
print(pdw_result1)

# Using fixed effects model residuals and fitted values for testing
fe_lm1 <- lm(residuals(fe_model1) ~ fitted(fe_model1))
bp1 <- bptest(fe_lm1)
print(bp1)

fe_model2 <- plm(HDI ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "within")
re_model2 <- plm(HDI ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "random")
hausman_test2 <- phtest(fe_model2, re_model2)
print(hausman_test2)
pdw_result2 <- pdwtest(fe_model2, alternative = "two.sided")
print(pdw_result2)

# Using fixed effects model residuals and fitted values for testing
fe_lm2 <- lm(residuals(fe_model2) ~ fitted(fe_model2))
bp2 <- bptest(fe_lm2)
print(bp2)

fe_model3 <- plm(TRADE...GDP ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "within")
re_model3 <- plm(TRADE...GDP ~ rule.of.law + property.rights + freedom.of.expression.index + civil.liberties.index + gender.equality.in.respect.for.civil.liberties, data = data, model = "random")
hausman_test3 <- phtest(fe_model3, re_model3)
print(hausman_test3)
pdw_result3 <- pdwtest(fe_model3, alternative = "two.sided")
print(pdw_result3)

# Using fixed effects model residuals and fitted values for testing
fe_lm3 <- lm(residuals(fe_model3) ~ fitted(fe_model3))
bp3 <- bptest(fe_lm3)
print(bp3)

# APPLY DIFFERENCING DUE TO AUTOCORRELATION

# Define dependent variables
dvs <- c("GDP_BILLIONS", "HDI", "TRADE...GDP")
independent_vars <- c("rule.of.law", "property.rights", "freedom.of.expression.index", 
                      "civil.liberties.index", "gender.equality.in.respect.for.civil.liberties")

# Load necessary libraries
library(plm)
library(lmtest)

# Differencing the variables
data_diff <- data %>%
  group_by(Country) %>%  # Group by Country
  mutate(across(c(dvs, independent_vars), 
                ~ (.-lag(.)) / lag(.) * 100,  # Calculate percentage change
                .names = "d_{.col}")) %>%
  ungroup()  # Ungroup to avoid unintended side effects

# Convert to panel data format
data_diff <- pdata.frame(data_diff, index = c("Country", "Year"))


# Function to apply differencing to a dataset
apply_percentage_differencing <- function(data, vars_to_diff, id_vars = c("Country", "Year")) {
  data <- data %>%
    mutate(across(all_of(vars_to_diff), ~ as.numeric(.))) %>%  # ← CONVERSÃO AQUI
    arrange(across(all_of(id_vars))) %>%
    group_by(across(all_of(id_vars[1]))) %>%
    mutate(across(
      all_of(vars_to_diff), 
      ~ ifelse(lag(.) == 0 | is.na(lag(.)), NA, (. - lag(.)) / lag(.) * 100), 
      .names = "d_{.col}"
    )) %>%
    ungroup()
  
  pdata.frame(data, index = id_vars)
}



vars_to_diff <- c("GDP_BILLIONS", "HDI", "TRADE...GDP", "rule.of.law", "property.rights", "freedom.of.expression.index", 
                  "civil.liberties.index", "gender.equality.in.respect.for.civil.liberties")

# Apply differencing to each dataset and store in a list
differenced_dfs <- lapply(dfs, function(df) apply_percentage_differencing(df, vars_to_diff))

dfs <- differenced_dfs

df1 <- dfs$df1
df2 <- dfs$df2
df3 <- dfs$df3
df4 <- dfs$df4


# Define dependent variables
dvs <- c("d_GDP_BILLIONS", "d_HDI", "d_TRADE...GDP")
independent_vars <- c("d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
                      "d_civil.liberties.index", "d_gender.equality.in.respect.for.civil.liberties")

# Outlier Evaluation

library(dplyr)

calculate_z_scores <- function(data, vars, id_vars = c("Country", "Year"), z_threshold = 4) {
  # Ensure the columns used for Z-score calculations are numeric and handle NA
  data <- data %>%
    mutate(across(all_of(vars), ~ as.numeric(.), .names = "{.col}_numeric")) %>%
    mutate(across(ends_with("_numeric"), ~ replace(., is.na(.), mean(., na.rm = TRUE)), .names = "{.col}_filled"))
  
  # Calculate Z-scores relative to the entire dataset
  z_score_data <- data %>%
    mutate(across(ends_with("_filled"), 
                  ~ (.-mean(.))/ifelse(sd(.) == 0, 1, sd(.)), 
                  .names = "z_{.col}"))
  
  # Flag outliers
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), 
                  ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), 
                  .names = "flag_{.col}"))
  
  # Select relevant columns for output
  output_results <- z_score_data %>%
    dplyr::select(
      dplyr::all_of(id_vars),
      dplyr::all_of(vars),
      dplyr::starts_with("z_"),
      dplyr::starts_with("flag_")
    )
  
  
  
  return(output_results)
}


# Define the variables for outlier detection
vars_to_check <- c("d_GDP_BILLIONS", "d_HDI", "d_TRADE...GDP", "d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
                      "d_civil.liberties.index", "d_gender.equality.in.respect.for.civil.liberties")

# Apply the function to detect outliers
outlier_results <- calculate_z_scores(data = data_diff, vars = vars_to_check)

# Count the number of outliers for each variable
outlier_counts <- outlier_results %>%
  dplyr::select(dplyr::starts_with("flag_")) %>%
  dplyr::summarise(across(everything(), ~ sum(. == "Outlier", na.rm = TRUE)))


# Rename columns for clarity
colnames(outlier_counts) <- sub("flag_", "", colnames(outlier_counts))

# Display the outlier counts
print(outlier_counts)




# Function to print flagged outliers for each variable
print_outliers_full <- function(outlier_results, vars) {
  for (var in vars) {
    # Identify the correct flag column for the variable
    flag_column <- paste0("flag_z_", var, "_numeric_filled")
    
    # Check if the flag column exists in the outlier_results
    if (flag_column %in% colnames(outlier_results)) {
      cat("\nOutliers for variable:", var, "\n")
      
      # Filter and print the country and year for flagged outliers
      outliers <- outlier_results %>%
        filter(!!sym(flag_column) == "Outlier") %>%
        dplyr::select(Country, Year)
      
      
      if (nrow(outliers) > 0) {
        # Print outliers, replacing NA with a readable format
        print(outliers, n = Inf, na.print = "NA")
      } else {
        cat("No outliers detected.\n")
      }
    } else {
      cat("\nNo flag column found for variable:", var, "\n")
    }
  }
}

# Call the function to print all outliers
print_outliers_full(outlier_results, vars_to_check)

str(outlier_results)

# Define outliers as a list of data frames
outlier_list <- list(
  d_GDP_BILLIONS = data.frame(
    Country = c("brazil", "brazil", "brazil", "bulgaria", "moldova", 
                "nicaragua", "russia"),
    Year = c("1992", "1993", "1994", "1997", "1993", "1991", "1993")
  ),
  d_HDI = data.frame(
    Country = c("cyprus", "ethiopia", "malawi", "republic of the congo", 
                "sierra leone"),
    Year = c("1993", "2005", "1995", "1997", "2000")
  ),
  d_TRADE...GDP = data.frame(
    Country = c("indonesia", "republic of the congo", "russia", "sierra leone"),
    Year = c("1998", "1994", "1992", "2000")
  ),
  d_rule.of.law = data.frame(
    Country = c("gambia, the", "indonesia", "indonesia", "maldives", 
                "republic of the congo", "republic of the congo", 
                "sierra leone"),
    Year = c("2017", "1998", "1999", "2019", "1991", "1997", "1996")
  ),
  d_property.rights = data.frame(
    Country = c("cuba", "djibouti", "ethiopia", "ethiopia", "guatemala", 
                "russia", "sierra leone", "tanzania"),
    Year = c("2001", "2002", "1991", "1992", "2000", "1992", "2002", "1991")
  ),
  d_freedom.of.expression.index = data.frame(
    Country = c("ethiopia", "gambia, the", "indonesia", "malawi", "malawi", 
                "republic of the congo", "togo"),
    Year = c("1991", "2017", "1998", "1993", "1994", "1991", "1991")
  ),
  d_civil.liberties.index = data.frame(
    Country = c("ethiopia", "gambia, the", "indonesia", "indonesia", 
                "malawi", "malawi", "malawi", "nicaragua", 
                "republic of the congo", "russia", "sierra leone"),
    Year = c("1991", "2017", "1998", "1999", "1993", "1994", "1995", 
             "2018", "1991", "1992", "2002")
  ),
  d_gender.equality.in.respect.for.civil.liberties = data.frame(
    Country = c("cameroon", "ethiopia", "ethiopia", "ethiopia", "fiji", 
                "indonesia", "indonesia", "mexico", "sierra leone", 
                "sierra leone"),
    Year = c("1992", "2014", "2018", "2020", "1997", "1998", "1999", 
             "2000", "1996", "2003")
  )
)



# Function to replace flagged outliers with NA for specific variables
replace_outliers_with_na <- function(df, outlier_list) {
  for (var in names(outlier_list)) {
    # Ensure column exists in df
    if (var %in% names(df)) {
      flagged <- outlier_list[[var]]
      for (i in seq_len(nrow(flagged))) {
        country <- flagged$Country[i]
        year <- flagged$Year[i]
        # Replace the specific variable value with NA
        df[df$Country == country & df$Year == year, var] <- NA
      }
    } else {
      warning(paste("Variable", var, "not found in dataset"))
    }
  }
  return(df)
}

# Assign meaningful names to the list elements
names(differenced_dfs) <- c("df1", "df2", "df3", "df4")

dfs <- differenced_dfs

# Apply the function to each dataset in dfs
dfs <- lapply(dfs, function(df) {
  replace_outliers_with_na(df, outlier_list)
})

# Apply the function to data_diff
data_diff <- replace_outliers_with_na(data_diff, outlier_list)

# Print the remaining rows for verification
cat("Rows remaining in data_diff after outlier removal:", nrow(data_diff), "\n")
for (i in seq_along(dfs)) {
  cat("Rows remaining in dfs dataset", i, "after outlier removal:", nrow(dfs[[i]]), "\n")
}


# Define dependent variables
dvs <- c("d_GDP_BILLIONS", "d_HDI", "d_TRADE...GDP")
independent_vars <- c("d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
                      "d_civil.liberties.index", "d_gender.equality.in.respect.for.civil.liberties")

# Overall Results

# Load necessary libraries
library(plm)
library(lmtest)
library(sandwich)
library(dplyr)
library(car)

# Ensure data_diff is in the correct panel data format
data_diff <- pdata.frame(data_diff, index = c("Country", "Year"))


# List to store model results and diagnostics
results_list_all <- list()

# Loop through each dependent variable
for (dv in dvs) {
  # Define the formula
  formula_string <- paste0(dv, " ~ ", paste(independent_vars, collapse = " + "))
  
  # Fit the fixed effects model
  fe_model <- plm(as.formula(formula_string), data = data_diff, model = "within", na.action = na.exclude)
  
  robust_se <- vcovBK(fe_model, cluster = "group")
  
  # Summarize the model with robust standard errors
  fe_summary <- summary(fe_model, vcov = robust_se)
  
  # Extract model coefficients
  model_coefficients <- as.data.frame(fe_summary$coefficients)
  colnames(model_coefficients) <- c("Estimate", "Std. Error", "t-value", "Pr(>|t|)")
  
  # Perform diagnostics
  dw_test <- pdwtest(fe_model, alternative = "two.sided")
  bp_test <- bptest(fe_model)
  vif_test <- vif(lm(as.formula(formula_string), data = data_diff))
  # Adjusted R-squared
  adj_r_squared <- fe_summary$r.squared[2]
  
  # Save results for this dependent variable
  results_list_all[[dv]] <- list(
    coefficients = model_coefficients,
    diagnostics = list(
      Durbin_Watson = dw_test,
      Breusch_Pagan = bp_test,
      Adjusted_R2 = adj_r_squared
    )
  )
  
  # Print results
  cat("\nFixed Effects Model Results for", dv, ":\n")
  print(model_coefficients)
  cat("\nDurbin-Watson Test for Serial Correlation in", dv, ":\n")
  print(dw_test)
  cat("\nAdjusted R-squared for", dv, ":\n")
  print(adj_r_squared)
  cat("\nBreusch-Pagan Test on Residuals for", dv, ":\n")
  print(bp_test)
  cat("\nCollinearity Test", dv, ":\n")
  print(vif_test)
  
}

# Save results_list for further analysis
cat("\nAnalysis completed. Results saved in 'results_list'.")


# Get coefficients

# Initialize an empty dataframe to store all coefficients
coefficients_df <- data.frame()

# Loop through each dependent variable
for (dv in dvs) {
  # Define the formula
  formula_string <- paste0(dv, " ~ ", paste(independent_vars, collapse = " + "))
  
  # Fit the fixed effects model
  fe_model <- plm(as.formula(formula_string), data = data_diff, model = "within", na.action = na.exclude)
  
  # Use robust standard errors (vcovBK)
  robust_se <- vcovBK(fe_model, cluster = "group")
  
  # Summarize the model with robust standard errors
  fe_summary <- summary(fe_model, vcov = robust_se)
  
  # Extract model coefficients
  model_coefficients <- as.data.frame(fe_summary$coefficients)
  colnames(model_coefficients) <- c("Estimate", "Std. Error", "t-value", "Pr(>|t|)")
  
  # Add predictor names as a column
  model_coefficients$Predictor <- rownames(fe_summary$coefficients)
  
  # Add dependent variable name as a column
  model_coefficients$Dependent_Variable <- dv
  
  # Add the coefficients to the overall dataframe
  coefficients_df <- rbind(coefficients_df, model_coefficients)
}

# Reset row names for the combined dataframe
rownames(coefficients_df) <- NULL

# Save or display the combined coefficients dataframe
print(coefficients_df)



# List to store model results and diagnostics for each dataset
results_list <- list()

# Loop through each differenced dataset
for (i in seq_along(differenced_dfs)) {
  # Get the current differenced dataset
  data_diff_loop <- differenced_dfs[[i]]
  
  # Initialize results for this dataset
  dataset_results <- list()
  
  # Loop through each dependent variable
  for (dv in dvs) {
    # Define the formula with differenced variables
    formula_string <- paste0(dv, " ~ ", paste(paste0(independent_vars), collapse = " + "))
    
    # Fit the fixed effects model
    fe_model <- plm(as.formula(formula_string), data = data_diff_loop, model = "within", na.action = na.exclude)
    
    # Extract robust standard errors
    robust_se <- vcovHC(fe_model, type = "HC1")
    
    # Summarize the model
    fe_summary <- summary(fe_model, vcov = robust_se)
    model_coefficients <- as.data.frame(fe_summary$coefficients)
    colnames(model_coefficients) <- c("Estimate", "Std. Error", "t-value", "Pr(>|t|)")
    
    # Perform diagnostics
    dw_test <- pdwtest(fe_model, alternative = "two.sided")
    vif_test <- vif(lm(as.formula(formula_string), data = data_diff_loop))
    fe_lm <- lm(residuals(fe_model) ~ fitted(fe_model))
    bp_test <- bptest(fe_lm)
    
    # Adjusted R-squared
    adj_r_squared <- fe_summary$r.squared[2]
    
    # Save results for this dependent variable
    dataset_results[[dv]] <- list(
      coefficients = model_coefficients,
      diagnostics = list(
        Durbin_Watson = dw_test,
        VIF = vif_test,
        Breusch_Pagan = bp_test,
        Adjusted_R2 = adj_r_squared
      )
    )
    
    # Print results to the console
    cat("\nFixed Effects Model Results for", dv, " (Dataset", i, "):\n")
    print(model_coefficients)
    cat("\nDurbin-Watson Test for Serial Correlation in", dv, " (Dataset", i, "):\n")
    print(dw_test)
    cat("\nVariance Inflation Factor (VIF) for", dv, " (Dataset", i, "):\n")
    print(vif_test)
    cat("\nAdjusted R-squared for", dv, " (Dataset", i, "):\n")
    print(adj_r_squared)
    cat("\nBreusch-Pagan Test on Residuals for", dv, " (Dataset", i, "):\n")
    print(bp_test)
  }
  
  # Save results for this dataset
  results_list[[paste0("Dataset_", i)]] <- dataset_results
}


# Take out civil liberties to resolve multicolinearity

# Load necessary libraries
library(plm)
library(lmtest)
library(sandwich)
library(dplyr)
library(car)

# Define dependent variables
dvs <- c("d_GDP_BILLIONS", "d_HDI", "d_TRADE...GDP")
independent_vars <- c("d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
                     "d_gender.equality.in.respect.for.civil.liberties")

# List to store model results and diagnostics for each dataset
results_list <- list()

# Loop through each differenced dataset
for (i in seq_along(differenced_dfs)) {
  # Get the current differenced dataset
  data_diff_loop <- differenced_dfs[[i]]
  
  # Initialize results for this dataset
  dataset_results <- list()
  
  # Loop through each dependent variable
  for (dv in dvs) {
    # Define the formula with differenced variables
    formula_string <- paste0(dv, " ~ ", paste(paste0(independent_vars), collapse = " + "))
    
    # Fit the fixed effects model
    fe_model <- plm(as.formula(formula_string), data = data_diff_loop, model = "within", na.action = na.exclude)
    
    # Extract robust standard errors
    robust_se <- vcovBK(fe_model, cluster = "group")
    
    # Summarize the model
    fe_summary <- summary(fe_model, vcov = robust_se)
    model_coefficients <- as.data.frame(fe_summary$coefficients)
    colnames(model_coefficients) <- c("Estimate", "Std. Error", "t-value", "Pr(>|t|)")
    
    # Perform diagnostics
    dw_test <- pdwtest(fe_model, alternative = "two.sided")
    vif_test <- vif(lm(as.formula(formula_string), data = data_diff_loop))
    fe_lm <- lm(residuals(fe_model) ~ fitted(fe_model))
    bp_test <- bptest(fe_lm)
    
    # Adjusted R-squared
    adj_r_squared <- fe_summary$r.squared[2]
    
    # Save results for this dependent variable
    dataset_results[[dv]] <- list(
      coefficients = model_coefficients,
      diagnostics = list(
        Durbin_Watson = dw_test,
        VIF = vif_test,
        Breusch_Pagan = bp_test,
        Adjusted_R2 = adj_r_squared
      )
    )
    
    # Print results to the console
    cat("\nFixed Effects Model Results for", dv, " (Dataset", i, "):\n")
    print(model_coefficients)
    cat("\nDurbin-Watson Test for Serial Correlation in", dv, " (Dataset", i, "):\n")
    print(dw_test)
    cat("\nVariance Inflation Factor (VIF) for", dv, " (Dataset", i, "):\n")
    print(vif_test)
    cat("\nAdjusted R-squared for", dv, " (Dataset", i, "):\n")
    print(adj_r_squared)
    cat("\nBreusch-Pagan Test on Residuals for", dv, " (Dataset", i, "):\n")
    print(bp_test)
  }
  
  # Save results for this dataset
  results_list[[paste0("Dataset_", i)]] <- dataset_results
}

# Extract coefficients from results_list and create single tables for each dataset
coefficients_tables <- list()

# Loop through each dataset in results_list
for (dataset_name in names(results_list)) {
  # Initialize an empty data frame to store coefficients for this dataset
  dataset_coefficients <- data.frame()
  
  # Get the models for this dataset
  dataset_models <- results_list[[dataset_name]]
  
  # Loop through each dependent variable in the dataset
  for (dv in names(dataset_models)) {
    # Get the coefficients for this DV
    coefficients <- dataset_models[[dv]]$coefficients
    
    # Add a column indicating the dependent variable
    coefficients$DV <- dv
    
    # Combine with the dataset-level table
    dataset_coefficients <- rbind(dataset_coefficients, coefficients)
  }
  
  # Add the final table for this dataset to the list
  coefficients_tables[[dataset_name]] <- dataset_coefficients
}

# Print or save the tables
for (dataset_name in names(coefficients_tables)) {
  cat("\nCoefficients for", dataset_name, ":\n")
  print(coefficients_tables[[dataset_name]])
}


modelresults_df1 <- coefficients_tables$Dataset_1
modelresults_df2 <- coefficients_tables$Dataset_2
modelresults_df3 <- coefficients_tables$Dataset_3
modelresults_df4 <- coefficients_tables$Dataset_4


# Load necessary libraries
library(ggplot2)
library(dplyr)

# Function to convert pdata.frame to plain data frame
strip_pdata_frame <- function(data) {
  as.data.frame(lapply(data, function(col) {
    if (inherits(col, "pseries")) as.numeric(col) else col
  }))
}

# Initialize a combined data frame
combined_data <- data.frame()

# Loop through the datasets and process them
for (i in seq_along(dfs)) {
  # Convert pdata.frame to plain data frame
  data_plain <- strip_pdata_frame(dfs[[i]])
  
  # Add a Group column for identification
  data_plain$Group <- paste0("Group ", i)
  
  # Ensure consistent column names
  if (i == 1) {
    reference_columns <- colnames(data_plain)
  } else {
    # Add missing columns with NA
    missing_cols <- setdiff(reference_columns, colnames(data_plain))
    for (col in missing_cols) {
      data_plain[[col]] <- NA
    }
    # Reorder columns to match the reference
    data_plain <- data_plain[, reference_columns]
  }
  
  # Combine datasets
  combined_data <- bind_rows(combined_data, data_plain)
}

# Directory to save scatterplots
output_dir <- "scatterplots_combined"
if (!dir.exists(output_dir)) dir.create(output_dir)

# Create scatterplots for each dependent variable
for (dv in dvs) {  # Use the column names directly as they are
  for (iv in independent_vars) {  # Use the independent variables as they are
    # Generate the scatterplot
    p <- ggplot(combined_data, aes_string(x = iv, y = dv, color = "Group")) +
      geom_point(alpha = 0.6) +
      geom_smooth(method = "lm", se = FALSE, linetype = "dashed", formula = y ~ x) +
      labs(
        title = paste("Scatterplot of", dv, "vs.", iv, "by Group"),
        x = iv,
        y = dv,
        color = "Group"
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(hjust = 0.5, size = 14),
        axis.title = element_text(size = 12),
        legend.position = "bottom"
      )
    
    # Save the plot
    plot_filename <- paste0(output_dir, "/Scatterplot_", dv, "_vs_", iv, "_Combined.png")
    ggsave(plot_filename, plot = p, width = 8, height = 6, dpi = 300)
  }
}

cat("Scatterplots saved in 'scatterplots_combined' directory.")



# Load necessary libraries
library(ggplot2)
library(dplyr)

# Calculate mean scores for each country and year
calculate_means <- function(data, vars_to_average, id_vars = c("Country", "Year")) {
  data %>%
    group_by(across(all_of(id_vars))) %>%
    summarize(across(all_of(vars_to_average), mean, na.rm = TRUE), .groups = "drop")
}

# List of variables to calculate mean scores
vars_to_average <- c("d_GDP_BILLIONS", "d_HDI", "d_TRADE...GDP")

# Calculate mean scores for each differenced dataframe
mean_scores <- bind_rows(lapply(seq_along(differenced_dfs), function(i) {
  df <- as.data.frame(differenced_dfs[[i]]) # Convert pdata.frame to data frame
  calculate_means(df, vars_to_average) %>%
    mutate(Group = paste0("DF", i)) # Add a group column
}))

# Ensure Year is numeric for plotting
mean_scores$Year <- as.numeric(as.character(mean_scores$Year))

# Directory to save the line charts
output_dir <- "line_charts_mean_scores"
if (!dir.exists(output_dir)) dir.create(output_dir)

# Create line charts for each variable
for (var in vars_to_average) {
  # Generate the line chart
  p <- ggplot(mean_scores, aes(x = Year, y = !!sym(var), color = Group, group = Group)) +
    geom_line(size = 1) +
    geom_point(size = 2) +
    labs(
      title = paste("Mean Scores of", var, "by Year"),
      x = "Year",
      y = paste("Mean", var),
      color = "Dataframe"
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14),
      axis.title = element_text(size = 12),
      legend.position = "bottom"
    )
  
  # Save the plot
  plot_filename <- paste0(output_dir, "/MeanLineChart_", var, ".png")
  ggsave(plot_filename, plot = p, width = 8, height = 6, dpi = 300)
}

cat("Mean score line charts saved in 'line_charts_mean_scores' directory.")


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
  "Clean Diff Data" = data_diff, 
  "Overall Model" = coefficients_df,
  "Group 1 Model" = modelresults_df1,
  "Group 2 Model" = modelresults_df2,
  "Group 3 Model" = modelresults_df3,
  "Group 4 Model" = modelresults_df4,
  "Outlier Results" = outlier_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")