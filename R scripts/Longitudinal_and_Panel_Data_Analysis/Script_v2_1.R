# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/bianalkahsem")

# Load required libraries
library(readr)
library(dplyr)

# Read the dataset
df_all <- read_csv("Final_Merged_Dataset.csv")

# Clean column names
names(df_all) <- tolower(names(df_all))
names(df_all) <- gsub("[ /%-]+", ".", names(df_all))
names(df_all) <- gsub("\\.+", ".", names(df_all))
names(df_all) <- gsub("\\.$", "", names(df_all))

# Standardize text fields
df_all$country <- tolower(trimws(df_all$country))
df_all$code <- trimws(df_all$code)
df_all$year <- as.numeric(df_all$year)

# Drop rows with missing key identifiers
df_all <- df_all %>% filter(!is.na(code) & !is.na(year))

# Ensure uniqueness between code and year
df_all <- df_all %>% arrange(code, year) %>% distinct(code, year, .keep_all = TRUE)


colnames(df_all)

# Rename variables for consistency
df_all <- df_all %>%
  rename(
    rule.of.law = rule_of_law,
    property.rights = property_rights,
    freedom.of.expression.index = freedom_of_expression_index,
    civil.liberties.index = civil_liberties_index,
    gender.equality.in.respect.for.civil.liberties = gender_equality_in_respect_for_civil_liberties,
    GDP_BILLIONS = gdp.billions,
    TRADE...GDP = trade_._gdp,
    population_thousands = population_thousands,
    natural_resources_index = natural.resources,
    territory_km2 = territory_km2
  )

colnames(df_all)
# Explicit numeric conversion correcting for commas (if necessary)
df_all$GDP_BILLIONS <- as.numeric(gsub(",", ".", df_all$GDP_BILLIONS))

# Define variables to be differenced
vars_to_diff <- c(
  "GDP_BILLIONS", "hdi", "TRADE...GDP",
  "rule.of.law", "property.rights", "freedom.of.expression.index", 
  "civil.liberties.index", "gender.equality.in.respect.for.civil.liberties",
  "population_thousands", "natural_resources_index", "territory_km2"
)

# Ensure numeric format explicitly
df_all[vars_to_diff] <- lapply(df_all[vars_to_diff], function(x) as.numeric(as.character(x)))

str(df_all)

# Apply corrected percentage differencing (with proper NA handling)
differenced_df <- df_all %>%
  arrange(code, year) %>%
  group_by(code) %>%
  mutate(across(
    all_of(vars_to_diff),
    ~ ifelse(is.na(.) | is.na(lag(.)) | lag(.) == 0, NA, ((. - lag(.)) / lag(.) * 100)),
    .names = "d_{.col}"
  )) %>%
  ungroup()

# Check results for Brazil
differenced_df %>%
  filter(code == "BRA") %>%
  select(year, GDP_BILLIONS, d_GDP_BILLIONS) %>%
  head(10)


colnames(differenced_df)

# Define dependent variables
dvs <- c("d_GDP_BILLIONS", "d_hdi", "d_TRADE...GDP")

independent_vars <- c(
  "d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
  "d_civil.liberties.index", "d_gender.equality.in.respect.for.civil.liberties",
  "d_population_thousands", "d_natural_resources_index", "territory_km2"
)

# Outlier Evaluation

library(dplyr)

calculate_z_scores <- function(data, vars, id_vars = c("country", "year"), z_threshold = 4) {
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
vars_to_check <- c("d_GDP_BILLIONS", "d_hdi", "d_TRADE...GDP", "d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
                      "d_civil.liberties.index", "d_gender.equality.in.respect.for.civil.liberties",
                   "d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
                   "d_civil.liberties.index", "d_gender.equality.in.respect.for.civil.liberties",
                   "d_population_thousands", "d_natural_resources_index", "territory_km2")

# Apply the function to detect outliers
outlier_results <- calculate_z_scores(data = differenced_df, vars = vars_to_check)

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
        dplyr::select(country, year)
      
      
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
# Define updated outliers as a list of data frames
outlier_list <- list(
  d_GDP_BILLIONS = data.frame(
    country = c("angola", "armenia", "democratic republic of the congo", "venezuela", "venezuela"),
    year = c(1996, 1994, 1994, 2018, 2019)
  ),
  d_hdi = data.frame(
    country = c("afghanistan", "afghanistan", "armenia", "azerbaijan", "burundi", "burundi", "burundi",
                "central african republic", "central african republic", "republic of the congo", "ecuador",
                "eritrea", "ethiopia", "haiti", "haiti", "iraq", "iraq", "kuwait", "liberia", "liberia",
                "libya", "burma/myanmar", "malawi", "rwanda", "rwanda", "sudan", "sudan", "sierra leone",
                "south sudan", "syria", "syria", "chad", "tajikistan", "tajikistan", "yemen", "zimbabwe"),
    year = c(1995, 2002, 1992, 2020, 1993, 1994, 2006, 1992, 2013, 1997, 2020, 2015, 2005, 2010, 2011, 1991,
             1992, 1991, 2000, 2003, 2011, 2009, 1995, 2002, 2004, 1994, 1999, 2000, 2012, 2012, 2013, 2004,
             1992, 1993, 2015, 2009)
  ),
  d_TRADE...GDP = data.frame(
    country = "iraq",
    year = 1997
  ),
  d_rule.of.law = data.frame(
    country = c("angola", "armenia", "azerbaijan", "azerbaijan", "burundi", "central african republic",
                "republic of the congo", "republic of the congo", "dominican republic", "eritrea", "georgia",
                "gambia, the", "croatia", "haiti", "indonesia", "indonesia", "iraq", "kyrgyzstan", "liberia",
                "liberia", "libya", "madagascar", "madagascar", "maldives", "maldives", "burma/myanmar",
                "burma/myanmar", "mauritania", "peru", "paraguay", "sudan", "sudan", "sudan", "sierra leone",
                "el salvador", "somalia", "somalia", "tajikistan", "tajikistan", "timor-leste", "tunisia",
                "uzbekistan"),
    year = c(2017, 2018, 1992, 1994, 2019, 2014, 1991, 1997, 2020, 1991, 2004, 2017, 2000, 1995, 1998, 1999,
             2003, 2010, 1991, 2004, 2011, 2009, 2014, 2008, 2019, 2010, 2011, 2006, 2001, 1992, 2005, 2017,
             2019, 1996, 1992, 1992, 2013, 1992, 2018, 2000, 2011, 2017)
  ),
  d_property.rights = data.frame(
    country = "albania",
    year = 1991
  ),
  d_freedom.of.expression.index = data.frame(
    country = c("afghanistan", "afghanistan", "angola", "albania", "burundi", "republic of the congo", "eritrea",
                "ethiopia", "gambia, the", "indonesia", "iraq", "cambodia", "cambodia", "libya", "malawi",
                "malawi", "el salvador", "somalia", "tajikistan", "timor-leste", "timor-leste", "tunisia",
                "uzbekistan"),
    year = c(2001, 2002, 1992, 1991, 1992, 1991, 1991, 1991, 2017, 1998, 2003, 1992, 1993, 2011, 1993, 1994,
             1992, 1991, 1998, 1998, 2000, 2011, 2017)
  ),
  d_civil.liberties.index = data.frame(
    country = c("afghanistan", "afghanistan", "angola", "albania", "belarus", "republic of the congo", "eritrea",
                "ethiopia", "gambia, the", "haiti", "indonesia", "indonesia", "iraq", "cambodia", "cambodia",
                "liberia", "libya", "burma/myanmar", "malawi", "malawi", "peru", "sudan", "el salvador",
                "somalia", "somalia", "timor-leste", "tunisia", "uzbekistan", "yemen", "south africa"),
    year = c(2001, 2002, 1992, 1991, 2020, 1991, 1991, 1991, 2017, 1995, 1998, 1999, 2003, 1992, 1993, 2004,
             2011, 2010, 1994, 1995, 2001, 2019, 1992, 1991, 1992, 2000, 2011, 2017, 1991, 1994)
  ),
  d_gender.equality.in.respect.for.civil.liberties = data.frame(
    country = c("bahrain", "central african republic", "haiti", "lebanon", "burma/myanmar", "mauritania",
                "nepal", "rwanda", "tajikistan", "tajikistan", "kosovo", "yemen", "south africa"),
    year = c(2020, 2020, 2019, 2000, 2016, 2019, 1992, 1994, 2010, 2020, 2005, 2005, 1994)
  ),
  d_population_thousands = data.frame(
    country = c("united arab emirates", "bahrain", "bahrain", "bahrain", "bosnia and herzegovina", 
                "bosnia and herzegovina", "bosnia and herzegovina", "bosnia and herzegovina", 
                "equatorial guinea", "equatorial guinea", "equatorial guinea", "jordan", "jordan", 
                "jordan", "jordan", "kuwait", "kuwait", "kuwait", "kuwait", "lebanon", "lebanon", 
                "lebanon", "lebanon", "lebanon", "lithuania", "lithuania", "lithuania", "lithuania", 
                "maldives", "maldives", "maldives", "maldives", "maldives", "maldives", "oman", 
                "oman", "oman", "oman", "oman", "oman", "qatar", "qatar", "qatar", "qatar", "qatar"),
    year = c(2011, 2017, 2018, 2019, 2012, 2013, 2014, 2015, 2011, 2012, 2013, 2011, 2012, 2013, 2014, 
             2011, 2012, 2013, 2014, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2012, 2013, 
             2014, 2015, 2016, 2017, 2011, 2012, 2013, 2014, 2015, 2016, 2011, 2012, 2013, 2014, 2015)
  ),
  d_natural_resources_index = data.frame(
    country = c("bosnia and herzegovina", "bosnia and herzegovina", "cyprus", "cyprus", "eritrea", 
                "ireland", "israel", "jordan", "luxembourg", "morocco", "north macedonia", "sudan", 
                "sudan", "sudan", "kosovo"),
    year = c(1995, 1997, 2007, 2017, 2011, 2006, 2004, 2008, 2008, 2008, 1992, 1992, 1995, 1999, 2004)
  )
)




# Function to replace flagged outliers with NA for specific variables
replace_outliers_with_na <- function(df, outlier_list) {
  for (var in names(outlier_list)) {
    # Ensure column exists in df
    if (var %in% names(df)) {
      flagged <- outlier_list[[var]]
      for (i in seq_len(nrow(flagged))) {
        country <- flagged$country[i]
        year <- flagged$year[i]
        # Replace the specific variable value with NA
        df[df$country == country & df$year == year, var] <- NA
      }
    } else {
      warning(paste("Variable", var, "not found in dataset"))
    }
  }
  return(df)
}


# Apply the function to each dataset in dfs
differenced_df_nooutlier <- replace_outliers_with_na(differenced_df, outlier_list)

# Apply the function to differenced_df
differenced_df <- replace_outliers_with_na(differenced_df, outlier_list)


# Define dependent variables
dvs <- c("d_GDP_BILLIONS", "d_hdi", "d_TRADE...GDP")


independent_vars <- c(
  "d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
  "d_civil.liberties.index", "d_gender.equality.in.respect.for.civil.liberties",
  "d_population_thousands", "d_natural_resources_index", "territory_km2"
)

# Overall Results

# Load necessary libraries
library(plm)
library(lmtest)
library(sandwich)
library(dplyr)
library(car)

# Ensure differenced_df is in the correct panel data format
differenced_df <- pdata.frame(differenced_df, index = c("code", "year"))


# List to store model results and diagnostics
results_list_all <- list()

# Loop through each dependent variable
for (dv in dvs) {
  # Define the formula
  formula_string <- paste0(dv, " ~ ", paste(independent_vars, collapse = " + "))
  
  # Fit the fixed effects model
  fe_model <- plm(as.formula(formula_string), data = differenced_df, model = "pooling", na.action = na.exclude)
  
  
  robust_se <- vcovBK(fe_model, cluster = "group")
  
  # Summarize the model with robust standard errors
  fe_summary <- summary(fe_model, vcov = robust_se)
  
  # Extract model coefficients
  model_coefficients <- as.data.frame(fe_summary$coefficients)
  colnames(model_coefficients) <- c("Estimate", "Std. Error", "t-value", "Pr(>|t|)")
  
  # Perform diagnostics
  dw_test <- pdwtest(fe_model, alternative = "two.sided")
  bp_test <- bptest(fe_model)
  vif_test <- vif(lm(as.formula(formula_string), data = differenced_df))
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
  fe_model <- plm(as.formula(formula_string), data = differenced_df, model = "within", na.action = na.exclude)
  
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


# Take out civil liberties to resolve multicolinearity

# Load necessary libraries
library(plm)
library(lmtest)
library(sandwich)
library(dplyr)
library(car)

# Define dependent variables
dvs <- c("d_GDP_BILLIONS", "d_hdi", "d_TRADE...GDP")

independent_vars <- c(
  "d_rule.of.law", "d_property.rights", "d_freedom.of.expression.index", 
   "d_gender.equality.in.respect.for.civil.liberties",
  "d_population_thousands", "d_natural_resources_index", "territory_km2"
)



# List to store model results and diagnostics
results_list_all_nocol <- list()

# Loop through each dependent variable
for (dv in dvs) {
  # Define the formula
  formula_string <- paste0(dv, " ~ ", paste(independent_vars, collapse = " + "))
  
  # Fit the fixed effects model
  fe_model <- plm(as.formula(formula_string), data = differenced_df, model = "within", na.action = na.exclude)
  
  robust_se <- vcovBK(fe_model, cluster = "group")
  
  # Summarize the model with robust standard errors
  fe_summary <- summary(fe_model, vcov = robust_se)
  
  # Extract model coefficients
  model_coefficients <- as.data.frame(fe_summary$coefficients)
  colnames(model_coefficients) <- c("Estimate", "Std. Error", "t-value", "Pr(>|t|)")
  
  # Perform diagnostics
  dw_test <- pdwtest(fe_model, alternative = "two.sided")
  bp_test <- bptest(fe_model)
  vif_test <- vif(lm(as.formula(formula_string), data = differenced_df))
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
coefficients_df_nocol <- data.frame()

# Loop through each dependent variable
for (dv in dvs) {
  # Define the formula
  formula_string <- paste0(dv, " ~ ", paste(independent_vars, collapse = " + "))
  
  # Fit the fixed effects model
  fe_model <- plm(as.formula(formula_string), data = differenced_df, model = "within", na.action = na.exclude)
  
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
  coefficients_df_nocol <- rbind(coefficients_df_nocol, model_coefficients)
}

# Reset row names for the combined dataframe
rownames(coefficients_df_nocol) <- NULL

# Save or display the combined coefficients dataframe
print(coefficients_df_nocol)


# Convert pdata.frame to regular data frame
plot_data <- as.data.frame(lapply(differenced_df, function(col) {
  if (inherits(col, "pseries")) as.numeric(col) else col
}))

library(ggplot2)

# Directory to save scatterplots
output_dir <- "scatterplots_combined"
if (!dir.exists(output_dir)) dir.create(output_dir)

# Create scatterplots for each dependent variable
for (dv in dvs) {
  for (iv in independent_vars) {
    # Generate the scatterplot without grouping
    p <- ggplot(plot_data, aes_string(x = iv, y = dv)) +
      geom_point(alpha = 0.6) +
      geom_smooth(method = "lm", se = FALSE, linetype = "dashed", formula = y ~ x) +
      labs(
        title = paste("Scatterplot of", dv, "vs.", iv),
        x = iv,
        y = dv
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(hjust = 0.5, size = 14),
        axis.title = element_text(size = 12)
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


# Calculate mean scores for each year across all countries
calculate_means <- function(data, vars_to_average, id_var = "year") {
  data %>%
    group_by(across(all_of(id_var))) %>%
    summarize(across(all_of(vars_to_average), mean, na.rm = TRUE), .groups = "drop")
}

# List of variables to calculate mean scores
vars_to_average <- c("d_GDP_BILLIONS", "d_hdi", "d_TRADE...GDP")

# Convert pdata.frame to regular data frame if needed
differenced_df <- as.data.frame(lapply(differenced_df, function(col) {
  if (inherits(col, "pseries")) as.numeric(col) else col
}))

# Calculate mean scores from the single differenced dataset
mean_scores <- calculate_means(differenced_df, vars_to_average)

# Ensure year is numeric for plotting
mean_scores$year <- as.numeric(as.character(mean_scores$year))

# Directory to save the line charts
output_dir <- "line_charts_mean_scores"
if (!dir.exists(output_dir)) dir.create(output_dir)

# Create line charts for each variable
for (var in vars_to_average) {
  # Generate the line chart
  p <- ggplot(mean_scores, aes(x = year, y = !!sym(var))) +
    geom_line(size = 1, color = "steelblue") +
    geom_point(size = 2, color = "steelblue") +
    labs(
      title = paste("Mean Scores of", var, "by year"),
      x = "year",
      y = paste("Mean", var)
    ) +
    theme_minimal() +
    theme(
      plot.title = element_text(hjust = 0.5, size = 14),
      axis.title = element_text(size = 12)
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
  "Clean Diff Data" = differenced_df, 
  "Model Results" = coefficients_df
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_NewModel.xlsx")










# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/bianalkahsem/New Order")

# Load required libraries
library(readr)
library(dplyr)
library(openxlsx)

df_all <- read.xlsx("v dem continous data new variables _ ECONOMIC VARIABLES NOV 2024 (2).xlsx")


# Clean column names
names(df_all) <- tolower(names(df_all))
names(df_all) <- gsub("[ /%-]+", ".", names(df_all))
names(df_all) <- gsub("\\.+", ".", names(df_all))
names(df_all) <- gsub("\\.$", "", names(df_all))

# Standardize text fields
df_all$country <- tolower(trimws(df_all$country))
df_all$year <- as.numeric(df_all$year)

# Drop rows with missing key identifiers
df_all <- df_all %>% filter(!is.na(country) & !is.na(year))

# Ensure uniqueness between code and year
df_all <- df_all %>% arrange(country, year) %>% distinct(country, year, .keep_all = TRUE)



# Load required packages
library(dplyr)
library(plm)
library(lmtest)
library(sandwich)
library(zoo)


# Compute differenced outcomes and all lagged democracy indicators
df_diff <- df_filtered %>%
  arrange(country, year) %>%
  group_by(country) %>%
  mutate(
    d_gdp     = ifelse(!is.na(dplyr::lag(gdp.billions)) & dplyr::lag(gdp.billions) != 0,
                       (gdp.billions - dplyr::lag(gdp.billions)) / dplyr::lag(gdp.billions) * 100, NA_real_),
    d_hdi     = ifelse(!is.na(dplyr::lag(hdi)) & dplyr::lag(hdi) != 0,
                       (hdi - dplyr::lag(hdi)) / dplyr::lag(hdi) * 100, NA_real_),
    d_trade   = ifelse(!is.na(dplyr::lag(trade.gdp)) & dplyr::lag(trade.gdp) != 0,
                       (trade.gdp - dplyr::lag(trade.gdp)) / dplyr::lag(trade.gdp) * 100, NA_real_),
    d_v2x     = v2x_regime_amb - dplyr::lag(v2x_regime_amb),
    lag1_v2x  = dplyr::lag(v2x_regime_amb, 1),
    ma3_v2x   = rowMeans(cbind(
      dplyr::lag(v2x_regime_amb, 0),
      dplyr::lag(v2x_regime_amb, 1),
      dplyr::lag(v2x_regime_amb, 2)
    ), na.rm = FALSE),
    cumlag2_v2x = dplyr::lag(v2x_regime_amb, 1) + dplyr::lag(v2x_regime_amb, 2),
    cumlag3_v2x = dplyr::lag(v2x_regime_amb, 1) + dplyr::lag(v2x_regime_amb, 2) + dplyr::lag(v2x_regime_amb, 3)
  ) %>%
  ungroup()

# Convert to panel data
df_panel <- pdata.frame(df_diff, index = c("country", "year"))

# Define function to run fixed effects model with robust standard errors
run_model <- function(dv, iv) {
  formula <- as.formula(paste(dv, "~", iv))
  model <- plm(formula, data = df_panel, model = "within", na.action = na.exclude)
  se <- vcovHC(model, type = "HC0", method = "arellano", cluster = "group")
  summary(model, vcov = se)
}

# Define DVs and democracy predictors
dvs <- c("d_gdp", "d_hdi", "d_trade")
predictors <- c("d_v2x", "lag1_v2x", "ma3_v2x", "cumlag2_v2x", "cumlag3_v2x")

colnames(df_diff)

# Run panel models
results <- list()

for (dv in dvs) {
  for (iv in predictors) {
    fml <- as.formula(paste0(dv, " ~ ", iv))
    model <- plm(fml, data = df_panel, model = "within", na.action = na.exclude)
    robust_se <- vcovHC(model, method = "arellano", type = "HC0", cluster = "group")
    summary_model <- summary(model, vcov = robust_se)
    print(summary_model)
    df_out <- as.data.frame(summary_model$coefficients)
    df_out$Predictor <- rownames(df_out)
    df_out$Dependent <- dv
    df_out$RegimeVar <- iv
    
    results[[paste(dv, iv, sep = "_")]] <- df_out
  }
}

# Combine and display results
final_results <- bind_rows(results)
print(final_results)


data_list <- list(
  "Democracy Score Model" = final_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_Dem_Model.xlsx")

