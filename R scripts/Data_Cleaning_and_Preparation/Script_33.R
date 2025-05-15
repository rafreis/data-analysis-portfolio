# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/joshua")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Dataclean.xlsx")

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

# Load required libraries
library(dplyr)
library(ggplot2)
library(psych)

# Step 1: Data Preparation
# Rename key variables for clarity
df <- df %>%
  rename(
    prox_pocus = Prox,
    dist_pocus = Dist,
    prox_formal = Prox.1,
    dist_formal = Dist.1,
    r_max = R_MAX,
    f_max = F_MAX
  )

# Convert numeric variables from character if necessary
numeric_vars <- c("Age", "X.packs", "X.years", "Pack.Years", 
                  "prox_pocus", "dist_pocus", "prox_formal", "dist_formal", 
                  "r_max", "f_max")

df[numeric_vars] <- lapply(df[numeric_vars], function(x) as.numeric(as.character(x)))

# Step 2: Descriptive Statistics
# Descriptive summary
desc_vars <- c("r_max", "f_max", "prox_pocus", "prox_formal", "dist_pocus", "dist_formal")

descriptive_stats <- describe(df[desc_vars], fast = FALSE)

# Show table
descriptive_stats

library(e1071)

library(e1071) 

calculate_stats <- function(data, variables) {
  results <- data.frame(Variable = character(),
                        Skewness = numeric(),
                        Kurtosis = numeric(),
                        Shapiro_Wilk_F = numeric(),
                        Shapiro_Wilk_p_value = numeric(),
                        KS_Statistic = numeric(),
                        KS_p_value = numeric(),
                        stringsAsFactors = FALSE)
  
  for (var in variables) {
    if (var %in% names(data)) {
      # Calculate skewness and kurtosis
      skew <- skewness(data[[var]], na.rm = TRUE)
      kurt <- kurtosis(data[[var]], na.rm = TRUE)
      
      # Perform Shapiro-Wilk test
      shapiro_test <- shapiro.test(data[[var]])
      
      # Perform Kolmogorov-Smirnov test
      ks_test <- ks.test(data[[var]], "pnorm", mean = mean(data[[var]], na.rm = TRUE), sd = sd(data[[var]], na.rm = TRUE))
      
      # Add results to the dataframe
      results <- rbind(results, c(var, skew, kurt, shapiro_test$statistic, shapiro_test$p.value, ks_test$statistic, ks_test$p.value))
    } else {
      warning(paste("Variable", var, "not found in the data. Skipping."))
    }
  }
  
  colnames(results) <- c("Variable", "Skewness", "Kurtosis", "Shapiro_Wilk_F", "Shapiro_Wilk_p_value", "KS_Statistic", "KS_p_value")
  return(results)
}

scales <- c("r_max"      ,         "f_max", "prox_pocus"     ,     "dist_pocus"     ,     "Date.of.formal"     
            ,"prox_formal")
colnames(df)

# Example usage with your data
df_normality_results <- calculate_stats(df, scales)


# Function to run paired t-tests comparing two columns in the same dataframe
run_paired_t_test_between_columns <- function(data, col1, col2, label = NULL) {
  # Remove rows with NA in either column
  df_clean <- data %>% filter(!is.na(.data[[col1]]) & !is.na(.data[[col2]]))
  
  # Run paired t-test
  t_result <- t.test(df_clean[[col1]], df_clean[[col2]], paired = TRUE)
  
  # Calculate effect size (Cohen's d for paired samples)
  diff_vals <- df_clean[[col2]] - df_clean[[col1]]
  effect_size <- mean(diff_vals) / sd(diff_vals)
  
  # Output summary
  result <- data.frame(
    Variable = ifelse(is.null(label), paste0(col1, "_vs_", col2), label),
    T_Value = t_result$statistic,
    P_Value = t_result$p.value,
    Effect_Size = effect_size,
    Num_Pairs = nrow(df_clean),
    Mean_Diff = mean(diff_vals),
    CI_Low = t_result$conf.int[1],
    CI_High = t_result$conf.int[2]
  )
  
  return(result)
}

# Run paired t-tests for all measurement pairs
paired_test_results <- rbind(
  run_paired_t_test_between_columns(df, "r_max", "f_max", "Resident vs Formal Max"),
  run_paired_t_test_between_columns(df, "prox_pocus", "prox_formal", "Resident vs Formal Prox"),
  run_paired_t_test_between_columns(df, "dist_pocus", "dist_formal", "Resident vs Formal Dist")
)

# View combined results
print(paired_test_results)


library(dplyr)
library(ggplot2)

# Select relevant columns and reshape into long format
df_long <- df %>%
  select(r_max, f_max, prox_pocus, prox_formal, dist_pocus, dist_formal) %>%
  tidyr::pivot_longer(cols = everything(), names_to = "Measurement", values_to = "Value") %>%
  filter(!is.na(Value))

# Calculate mean and 95% CI
summary_stats <- df_long %>%
  group_by(Measurement) %>%
  summarise(
    Mean = mean(Value),
    SD = sd(Value),
    N = n(),
    SE = SD / sqrt(N),
    CI_low = Mean - qt(0.975, df = N - 1) * SE,
    CI_high = Mean + qt(0.975, df = N - 1) * SE
  )

print(summary_stats)


# Order the measurements for better layout
summary_stats$Measurement <- factor(summary_stats$Measurement, levels = summary_stats$Measurement)

ggplot(summary_stats, aes(x = Measurement, y = Mean)) +
  geom_point(size = 3) +
  geom_errorbar(aes(ymin = CI_low, ymax = CI_high), width = 0.2) +
  labs(title = "Mean ± 95% CI for Aortic Measurements",
       x = "Measurement Type",
       y = "Mean (cm)") +
  theme_minimal(base_size = 14) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

library(tidyr)

# Reshape and relabel to separate source and measurement
df_grouped <- df %>%
  select(r_max, f_max, prox_pocus, prox_formal, dist_pocus, dist_formal) %>%
  rename(
    `Max (Resident)` = r_max,
    `Max (Formal)` = f_max,
    `Prox (Resident)` = prox_pocus,
    `Prox (Formal)` = prox_formal,
    `Dist (Resident)` = dist_pocus,
    `Dist (Formal)` = dist_formal
  ) %>%
  pivot_longer(cols = everything(), names_to = "Measurement", values_to = "Value") %>%
  separate(Measurement, into = c("Type", "Source"), sep = " \\(", extra = "merge", fill = "right") %>%
  mutate(Source = gsub("\\)", "", Source))

# Compute stats
summary_stats_grouped <- df_grouped %>%
  filter(!is.na(Value)) %>%
  group_by(Type, Source) %>%
  summarise(
    Mean = mean(Value),
    SD = sd(Value),
    N = n(),
    SE = SD / sqrt(N),
    CI_low = Mean - qt(0.975, df = N - 1) * SE,
    CI_high = Mean + qt(0.975, df = N - 1) * SE,
    .groups = "drop"
  )

# Save the current plot to PNG
ggplot(summary_stats_grouped, aes(x = Type, y = Mean, color = Source)) +
  geom_point(position = position_dodge(width = 0.4), size = 3) +
  geom_errorbar(aes(ymin = CI_low, ymax = CI_high), 
                position = position_dodge(width = 0.4), width = 0.2) +
  labs(title = "Mean ± 95% CI for Resident vs Formal U/S",
       x = "Measurement Type",
       y = "Mean (cm)",
       color = "Source") +
  theme_minimal(base_size = 14) +
  theme(axis.text.x = element_text(angle = 0, hjust = 0.5))

# Export
ggsave("comparison_ci_plot.png", width = 8, height = 6, dpi = 300)


# Bland-Altman plot for r_max vs f_max
df_ba <- df %>%
  filter(!is.na(r_max) & !is.na(f_max)) %>%
  mutate(
    mean_val = (r_max + f_max) / 2,
    diff_val = r_max - f_max
  )

library(irr)

run_agreement_analysis <- function(df, col1, col2, label) {
  # Clean data
  df_sub <- df %>%
    filter(!is.na(.data[[col1]]) & !is.na(.data[[col2]])) %>%
    mutate(
      mean_val = (.data[[col1]] + .data[[col2]]) / 2,
      diff_val = .data[[col1]] - .data[[col2]]
    )
  
  # Calculate stats
  mean_diff <- mean(df_sub$diff_val)
  sd_diff <- sd(df_sub$diff_val)
  loa_upper <- mean_diff + 1.96 * sd_diff
  loa_lower <- mean_diff - 1.96 * sd_diff
  
  # Plot Bland-Altman
  p <- ggplot(df_sub, aes(x = mean_val, y = diff_val)) +
    geom_point(size = 2) +
    geom_hline(yintercept = mean_diff, color = "blue", linetype = "solid") +
    geom_hline(yintercept = loa_upper, color = "red", linetype = "dashed") +
    geom_hline(yintercept = loa_lower, color = "red", linetype = "dashed") +
    labs(title = paste("Bland-Altman Plot:", label),
         x = "Mean of Measurements (cm)",
         y = "Difference (Resident - Formal)") +
    theme_minimal(base_size = 14)
  
  # Save the plot
  filename <- paste0("bland_altman_", gsub(" ", "_", tolower(label)), ".png")
  ggsave(filename, p, width = 8, height = 6, dpi = 300)
  
  # ICC
  icc_data <- df_sub %>% select(all_of(col1), all_of(col2))
  icc_res <- icc(icc_data, model = "twoway", type = "agreement", unit = "single")
  
  cat("\n\n-----------------------------\n")
  cat("ICC Results for:", label, "\n")
  print(icc_res)
}

# Run agreement analysis for all measurement types
run_agreement_analysis(df, "r_max", "f_max", "Max (Resident vs Formal)")
run_agreement_analysis(df, "prox_pocus", "prox_formal", "Prox (Resident vs Formal)")
run_agreement_analysis(df, "dist_pocus", "dist_formal", "Dist (Resident vs Formal)")


# Install if needed: install.packages("DescTools")
library(DescTools)

# Run Lin's CCC for each pair
ccc_rmax <- CCC(df$r_max, df$f_max, na.rm = TRUE)
ccc_prox <- CCC(df$prox_pocus, df$prox_formal, na.rm = TRUE)
ccc_dist <- CCC(df$dist_pocus, df$dist_formal, na.rm = TRUE)

# Print results
cat("\nLin's CCC - R_MAX vs F_MAX:\n")
print(ccc_rmax)

cat("\nLin's CCC - PROX:\n")
print(ccc_prox)

cat("\nLin's CCC - DIST:\n")
print(ccc_dist)

str(ccc_dist)

## Export Results

library(openxlsx)

save_apa_formatted_excel <- function(data_list, filename) {
  wb <- createWorkbook()
  
  for (i in seq_along(data_list)) {
    sheet_name <- names(data_list)[i]
    if (is.null(sheet_name) || sheet_name == "") sheet_name <- paste("Sheet", i)
    addWorksheet(wb, sheet_name)
    
    # Prepare data
    data <- data_list[[i]]
    if (is.matrix(data)) data <- as.data.frame(data)
    
    # Clean rownames if present
    if (!is.null(rownames(data))) {
      data <- cbind("Index" = rownames(data), data)
      rownames(data) <- NULL
    }
    
    # Format numeric columns to 2-3 decimals
    data <- data.frame(lapply(data, function(col) {
      if (is.numeric(col)) round(col, 3) else col
    }), check.names = FALSE)
    
    # Write data to sheet
    writeData(wb, sheet_name, data, startRow = 1, startCol = 1, headerStyle = createStyle(textDecoration = "bold", halign = "center"))
    
    # Define APA header and border styles
    header_style <- createStyle(fontName = "Calibri", fontSize = 11,
                                textDecoration = "bold", halign = "center",
                                border = "TopBottom", borderColour = "black", borderStyle = "thin")
    
    cell_style <- createStyle(fontName = "Calibri", fontSize = 11,
                              halign = "center", valign = "center")
    
    # Apply styles
    addStyle(wb, sheet = sheet_name, style = header_style, rows = 1, cols = 1:ncol(data), gridExpand = TRUE)
    addStyle(wb, sheet = sheet_name, style = cell_style, rows = 2:(nrow(data) + 1), cols = 1:ncol(data), gridExpand = TRUE)
    
    # Add bottom border after last row
    bottom_border_style <- createStyle(border = "bottom", borderColour = "black", borderStyle = "thin")
    addStyle(wb, sheet_name, style = bottom_border_style, rows = nrow(data) + 1, cols = 1:ncol(data), gridExpand = TRUE)
    
    # Auto column widths
    setColWidths(wb, sheet_name, cols = 1:ncol(data), widths = "auto")
  }
  
  saveWorkbook(wb, filename, overwrite = TRUE)
}


# Helper function to extract ICC info from an icclist object
extract_icc_summary <- function(icc_obj, label) {
  data.frame(
    Variable = label,
    ICC = icc_obj$value,
    ICC_Lower = icc_obj$lbound,
    ICC_Upper = icc_obj$ubound,
    F_value = icc_obj$Fvalue,
    df1 = icc_obj$df1,
    df2 = icc_obj$df2,
    p_value = icc_obj$p.value
  )
}

# Run ICC and extract for all measurement types
icc_rmax <- icc(df %>% select(r_max, f_max), model = "twoway", type = "agreement", unit = "single")
icc_prox <- icc(df %>% select(prox_pocus, prox_formal), model = "twoway", type = "agreement", unit = "single")
icc_dist <- icc(df %>% select(dist_pocus, dist_formal), model = "twoway", type = "agreement", unit = "single")

# Combine all into one table
icc_table <- rbind(
  extract_icc_summary(icc_rmax, "Max"),
  extract_icc_summary(icc_prox, "Prox"),
  extract_icc_summary(icc_dist, "Dist")
)

# Function to extract CCC estimates
extract_ccc_summary <- function(ccc_obj, label) {
  data.frame(
    Variable = label,
    CCC = ccc_obj$rho.c$est,
    CCC_Lower = ccc_obj$rho.c$lwr.ci,
    CCC_Upper = ccc_obj$rho.c$upr.ci,
    Accuracy_Cb = ccc_obj$C.b,
    Shift_Location = ccc_obj$l.shift,
    Shift_Scale = ccc_obj$s.shift
  )
}

# Combine all into one table
ccc_table <- rbind(
  extract_ccc_summary(ccc_rmax, "Max"),
  extract_ccc_summary(ccc_prox, "Prox"),
  extract_ccc_summary(ccc_dist, "Dist")
)



# Compose list of tables to export
data_list <- list(
  "Descriptive Statistics" = descriptive_stats,
  "Normality Tests" = df_normality_results,
  "Paired T-Tests" = paired_test_results,
  "Means + 95% CI" = summary_stats_grouped,
  "Lin's CCC" = ccc_table,
  "ICC Results" = icc_table
)

# Save to Excel file
save_apa_formatted_excel(data_list, "AAA_Analysis_Tables_APA.xlsx")
