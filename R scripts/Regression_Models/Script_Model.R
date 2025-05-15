# 1. Load libraries
library(lme4)
library(lmerTest)  # For p-values
library(broom.mixed)  # For tidy outputs
library(ggplot2)
library(dplyr)
library(readr)

# 2. Load dataset
df_final <- read_csv("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/df_final_cleaned.csv")

# Clean as character
df_final$CONSTRUCTION_AGE_BAND_clean <- as.character(df_final$CONSTRUCTION_AGE_BAND)

unique_vals <- unique(df_final$CONSTRUCTION_AGE_BAND_clean)
sort(unique_vals)

# Define recoding
df_final$CONSTRUCTION_AGE_BAND_5grp <- case_when(
  df_final$CONSTRUCTION_AGE_BAND_clean %in% c(
    "England and Wales: before 1900"
  ) | df_final$CONSTRUCTION_AGE_BAND_clean %in% as.character(1700:1899) ~ "Pre-1900",
  
  df_final$CONSTRUCTION_AGE_BAND_clean %in% c(
    "England and Wales: 1900-1929", "England and Wales: 1930-1949"
  ) | df_final$CONSTRUCTION_AGE_BAND_clean %in% as.character(1900:1945) ~ "1900–1945",
  
  df_final$CONSTRUCTION_AGE_BAND_clean %in% c(
    "England and Wales: 1950-1966", "England and Wales: 1967-1975", "England and Wales: 1976-1982"
  ) | df_final$CONSTRUCTION_AGE_BAND_clean %in% as.character(1946:1982) ~ "1946–1982",
  
  df_final$CONSTRUCTION_AGE_BAND_clean %in% c(
    "England and Wales: 1983-1990", "England and Wales: 1991-1995", "England and Wales: 1996-2002"
  ) | df_final$CONSTRUCTION_AGE_BAND_clean %in% as.character(1983:1999) ~ "1983–1999",
  
  df_final$CONSTRUCTION_AGE_BAND_clean %in% c(
    "England and Wales: 2003-2006", "England and Wales: 2007-2011", "England and Wales: 2007 onwards"
  ) | df_final$CONSTRUCTION_AGE_BAND_clean %in% as.character(2000:2011) ~ "2000–2011",
  
  df_final$CONSTRUCTION_AGE_BAND_clean %in% c(
    "England and Wales: 2012 onwards"
  ) | df_final$CONSTRUCTION_AGE_BAND_clean %in% as.character(2012:2035) ~ "Post-2012",
  
  is.na(df_final$CONSTRUCTION_AGE_BAND_clean) |
    df_final$CONSTRUCTION_AGE_BAND_clean %in% c("NO DATA!", "INVALID!") ~ "Unknown"
)


# Save to CSV
write_csv(df_final, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/df_final_cleaned_R.csv")

library(dplyr)
library(readr)

# Define grouping variables
group_vars <- c(
  "PROPERTY_TYPE", "BUILT_FORM", "CONSTRUCTION_AGE_BAND_5grp", "CURRENT_ENERGY_RATING",
  "has_double_glazing", "has_efficient_windows", "has_efficient_walls", "has_efficient_roof",
  "has_efficient_hot_water", "has_efficient_mainheat", "has_efficient_lighting",
  "has_mains_gas", "has_condensing_boiler", "has_solar_water", "has_pv",
  "AreaName", "Region"
)

# Keep only those that exist in df
group_vars <- group_vars[group_vars %in% colnames(df_final)]

# Generate table
descriptive_table <- purrr::map_dfr(group_vars, function(var) {
  df_final %>%
    group_by(.data[[var]]) %>%
    summarise(
      mean_price_real = mean(price_real, na.rm = TRUE),
      sd_price_real = sd(price_real, na.rm = TRUE),
      mean_log_price_real = mean(log_price_real, na.rm = TRUE),
      sd_log_price_real = sd(log_price_real, na.rm = TRUE),
      n = n(),
      .groups = "drop"
    ) %>%
    mutate(
      grouping_variable = var,
      level = as.character(.data[[var]])
    ) %>%
    select(grouping_variable, level, everything(), -all_of(var))
})

# Save
write_csv(descriptive_table, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/univariatetests/disaggregated_means_sds.csv")

colnames(df_final)


# Select only model-relevant columns
df_final <- df_final %>%
  select(
    log_price_real,
    has_double_glazing, has_efficient_windows, has_efficient_walls, has_efficient_roof,
    has_efficient_hot_water, has_efficient_mainheat, has_efficient_lighting, has_mains_gas,
    has_solar_water, has_pv,
    CONSTRUCTION_AGE_BAND_5grp,
    log_TOTAL_FLOOR_AREA,
    PROPERTY_TYPE, BUILT_FORM, quarter,
    postcode_area
  )

# Ensure all relevant columns are factors
factor_vars <- c(
  "PROPERTY_TYPE", "BUILT_FORM", "quarter", "postcode_area",
  "has_double_glazing", "has_efficient_windows", "has_efficient_walls", "has_efficient_roof",
  "has_efficient_hot_water", "has_efficient_mainheat", "has_efficient_lighting", "has_mains_gas",
  "has_solar_water", "has_pv", "CONSTRUCTION_AGE_BAND_5grp"
)

df_final[factor_vars] <- lapply(df_final[factor_vars], as.factor)

memory.limit(size = 16000)
gc()                      # Call again to free additional memory

# 4. Fit LMM
lmm_model_all <- lmer(
  log_price_real ~ has_double_glazing + has_efficient_windows + has_efficient_walls + has_efficient_roof +
    has_efficient_hot_water + has_efficient_mainheat + has_efficient_lighting + has_mains_gas +
    has_solar_water + has_pv + CONSTRUCTION_AGE_BAND_5grp +
    log_TOTAL_FLOOR_AREA + PROPERTY_TYPE + BUILT_FORM + quarter + (1 | postcode_area),
  data = df_final,
  REML = FALSE
)

# 5. Summary
summary(lmm_model_all)

#Save Model
saveRDS(lmm_model_all, file = "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/lmm_model_all.rds")

lmm_model_all <- readRDS("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/lmm_model_all.rds")

# 6. Save tidy Coefficient table
coef_table <- broom.mixed::tidy(lmm_model_all, effects = "fixed")
write_csv(coef_table, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/lmm_results/lmm_coefficients_r.csv")

# 7. Goodness-of-fit statistics (APA)
cat("\n📊 Model Fit Statistics (APA style):\n")
cat("- Log-Likelihood:", logLik(lmm_model_all)[1], "\n")
cat("- AIC:", AIC(lmm_model_all), "\n")
cat("- BIC:", BIC(lmm_model_all), "\n")
cat("- Number of Observations:", nobs(lmm_model_all), "\n")
cat("- Number of Groups (postcode_area):", length(unique(df_final$postcode_area)), "\n")

# 8. Feature Importance Plot
# Clean and filter coefficients
library(stringr)
library(ggplot2)
library(dplyr)

# Clean and filter coefficients
coef_table_clean <- coef_table %>%
  filter(term != "(Intercept)") %>%  # Remove intercept
  filter(
    !str_detect(term, "quarter|PROPERTY_TYPE|BUILT_FORM|log_TOTAL_FLOOR_AREA|CONSTRUCTION_AGE_BAND_5grp")
  ) %>%
  mutate(
    term_clean = str_replace_all(term, "TRUE", ""),
    term_clean = str_replace_all(term_clean, "_", " "),
    abs_estimate = abs(estimate),
    sig = case_when(
      p.value < 0.001 ~ "***",
      p.value < 0.01 ~ "**",
      p.value < 0.05 ~ "*",
      p.value < 0.1 ~ ".",
      TRUE ~ ""
    ),
    label = paste0(formatC(estimate, format = "f", digits = 3), sig)
  ) %>%
  arrange(abs_estimate)

# Plot
ggplot(coef_table_clean, aes(x = reorder(term_clean, abs_estimate), y = estimate)) +
  geom_col(fill = "gray40") +
  geom_text(aes(label = label), hjust = ifelse(coef_table_clean$estimate > 0, -0.1, 1.1), size = 3.2) +
  coord_flip() +
  labs(
    title = "Feature Importance (Coefficient Magnitude)",
    x = "Predictor",
    y = "Coefficient Estimate"
  ) +
  theme_minimal() +
  theme(plot.title = element_text(size = 14, face = "bold")) +
  expand_limits(y = c(min(coef_table_clean$estimate) - 0.05, max(coef_table_clean$estimate) + 0.05))


ggsave("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/lmm_results/feature_importance_cleaned.png", dpi = 300)


# 9. Residual Analysis

used_rows <- as.numeric(rownames(model.frame(lmm_model_all)))


df_model_used <- df_final[used_rows, ]

df_model_used$fitted <- predict(lmm_model_all)
df_model_used$residuals <- df_model_used$log_price_real - df_model_used$fitted

# Subsample for residual plots
set.seed(42)
df_sample <- df_model_used %>% sample_n(500000)

# Residuals vs Fitted
ggplot(df_sample, aes(x = fitted, y = residuals)) +
  geom_point(alpha = 0.2) +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
  labs(title = "Residuals vs Fitted Values",
       x = "Fitted log(Price)",
       y = "Residuals") +
  theme_minimal()

ggsave("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/lmm_results/residuals_vs_fitted_r.png", dpi = 300)

# Histogram of Residuals
ggplot(df_model_used, aes(x = residuals)) +
  geom_histogram(bins = 100, color = "black", fill = "lightblue") +
  labs(title = "Histogram of Residuals",
       x = "Residuals",
       y = "Count") +
  theme_minimal()


ggsave("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/lmm_results/residuals_histogram_r.png", dpi = 300)

# 10. Additional Model Diagnostics

# 10.1 Intra-Class Correlation (ICC)
library(performance)
icc_val <- icc(lmm_model_all)
print(icc_val)

# 10.2 Marginal and Conditional R²
r2_vals <- r2_nakagawa(lmm_model_all)
print(r2_vals)

# 10.3 Collinearity (VIF)
vif_vals <- check_collinearity(lmm_model_all)
print(vif_vals)

# Convert VIF results to dataframe
vif_df <- as.data.frame(vif_vals)

# Export to CSV
write_csv(vif_df, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/lmm_results/vif_table.csv")


# Define the list of variables to summarize
vars_to_check <- c(
  "has_double_glazing", "has_efficient_windows", "has_efficient_walls", "has_efficient_roof",
  "has_efficient_hot_water", "has_efficient_mainheat", "has_efficient_lighting", "has_mains_gas",
  "has_condensing_boiler", "has_solar_water", "has_pv",
  "PROPERTY_TYPE", "BUILT_FORM", "quarter", "CONSTRUCTION_AGE_BAND_5grp"
)

# Function to get N and % for each factor variable
freq_summary <- function(df, var) {
  df %>%
    count(!!sym(var)) %>%
    mutate(
      variable = var,
      percent = 100 * n / sum(n)
    ) %>%
    rename(level = !!sym(var)) %>%
    select(variable, level, n, percent)
}

# Apply across all variables
freq_tables <- bind_rows(lapply(vars_to_check, function(v) freq_summary(df_final, v)))

# Save to CSV
write_csv(freq_tables, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/descriptives/predictor_frequencies.csv")

# Preview
print(freq_tables)




library(lme4)
library(broom.mixed)
library(dplyr)
library(readr)

# Assuming df_final is already loaded and preprocessed
# Create output folder if it doesn't exist
output_dir <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/models"
dir.create(output_dir, showWarnings = FALSE)

# Prepare combinations
subgroup_df <- df_final %>%
  filter(!is.na(PROPERTY_TYPE) & !is.na(postcode_area)) %>%
  count(PROPERTY_TYPE, postcode_area) %>%
  filter(n >= 500)  # Optionally skip low-N groups

# Initialize log list
log_list <- list()

# Loop over each subgroup
for (i in seq_len(nrow(subgroup_df))) {
  this_type <- subgroup_df$PROPERTY_TYPE[i]
  this_post <- subgroup_df$postcode_area[i]
  n_rows <- subgroup_df$n[i]
  
  message("⏳ Running model [", i, "/", nrow(subgroup_df), "] — ", this_type, " + ", this_post, " (N = ", n_rows, ")")
  
  df_sub <- df_final %>%
    filter(PROPERTY_TYPE == this_type, postcode_area == this_post)
  
  t0 <- Sys.time()
  
  model <- tryCatch({
    lmer(
      log_price_real ~ has_double_glazing + has_efficient_windows + has_efficient_walls + has_efficient_roof +
        has_efficient_hot_water + has_efficient_mainheat + has_efficient_lighting + has_mains_gas +
        has_solar_water + has_pv + CONSTRUCTION_AGE_BAND_5grp +
        log_TOTAL_FLOOR_AREA + PROPERTY_TYPE + BUILT_FORM + quarter,
      data = df_sub,
      REML = FALSE
    )
  }, error = function(e) {
    message("❌ Error: ", e$message)
    return(NULL)
  })
  
  t1 <- Sys.time()
  duration <- round(as.numeric(difftime(t1, t0, units = "mins")), 2)
  
  if (!is.null(model)) {
    file_prefix <- file.path(output_dir, paste0("model_", this_type, "_", this_post))
    saveRDS(model, paste0(file_prefix, ".rds"))
    write_csv(tidy(model, effects = "fixed"), paste0(file_prefix, "_coef.csv"))
  }
  
  log_list[[i]] <- data.frame(
    property_type = this_type,
    postcode_area = this_post,
    n = n_rows,
    minutes = duration,
    status = ifelse(is.null(model), "error", "success"),
    stringsAsFactors = FALSE
  )
  
  gc()
}

# Save full log
log_df <- bind_rows(log_list)
write_csv(log_df, file.path(output_dir, "lmm_subgroup_log.csv"))
message("\n✅ All models completed. Log saved.")





















unique_vals <- unique(df_final$CONSTRUCTION_AGE_BAND_5grp)
sort(unique_vals)

df_final$CONSTRUCTION_AGE_BAND_5grp <- factor(df_final$CONSTRUCTION_AGE_BAND_5grp,
                                              levels = c("Pre-1900", "1900–1945", "1946–1982",
                                                         "1983–1999", "2000–2011", "Post-2012", "Unknown"),
                                              ordered = FALSE)



library(broom)
library(readr)
library(dplyr)

# Subgrupos desejados
targets <- list(
  list(property_type = "Flat", postcode_area = "SW")
)

# Caminho para salvar resultados
output_dir <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/georgemoss1234/models"
dir.create(output_dir, showWarnings = FALSE)

str(df_final)
# Log
log_list <- list()

for (i in seq_along(targets)) {
  pt <- targets[[i]]$property_type

  pc <- targets[[i]]$postcode_area
  
  df_sub <- df_final %>%
    filter(PROPERTY_TYPE == pt, postcode_area == pc)
  
  n_obs <- nrow(df_sub)
  message("\n⏳ Running fixed-effects model for ", pt, " + ", bf, " @ ", pc, " (n = ", n_obs, ")")
  
  t0 <- Sys.time()
  
  model <- tryCatch({
    lm(
      log_price_real ~ has_double_glazing + has_efficient_windows + has_efficient_walls + has_efficient_roof +
        has_efficient_hot_water + has_efficient_mainheat + has_efficient_lighting + has_mains_gas +
        has_solar_water + has_pv + CONSTRUCTION_AGE_BAND_5grp +
        log_TOTAL_FLOOR_AREA +  quarter,
      data = df_sub
    )
  }, error = function(e) {
    message("❌ Model failed: ", e$message)
    return(NULL)
  })
  
  t1 <- Sys.time()
  duration <- round(as.numeric(difftime(t1, t0, units = "mins")), 2)
  
  if (!is.null(model)) {
    prefix <- paste0("lm_", gsub(" ", "", pt), "_", gsub("-", "", bf), "_", pc)
    
    write_csv(tidy(model), file.path(output_dir, paste0(prefix, "_coef.csv")))
    
    # Salva output de summary no console
    message("\n📄 Summary for ", prefix, ":")
    print(summary(model))
    
    status <- "success"
  } else {
    status <- "error"
  }
  
  log_list[[i]] <- data.frame(
    property_type = pt,

    postcode_area = pc,
    n = n_obs,
    minutes = duration,
    status = status
  )
  # Calcular VIF
  library(performance)
  vif_vals <- check_collinearity(model)
  
  print(vif_vals)
  
  
  
  gc()
}

# Export log
log_df <- bind_rows(log_list)
print(log_df)

colnames(df_final)
