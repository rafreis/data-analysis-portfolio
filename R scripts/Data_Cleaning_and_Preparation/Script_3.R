# Set working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/Ilza_hilman")


# Define only the final cleaned files to load
final_files <- c(
  "cleaned_FFQ_Total_Polyphenols_&_classes.csv",
  "cleaned_TCA_NutrxFood_final.csv",
  "cleaned_1_FFQ_TCA_Nutrients.csv",
  "cleaned_Recall_Polyphenols_(day_1_and_2).csv",
  "cleaned_FFQ_Kcal,Nutrients_(same_data_as_in_1._FFQ_TCA_Nutrients).csv",
  "cleaned_Recalls_Kcal,_Nutrients_(both_days).csv"
)

# Load and print structure for each
library(readr)
library(dplyr)

for (file in final_files) {
  cat("\n\n---", file, "---\n")
  df <- read_csv(file, show_col_types = FALSE)
  print(glimpse(df))
}

# Load required packages
library(tidyverse)
library(readr)
library(stringr)
library(janitor)

ffq_long <- read_csv("cleaned_1_FFQ_TCA_Nutrients.csv", show_col_types = FALSE)

ffq_long %>% filter(is.na(Nutr_Name)) %>% count(Participant_ID)


# Check for duplicates after full cleaning
library(dplyr)
library(stringr)

ffq_long %>%
  filter(!is.na(Nutr_Name)) %>%
  distinct(Nutr_Name) %>%
  arrange(Nutr_Name) %>%
  pull(Nutr_Name)


ffq_wide <- ffq_long %>%
  filter(!is.na(Nutr_Name)) %>%
  pivot_wider(id_cols = Participant_ID, names_from = Nutr_Name, values_from = Intake) %>%
  rename_with(~ str_replace_all(., "[()]", "")) %>%
  rename_with(~ str_squish(str_replace_all(., " ", "_"))) %>%
  rename_with(~ paste0(., "_FFQ"), -1)


# 2. Load FFQ polyphenols and merge
ffq_poly <- read_csv("cleaned_FFQ_Total_Polyphenols_&_classes.csv", show_col_types = FALSE) %>%
  rename_with(~ str_replace_all(., "[()]", "")) %>%
  rename_with(~ str_squish(str_replace_all(., " ", "_"))) %>%
  rename_with(~ paste0(., "_FFQ"), -1)

ffq_final <- ffq_wide %>%
  left_join(ffq_poly, by = "Participant_ID") %>%
  rename_with(~ str_replace_all(., "[()]", "")) %>%
  rename_with(~ str_squish(str_replace_all(., " ", "_"))) 

# 3. Load and aggregate recall nutrients
recall_nut <- read_csv("cleaned_Recalls_Kcal,_Nutrients_(both_days).csv", show_col_types = FALSE)

recall_nut_avg <- recall_nut %>%
  group_by(Participant_ID) %>%
  summarise(across(where(is.numeric), mean, na.rm = TRUE)) %>%
  rename_with(~ str_replace_all(., "[()]", "")) %>%
  rename_with(~ str_squish(str_replace_all(., " ", "_"))) %>%
  rename_with(~paste0(., "_Recall"), -1)

# 4. Load and aggregate recall polyphenols
recall_poly <- read_csv("cleaned_Recall_Polyphenols_(day_1_and_2).csv", show_col_types = FALSE)

recall_poly_avg <- recall_poly %>%
  group_by(Participant_ID) %>%
  summarise(across(where(is.numeric), ~ mean(.x, na.rm = TRUE))) %>%
  rename_with(~ str_replace_all(., "[()]", "")) %>%
  rename_with(~ str_squish(str_replace_all(., " ", "_"))) %>%
  rename_with(~ paste0(., "_Recall"), -1)

# Fix column names in recall_nut_avg
recall_nut_avg <- recall_nut_avg %>%
  rename(
    sodium_Recall = sodiummg_Recall,
    calcium_Recall = calcium_mg_Recall,
    magnesium_Recall = magnesium_mg_Recall,
    potassium_Recall = potassium_mg_Recall,
    phosphorus_Recall = phosphorus_mg_Recall,
    carotene_mcg_Recall = `carotene_mcg_Recall`,
    iron_mg_Recall = iron_mg_Recall,  # already matches, included for clarity
    selenium_Recall = selenium_mcg_Recall,
    iodine_mcg_Recall = iodine_mcg_Recall,
    folate_Recall = folate_mcg_Recall,
    B12_Recall = B12_mcg_Recall,
    cinza_Recall = cinza_g_Recall
  )

# Fix column names in recall_poly_avg
recall_poly_avg <- recall_poly_avg %>%
  rename(
    Proanthocyanidins_Recall = `Proanthocyanidins,_total_Recall`,
    Naphtoquinones_Recall = Naphtoquinones_Recall,  # no renaming needed
    Lignans_Recall = Lignans_Recall,                # no renaming needed
    Total_Other_PPs_Recall = Total_other_PP_Recall
  )

# 5. Merge all
final_data <- ffq_final %>%
  inner_join(recall_nut_avg, by = "Participant_ID") %>%
  inner_join(recall_poly_avg, by = "Participant_ID")

colnames(final_data)

# Reorder columns in final_data for SPSS organization

# Define nutrient and polyphenol names (without suffixes)
nutrients <- c(
  "energy_kcal", "energy_kJ", "fat_g", "saturated_fat_g", "monounsatured_fat_g", "polyunsaturated_fat_g",
  "trans_fat_g", "linoleic_acid_g", "carbohydrate_g", "sugar_g", "starch_g", "oligosacharides_g",
  "protein_g", "fibre_g", "salt_g", "alcohol_g", "water_g", "cholesterol_mg", "organic_acids_g",
  "vit_a_mcg", "carotene_mcg", "vitamin_d_mcg", "vitamin_C_mg", "alpha_tocoferol_mg", "thiaminemg",
  "riboflavin_mg", "niacin_mg", "niacin_equivalent__mg", "tryptofan_mg", "B6_mg", "B12", "folate",
  "cinza", "sodium", "potassium", "calcium", "phosphorus", "magnesium", "iron_mg", "zinc", "selenium", "iodine_mcg"
)

polyphenols <- c(
  "Flavonols", "Isoflavonoids", "Anthocyanins", "Chalcones", "Dihydrochalcones", "Dihydroflavonols", "Catechins",
  "Theaflavins", "Proanthocyanidins", "Total_Flavanols", "Flavanones", "Flavones", "Total_Flavonoids",
  "Alkylphenols", "Alkylmethoxyphenols", "Furanocoumarins", "Hydroxyphenylpropenes", "Naphtoquinones",
  "Other_polyphenols", "Phenolic_terpenes", "Hydroxybenzaldehydes", "Hydroxycoumarins", "Tyrosols",
  "Total_Other_PPs", "Hydroxybenzoketones", "Hydroxyphenylacetic_acids", "Hydroxyphenylpropanoic_acids",
  "Hydroxycinnamic_acids", "Hydroxybenzoic_acids", "Total_Phenolic_Acids", "Lignans", "Stilbenes", "Total_Polyphenols"
)


colnames(ffq_poly)
colnames(recall_nut_avg)
colnames(recall_poly_avg)

unique(ffq_long$Nutr_Name)


# Interleave FFQ and Recall columns for each nutrient
nutrient_pairs <- unlist(lapply(nutrients, function(n) c(paste0(n, "_FFQ"), paste0(n, "_Recall"))))
polyphenol_pairs <- unlist(lapply(polyphenols, function(p) c(paste0(p, "_FFQ"), paste0(p, "_Recall"))))

# Define final column order
ordered_columns <- c("Participant_ID", nutrient_pairs, polyphenol_pairs)


# Reorder if columns exist
final_data <- final_data[, intersect(ordered_columns, colnames(final_data))]

# Apply log(x + 1) to all numeric columns
numeric_vars <- sapply(final_data, is.numeric)

for (var in names(final_data)[numeric_vars]) {
  new_var <- paste0("ln_", var)
  final_data[[new_var]] <- log(final_data[[var]] + 1)
}

# Save reordered file
write_csv(final_data, "merged_dataset_for_validation_ordered.csv")

# Create a copy of final_data to apply outlier removal
final_data_clean <- final_data

# Identify all log-transformed variables
ln_vars <- grep("^ln_", names(final_data), value = TRUE)

# Loop over ln_ variables
for (var in ln_vars) {
  # Calculate z-scores (excluding NA)
  z_scores <- scale(final_data[[var]])
  
  # Identify outliers: z > 3.5 or z < -3.5
  outlier_mask <- abs(z_scores) > 3.5
  
  # Count how many outliers
  outlier_count <- sum(outlier_mask, na.rm = TRUE)
  cat(var, "- Outliers > |3.5|:", outlier_count, "\n")
  
  # Replace outliers by NA in the clean dataset
  final_data_clean[[var]][outlier_mask] <- NA
}

# Check replacement was successful
for (var in ln_vars) {
  z_scores <- scale(final_data[[var]])
  outlier_mask <- abs(z_scores) > 3.5
  
  # Should return TRUE if all outliers were replaced with NA
  success <- all(is.na(final_data_clean[[var]][outlier_mask]))
  
  if (!success) {
    warning(paste("Replacement failed for", var))
  } else {
    cat("Replacement OK for", var, "\n")
  }
}

write_csv(final_data_clean, "merged_dataset_for_validation_ordered_nooutliers.csv")

colnames(final_data_clean)

cat(paste0(
  "NONPAR CORR\n  /VARIABLES=ln_", nutrients, "_FFQ ln_", nutrients, "_Recall\n",
  "  /PRINT=SPEARMAN ONETAIL NOSIG\n  /MISSING=PAIRWISE.\n\n"
))



# Get all column names
all_vars <- colnames(final_data_clean)

# Function to get all valid FFQ–Recall pairs
generate_spss_np_tests <- function(prefix = "") {
  ffq_vars <- grep(paste0("^", prefix, ".*_FFQ$"), all_vars, value = TRUE)
  syntax_blocks <- c()
  
  for (ffq_var in ffq_vars) {
    recall_var <- sub("_FFQ$", "_Recall", ffq_var)
    if (recall_var %in% all_vars) {
      block <- paste0(
        "*Nonparametric Tests: ", ffq_var, " vs ", recall_var, ".\n",
        "NPTESTS\n",
        "  /RELATED TEST(", ffq_var, " ", recall_var, ")\n",
        "  /MISSING SCOPE=ANALYSIS USERMISSING=EXCLUDE\n",
        "  /CRITERIA ALPHA=0.05  CILEVEL=95.\n"
      )
      syntax_blocks <- c(syntax_blocks, block)
    }
  }
  
  return(syntax_blocks)
}

# Generate for both raw and ln_ variables
ln_blocks <- generate_spss_np_tests("ln_")

# Combine all
final_syntax <- c(ln_blocks)

# Print to console
cat(paste(final_syntax, collapse = "\n"))



all_vars <- c(nutrients, polyphenols)


# Generate SPSS syntax
for (var in all_vars) {
  ffq <- paste0(var, "_FFQ")
  recall <- paste0(var, "_Recall")
  pct_change <- paste0(var, "_pct_diff")
  agreement <- paste0(var, "_10pct_agreement")
  
  cat(sprintf("COMPUTE %s = ((%s - %s) / %s) * 100.\n", pct_change, recall, ffq, ffq))
  cat(sprintf("VARIABLE LABELS %s '%% Difference: Recall vs FFQ'.\n", pct_change))
  
  cat(sprintf("COMPUTE %s = (ABS(%s) <= 10).\n", agreement, pct_change))
  cat(sprintf("VARIABLE LABELS %s 'Recall within ±10%% of FFQ (1=Yes, 0=No)'.\n", agreement))
  cat("EXECUTE.\n\n")
}


library(tidyverse)

# Identify all matched ln_ variable pairs
ln_ffq_vars <- names(final_data) %>% str_subset("^ln_.*_FFQ$")
ln_recall_vars <- names(final_data) %>% str_subset("^ln_.*_Recall$")
ln_common <- intersect(str_remove(ln_ffq_vars, "_FFQ$"), str_remove(ln_recall_vars, "_Recall$"))

# Preallocate results
results <- vector("list", length(ln_common))

for (i in seq_along(ln_common)) {
  var <- ln_common[i]
  ffq_col <- paste0(var, "_FFQ")
  recall_col <- paste0(var, "_Recall")
  
  # Drop missing
  df <- final_data[, c(ffq_col, recall_col)] %>% drop_na()
  if (nrow(df) < 3) next
  
  # Run Wilcoxon signed-rank test (paired = TRUE)
  wtest <- suppressWarnings(wilcox.test(df[[ffq_col]], df[[recall_col]], paired = TRUE, exact = FALSE))
  
  # Spearman correlation
  ctest <- suppressWarnings(cor.test(df[[ffq_col]], df[[recall_col]], method = "spearman", exact = FALSE))
  
  results[[i]] <- tibble(
    Variable = var,
    W_statistic = unname(wtest$statistic),
    Wilcoxon_p = wtest$p.value,
    Spearman_rho = unname(ctest$estimate),
    Spearman_p = ctest$p.value,
    Wilcoxon_signif = if_else(wtest$p.value < 0.05, "Yes", "No"),
    Spearman_signif = if_else(ctest$p.value < 0.05, "Yes", "No")
  )
}

# Combine results
final_results <- bind_rows(results)
print(final_results)


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
  "Results" = final_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Table.xlsx")

library(tidyverse)

# Identify all matched log-transformed variable pairs
ln_ffq_vars <- names(final_data) %>% str_subset("^ln_.*_FFQ$")
ln_recall_vars <- names(final_data) %>% str_subset("^ln_.*_Recall$")
ln_common <- intersect(str_remove(ln_ffq_vars, "_FFQ$"), str_remove(ln_recall_vars, "_Recall$"))

# Create a folder to store plots if desired
dir.create("bland_altman_plots", showWarnings = FALSE)

# Loop over pairs to generate plots
for (var in ln_common) {
  ffq_col <- paste0(var, "_FFQ")
  recall_col <- paste0(var, "_Recall")
  
  df <- final_data[, c(ffq_col, recall_col)] %>%
    drop_na() %>%
    mutate(
      mean_value = rowMeans(across(everything())),
      diff_value = .data[[ffq_col]] - .data[[recall_col]],
      upper_limit = mean(diff_value) + 1.96 * sd(diff_value),
      lower_limit = mean(diff_value) - 1.96 * sd(diff_value),
      mean_diff = mean(diff_value)
    )
  
  if (nrow(df) < 3) next
  
  p <- ggplot(df, aes(x = mean_value, y = diff_value)) +
    geom_point(alpha = 0.5) +
    geom_hline(aes(yintercept = mean_diff), color = "blue", linetype = "solid") +
    geom_hline(aes(yintercept = upper_limit), color = "red", linetype = "dashed") +
    geom_hline(aes(yintercept = lower_limit), color = "red", linetype = "dashed") +
    labs(
      title = paste("Bland–Altman Plot:", var),
      x = "Mean of FFQ and Recall (log-transformed)",
      y = "Difference (FFQ - Recall)"
    ) +
    theme_minimal()
  
  # Save each plot as PNG
  ggsave(filename = paste0("bland_altman_plots/BA_", var, ".png"), plot = p, width = 7, height = 5)
}
