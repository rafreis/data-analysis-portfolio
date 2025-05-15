# Load libraries
library(tidyverse)

# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/brina")
df <- read_csv("data_downloads_T36FOKHVGW_2025-02-20Assessing_Understanding_and_Regulating_Affect_AURA.csv")

# Ensure column names are clean
names(df) <- make.names(names(df))

# Visualize basic structure
glimpse(df)

# Define emotion variables by valence
positive_emotions <- c("ENTHU_ES", "EXCIT_ES", "RLX_ES", "PEACE_ES")
negative_emotions <- c("ANX_ES", "DULL_ES", "IRR_ES", "SAD_ES")

# Define prediction and retrospective variables
positive_pred <- c("PRED_ENTHU_ES", "PRED_EXCIT_ES", "PRED_RLX_ES", "PRED_PEACE_ES")
positive_retro <- c("RETRO_ENTHU_ES", "RETRO_EXCIT_ES", "RETRO_RLX_ES", "RETRO_PEACE_ES")

negative_pred <- c("PRED_ANX_ES", "PRED_DULL_ES", "PRED_IRR_ES", "PRED_SAD_ES")
negative_retro <- c("RETRO_ANX_ES", "RETRO_DULL_ES", "RETRO_IRR_ES", "RETRO_SAD_ES")

df_filtered <- df %>%
  group_by(UUID) %>%
  filter(n() >= 4) %>%
  ungroup()

# Remove rows with NA in any emotion block
df_pos_clean <- df %>% filter(if_all(all_of(positive_emotions), ~ !is.na(.)))
df_neg_clean <- df %>% filter(if_all(all_of(negative_emotions), ~ !is.na(.)))


calculate_ed <- function(dat, emotions, ..., allow_neg_icc = FALSE, fisher_transform_icc = TRUE) {
  
  center <- function(x){
    x - mean(x, na.rm = TRUE)
  }
  
  dat$row_id <- 1:nrow(dat)
  
  c_dat <- dat %>% group_by(...) %>% mutate(across(all_of(emotions), center))
  
  # Calculate person-level metrics
  
  person_dat <- c_dat %>% group_by(...) %>% group_split( .keep = TRUE) %>% lapply(., function(x)
    # Keep the ID vars
    cbind(x %>% select(row_id, ...),
          
          # Calculate emotional variance
          m_emo_var = sum(diag(var(x[, emotions]))),
          # Calculate classic (non) ED
          c_nonED = psych::ICC(x[emotions], missing = TRUE, lmer = F)$results[6, 2]
          
    )) %>%
    do.call('rbind', .)
  
  c_dat <- left_join(c_dat, person_dat)
  
  c_dat$momentary_squared_sum <- c_dat[emotions] %>% apply(., 1, function(x) (mean(x, na.rm = TRUE)*length(emotions))^2)
  
  c_dat$m_nonED <- c_dat$momentary_squared_sum/c_dat$m_emo_var
  c_dat$m_ED <- c_dat$momentary_squared_sum/c_dat$m_emo_var*-1
  
  
  if (fisher_transform_icc) {
    # We can safely transform before we flip to calculate c_ED or doing filtering, because fisher transform is
    # symmetrical around 0 and retains 0s.
    c_dat$c_nonED <- psych::fisherz(c_dat$c_nonED)
  }
  
  c_dat$c_ED <- c_dat$c_nonED*-1
  
  c_dat <- c_dat %>% group_by(...) %>% mutate(L2_nonED = sum(m_ED, na.rm = TRUE)/(n()-1))
  
  c_dat$L2_ED <- c_dat$L2_nonED*-1
  
  if (!allow_neg_icc) {
    # Set all measures of emotion differentiation to NA if the c_ED is negative
    c_dat <- c_dat %>% mutate(across(c('m_nonED', 'm_ED', 'c_nonED', 'c_ED', 'L2_nonED', 'L2_ED'),
                                     function(x) ifelse(c_nonED < 0, NA, x)))
  }
  
  
  out <- left_join(dat, select(c_dat, row_id, m_nonED, m_ED, c_nonED, c_ED, L2_nonED, L2_ED)) %>% select(-row_id)
  
  return(out)
}

library(psych)

# Step 1 – Emotional Granularity (ICC2 and Fisher's Z)
df_with_ed_pos <- calculate_ed(df_pos_clean, emotions = positive_emotions, UUID)
df_with_ed_neg <- calculate_ed(df_neg_clean, emotions = negative_emotions, UUID)

granularity_pos <- df_with_ed_pos %>% distinct(UUID, c_ED) %>% rename(granularity_pos = c_ED)
granularity_neg <- df_with_ed_neg %>% distinct(UUID, c_ED) %>% rename(granularity_neg = c_ED)


# Affective Forecasting Accuracy

df_eod <- df_filtered %>%
  group_by(UUID, DayNumber_DAY_ES) %>%
  filter(!all(is.na(Time_Local))) %>%  # drop groups where all times are NA
  filter(Time_Local == max(Time_Local, na.rm = TRUE)) %>%
  ungroup()

# Step 2: Create Day t+1 retrospective scores using only end-of-day surveys
df_shifted <- df_eod %>%
  mutate(DayNumber_DAY_ES = DayNumber_DAY_ES - 1) %>%
  select(UUID, DayNumber_DAY_ES, all_of(c(positive_retro, negative_retro))) %>%
  rename_with(~ paste0("SHIFTED_", .), all_of(c(positive_retro, negative_retro)))

# Step 3: Join Day t predictions with Day t+1 end-of-day retrospectives
df_matched <- df_filtered %>%
  inner_join(df_shifted, by = c("UUID", "DayNumber_DAY_ES"))

# Step 4: Compute absolute error per emotion
compute_abs_errors <- function(pred_vars, shifted_retro_vars, df) {
  map2_dfc(pred_vars, shifted_retro_vars, ~ abs(df[[.x]] - df[[.y]])) %>%
    set_names(~ paste0("err_", pred_vars))
}


df_errors <- df_matched %>%
  bind_cols(
    compute_abs_errors(positive_pred, paste0("SHIFTED_", positive_retro), df_matched),
    compute_abs_errors(negative_pred, paste0("SHIFTED_", negative_retro), df_matched)
  )

glimpse(df_errors)

# Step 5: Compute composite scores
df_accuracy <- df_errors %>%
  mutate(
    pos_accuracy = rowMeans(select(., starts_with("err_PRED_ENTHU_ES"), starts_with("err_PRED_EXCIT_ES"),
                                   starts_with("err_PRED_RLX_ES"), starts_with("err_PRED_PEACE_ES")), na.rm = TRUE),
    neg_accuracy = rowMeans(select(., starts_with("err_PRED_ANX_ES"), starts_with("err_PRED_DULL_ES"),
                                   starts_with("err_PRED_IRR_ES"), starts_with("err_PRED_SAD_ES")), na.rm = TRUE)
  ) %>%
  select(UUID, DayNumber_DAY_ES, pos_accuracy, neg_accuracy)

df_accuracy_summary <- df_accuracy %>%
  group_by(UUID) %>%
  summarise(
    pos_accuracy_mean = mean(pos_accuracy, na.rm = TRUE),
    pos_accuracy_sd   = sd(pos_accuracy, na.rm = TRUE),
    pos_accuracy_sem  = sd(pos_accuracy, na.rm = TRUE) / sqrt(sum(!is.na(pos_accuracy))),
    neg_accuracy_mean = mean(neg_accuracy, na.rm = TRUE),
    neg_accuracy_sd   = sd(neg_accuracy, na.rm = TRUE),
    neg_accuracy_sem  = sd(neg_accuracy, na.rm = TRUE) / sqrt(sum(!is.na(neg_accuracy)))
  )


df_accuracy_summary_reverse <- df_accuracy_summary %>%
  mutate(
    pos_accuracy_rev = 100 - pos_accuracy_mean,
    neg_accuracy_rev = 100 - neg_accuracy_mean
  )

# List of error variables
pos_error_vars <- c("err_PRED_ENTHU_ES", "err_PRED_EXCIT_ES", "err_PRED_RLX_ES", "err_PRED_PEACE_ES")
neg_error_vars <- c("err_PRED_ANX_ES", "err_PRED_DULL_ES", "err_PRED_IRR_ES", "err_PRED_SAD_ES")

# Recalculate composite accuracy from df_errors
df_accuracy_summary_filt <- df_errors %>%
  group_by(UUID) %>%
  summarise(
    n_days = n(),
    pos_accuracy_mean = rowMeans(across(all_of(pos_error_vars)), na.rm = TRUE) %>% mean(na.rm = TRUE),
    pos_accuracy_sd   = rowMeans(across(all_of(pos_error_vars)), na.rm = TRUE) %>% sd(na.rm = TRUE),
    pos_accuracy_sem  = pos_accuracy_sd / sqrt(n_days),
    neg_accuracy_mean = rowMeans(across(all_of(neg_error_vars)), na.rm = TRUE) %>% mean(na.rm = TRUE),
    neg_accuracy_sd   = rowMeans(across(all_of(neg_error_vars)), na.rm = TRUE) %>% sd(na.rm = TRUE),
    neg_accuracy_sem  = neg_accuracy_sd / sqrt(n_days)
  ) %>%
  filter(n_days >= 3)


df_accuracy_summary_filt_reverse <- df_accuracy_summary_filt %>%
  mutate(
    pos_accuracy_rev = 100 - pos_accuracy_mean,
    neg_accuracy_rev = 100 - neg_accuracy_mean
  )

# Merge everything into a final dataframe

df_final <- df_filtered %>%
  select(UUID) %>%
  distinct() %>%
  left_join(granularity_pos, by = "UUID") %>%
  left_join(granularity_neg, by = "UUID") %>%
  left_join(df_accuracy_summary_filt_reverse, by = "UUID") 


# Identify all baseline variables
baseline_vars <- names(df) %>% str_subset("_BL$")

# Count how many unique values each UUID has across time for each BL variable
check_variability <- df %>%
  select(UUID, all_of(baseline_vars)) %>%
  group_by(UUID) %>%
  summarise(across(everything(), ~ n_distinct(.x[!is.na(.x)]))) %>%
  ungroup()

# Gather results in long format and filter for non-constant variables
nonconstant_vars <- check_variability %>%
  pivot_longer(-UUID, names_to = "variable", values_to = "n_unique") %>%
  filter(n_unique > 1)

# View
nonconstant_vars %>% distinct(variable)

# Merge Well-being variables

df_final <- df_final %>%
  left_join(
    df_matched %>%
      group_by(UUID) %>%
      slice(1) %>%
      ungroup() %>%
      mutate(
        SWL_Score = rowMeans(select(., starts_with("SWL_")), na.rm = TRUE),
        PWB_Score = rowMeans(select(., starts_with("PWB_SF_")), na.rm = TRUE),
        GAD_Score = rowMeans(select(., starts_with("GAD7_")), na.rm = TRUE),
        CESD_Score = rowMeans(select(., starts_with("CESD_")), na.rm = TRUE),
        GAD_Rev   = 3 - GAD_Score,
        CESD_Rev  = 3 - CESD_Score
      ),
    by = "UUID"
  )


library(mediation)


# Adapted mediation function for AURA-style analyses
run_mediation_aura <- function(df, iv, mediator, dv, sims = 5000, diagnostics = FALSE, save_plots = FALSE, plot_prefix = NULL) {
  df$X_iv <- df[[iv]]
  df$X_med <- df[[mediator]]
  df$X_dv <- df[[dv]]
  
  model_m <- lm(X_med ~ X_iv, data = df)
  model_y <- lm(X_dv ~ X_iv + X_med, data = df)
  
  med_out <- mediate(
    model.m = model_m,
    model.y = model_y,
    treat = "X_iv",
    mediator = "X_med",
    boot = TRUE,
    sims = sims
  )
  
  results_df <- data.frame(
    Metric = c("ACME (mediation)", "ADE (direct)", "Total Effect", "Prop. Mediated"),
    Estimate = c(med_out$d0, med_out$z0, med_out$tau.coef, med_out$n0),
    `95% CI Lower` = c(med_out$d0.ci[1], med_out$z0.ci[1], med_out$tau.ci[1], med_out$n0.ci[1]),
    `95% CI Upper` = c(med_out$d0.ci[2], med_out$z0.ci[2], med_out$tau.ci[2], med_out$n0.ci[2]),
    `p-value` = c(med_out$d0.p, med_out$z0.p, med_out$tau.p, med_out$n0.p)
  )
  
  if (diagnostics || save_plots) {
    # Plot for mediator model
    if (save_plots) png(paste0(plot_prefix, "_mediator_model.png"), width = 800, height = 800)
    par(mfrow = c(2, 2))
    plot(model_m, main = "Mediator model")
    if (save_plots) dev.off()
    
    # Plot for outcome model
    if (save_plots) png(paste0(plot_prefix, "_outcome_model.png"), width = 800, height = 800)
    par(mfrow = c(2, 2))
    plot(model_y, main = "Outcome model")
    if (save_plots) dev.off()
    
    par(mfrow = c(1, 1))  # Reset layout
  }
  
  return(results_df)
}

# Positive valence mediation models
df_results_pos_swl  <- run_mediation_aura(df_final, "granularity_pos", "pos_accuracy_rev", "SWL_Score",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_pos_SWL")

df_results_pos_pwb  <- run_mediation_aura(df_final, "granularity_pos", "pos_accuracy_rev", "PWB_Score",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_pos_PWB")

df_results_pos_gad  <- run_mediation_aura(df_final, "granularity_pos", "pos_accuracy_rev", "GAD_Rev",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_pos_GAD")

df_results_pos_cesd <- run_mediation_aura(df_final, "granularity_pos", "pos_accuracy_rev", "CESD_Rev",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_pos_CESD")

# Negative valence mediation models
df_results_neg_swl  <- run_mediation_aura(df_final, "granularity_neg", "neg_accuracy_rev", "SWL_Score",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_neg_SWL")

df_results_neg_pwb  <- run_mediation_aura(df_final, "granularity_neg", "neg_accuracy_rev", "PWB_Score",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_neg_PWB")

df_results_neg_gad  <- run_mediation_aura(df_final, "granularity_neg", "neg_accuracy_rev", "GAD_Rev",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_neg_GAD")

df_results_neg_cesd <- run_mediation_aura(df_final, "granularity_neg", "neg_accuracy_rev", "CESD_Rev",
                                          sims = 5000, diagnostics = TRUE, save_plots = TRUE, plot_prefix = "mediation_neg_CESD")


library(dplyr)

df_mediation_all <- bind_rows(
  "pos_SWL"  = df_results_pos_swl,
  "pos_PWB"  = df_results_pos_pwb,
  "pos_GAD"  = df_results_pos_gad,
  "pos_CESD" = df_results_pos_cesd,
  "neg_SWL"  = df_results_neg_swl,
  "neg_PWB"  = df_results_neg_pwb,
  "neg_GAD"  = df_results_neg_gad,
  "neg_CESD" = df_results_neg_cesd,
  .id = "Model"
)


# Extra Visualizations

# Load ggplot2 if not already loaded
library(ggplot2)
library(dplyr)
library(tidyr)

# 1. Scatterplot: Granularity (positive) vs Forecasting Accuracy (positive)
# Title: Positive Granularity vs Forecasting Accuracy
p1 <- ggplot(df_final, aes(x = granularity_pos, y = pos_accuracy_rev)) +
  geom_point(alpha = 0.5) +
  geom_smooth(method = "lm", se = TRUE, color = "blue") +
  labs(title = "Positive Granularity vs Forecasting Accuracy",
       x = "Positive Emotional Granularity (Fisher-Z ICC × -1)",
       y = "Positive AF Accuracy (Reversed Error)") +
  theme_minimal()
print(p1)
ggsave("granularity_pos_vs_accuracy.png", plot = p1, width = 8, height = 6, dpi = 300)

# 2. Scatterplot: Granularity (negative) vs Forecasting Accuracy (negative)
# Title: Negative Granularity vs Forecasting Accuracy
p2 <- ggplot(df_final, aes(x = granularity_neg, y = neg_accuracy_rev)) +
  geom_point(alpha = 0.5) +
  geom_smooth(method = "lm", se = TRUE, color = "red") +
  labs(title = "Negative Granularity vs Forecasting Accuracy",
       x = "Negative Emotional Granularity (Fisher-Z ICC × -1)",
       y = "Negative AF Accuracy (Reversed Error)") +
  theme_minimal()
print(p2)
ggsave("granularity_neg_vs_accuracy.png", plot = p2, width = 8, height = 6, dpi = 300)

# 3. Faceted scatterplots: Forecasting Accuracy vs Well-Being Outcomes
# Title: Positive Forecasting Accuracy vs Well-Being Outcomes
df_long <- df_final %>%
  dplyr::select(UUID, pos_accuracy_rev, SWL_Score, PWB_Score, GAD_Rev, CESD_Rev) %>%
  pivot_longer(cols = c(SWL_Score, PWB_Score, GAD_Rev, CESD_Rev),
               names_to = "Outcome", values_to = "Score")


p3 <- ggplot(df_long, aes(x = pos_accuracy_rev, y = Score)) +
  geom_point(alpha = 0.4) +
  geom_smooth(method = "lm", se = FALSE, color = "darkgreen") +
  facet_wrap(~ Outcome, scales = "free_y") +
  labs(title = "Positive Forecasting Accuracy vs Well-Being Outcomes",
       x = "Positive AF Accuracy (Reversed Error)", y = "Outcome Score") +
  theme_minimal()
print(p3)
ggsave("accuracy_vs_wellbeing_facets.png", plot = p3, width = 10, height = 6, dpi = 300)

colnames(df_mediation_all)
# 4. Forest plot of effects (ACME, ADE, Total)
# Title: Effect Estimates from Mediation Models
df_plot <- df_mediation_all %>%
  rename(`95% CI Lower` = X95..CI.Lower,
         `95% CI Upper` = X95..CI.Upper) %>%
  filter(Metric %in% c("ACME (mediation)", "ADE (direct)", "Total Effect")) %>%
  mutate(Metric = factor(Metric, levels = c("ACME (mediation)", "ADE (direct)", "Total Effect")))


p4 <- ggplot(df_plot, aes(x = Estimate, y = Model, color = Metric)) +
  geom_point(position = position_dodge(width = 0.5)) +
  geom_errorbar(aes(xmin = `95% CI Lower`, xmax = `95% CI Upper`), width = 0.2,
                position = position_dodge(width = 0.5)) +
  facet_wrap(~ Metric, scales = "free_x") +
  labs(title = "Effect Estimates from Mediation Models", x = "Effect Size", y = "") +
  theme_minimal()

print(p4)
ggsave("mediation_effects_forest.png", plot = p4, width = 10, height = 6, dpi = 300)


# 5. Barplot of Proportion Mediated with value labels
p5 <- df_mediation_all %>%
  filter(Metric == "Prop. Mediated") %>%
  ggplot(aes(x = reorder(Model, Estimate), y = Estimate)) +
  geom_col(fill = "steelblue") +
  geom_text(aes(label = round(Estimate, 2)), hjust = -0.1, size = 3.5) +
  coord_flip() +
  labs(title = "Proportion of Effect Mediated by Forecasting Accuracy",
       x = "Model", y = "Proportion Mediated") +
  theme_minimal() +
  ylim(0, max(df_mediation_all$Estimate[df_mediation_all$Metric == "Prop. Mediated"], na.rm = TRUE) + 0.05)

print(p5)
ggsave("proportion_mediated_labeled.png", plot = p5, width = 8, height = 6, dpi = 300)

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
  "Final EOD Data" = df_final,
  "Accuracy Summary" = df_accuracy_summary_filt_reverse,
  "Model Results" = df_plot
)

# Save to Excel with APA formatting
#save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")

library(ggplot2)
library(tidyr)
library(dplyr)

# Prepare data in long format
df_wellbeing_long <- df_final %>%
  dplyr::select(UUID, SWL_Score, PWB_Score, GAD_Rev, CESD_Rev) %>%
  pivot_longer(cols = -UUID, names_to = "Wellbeing_Variable", values_to = "Score")

# Boxplot
p_box <- ggplot(df_wellbeing_long, aes(x = Wellbeing_Variable, y = Score, fill = Wellbeing_Variable)) +
  geom_boxplot(alpha = 0.7, outlier.color = "black") +
  labs(title = "Distribution of Well-Being Variables",
       x = "Well-Being Measure",
       y = "Score") +
  theme_minimal() +
  theme(legend.position = "none")

print(p_box)
ggsave("wellbeing_boxplots.png", plot = p_box, width = 8, height = 6, dpi = 300)








# New Requests

# Load libraries
library(tidyverse)
library(psych)
library(corrr)
library(broom)

class(df_final)

df_final <- as.data.frame(df_final)

# Prepare base dataset
df_summary <- df_final %>%
  dplyr::select(UUID, granularity_pos, granularity_neg,
                pos_accuracy_rev, neg_accuracy_rev,
                SWL_Score, PWB_Score, GAD_Rev, CESD_Rev)


# 1. Descriptives (Sample size, Age, Gender)
desc_table <- df %>%
  filter(UUID %in% df_summary$UUID) %>%
  distinct(UUID, AGE_BL, GENDER_BL) %>%
  summarise(
    N_final = n(),
    M_age = mean(AGE_BL, na.rm = TRUE),
    SD_age = sd(AGE_BL, na.rm = TRUE),
    Female = sum(GENDER_BL == 1, na.rm = TRUE),
    Male = sum(GENDER_BL == 2, na.rm = TRUE),
    Other = sum(!GENDER_BL %in% c(1, 2), na.rm = TRUE)
  )

# 2. Correlation matrix
cor_matrix <- df_summary %>%
  dplyr::select(-UUID) %>%
  correlate() %>%
  stretch(na.rm = TRUE) %>%
  mutate(
    r = round(r, 3),
    p = map2_dbl(x, y, ~ cor.test(df_summary[[.x]], df_summary[[.y]])$p.value),
    p = round(p, 3)
  )

# Filter only relevant H1 and H2 pairs
vars_granularity <- c("granularity_pos", "granularity_neg")
vars_accuracy <- c("pos_accuracy_rev", "neg_accuracy_rev")
vars_wellbeing <- c("SWL_Score", "PWB_Score", "GAD_Rev", "CESD_Rev")

h1_pairs <- expand.grid(vars_granularity, vars_accuracy) %>% as_tibble()
h2_pairs <- expand.grid(c(vars_granularity, vars_accuracy), vars_wellbeing) %>% as_tibble()

# Extract H1
cor_h1 <- cor_matrix %>%
  filter((x %in% h1_pairs$Var1 & y %in% h1_pairs$Var2) |
           (y %in% h1_pairs$Var1 & x %in% h1_pairs$Var2))

# Extract H2
cor_h2 <- cor_matrix %>%
  filter((x %in% h2_pairs$Var1 & y %in% h2_pairs$Var2) |
           (y %in% h2_pairs$Var1 & x %in% h2_pairs$Var2))

# Save results for export
list_results <- list(
  "Descriptives_Age_Gender" = desc_table,
  "Correlations_H1" = cor_h1,
  "Correlations_H2" = cor_h2
)

print(cor_h1)
print(cor_h2, n=50)

# View in Excel-friendly dataframes
library(openxlsx)
write.xlsx(list_results, "summary_tables_brina.xlsx")
