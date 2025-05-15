# Update the scale_name_mapping vector
scale_name_mapping <- c(
  "six_min_walk_test_distance_in_meter" = "6-Minuten gehtest (MGT)",
  "development_1" = "Patienten Souverantitaet 19 Items (PS19)",
  "development_2" = "Patienten Souverantitaet 8 Dimensionen (PS8D)",
  "development_3" = "Patienten Souverantitaet Sicht Pflegender 19 Items (PS19-P)",
  "development_4" = "Patienten Souverantitaet Sicht Pflegender 8 Dimensionen (PS8D-P)",
  "development_5" = "Effekt Pflegender (ICG)",
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe",
  "Blood_pressure___at_rest__systolic" = "BD systolisch",
  "Blood_pressure___at_rest__diastolic_" = "BD diastolisch"
)

# For Control and Treatment groups
group_name_mapping <- c(
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe"
)


setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/chippie77/NewPlots")
# Function for Boxplot Comparison
create_boxplot_comparison <- function(data, variables, factor_column1, factor_column2, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  library(stringr) # For string wrapping
  
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(
      Variable = ifelse(Variable %in% names(scale_name_mapping), scale_name_mapping[Variable], Variable),
      !!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1),
      !!sym(factor_column2) := factor(ifelse(!!sym(factor_column2) %in% names(group_name_mapping), group_name_mapping[!!sym(factor_column2)], !!sym(factor_column2)))
    )
  
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Value, fill = !!sym(factor_column2))) +
    geom_boxplot() +
    facet_wrap(~ Variable, scales = "free_y", labeller = labeller(Variable = label_wrap_gen(25))) +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Boxplot Vergleich", 
      x = "Endpunkte",  
      y = "Wert",
      fill = "Gruppe"
    ) +
    scale_fill_discrete(name = "Gruppe")
  
  return(p)
}

# For boxplots
scales_diga <- c("six_min_walk_test_distance_in_meter", "development_1", "development_2", "development_3", "development_4", "development_5", "Blood_pressure___at_rest__systolic", "Blood_pressure___at_rest__diastolic_")

scales_dipa <- scales_diga

plot <- create_boxplot_comparison(df_diga, scales_diga, "Pre_Post", "Classification", c("Pre", "Post"))
print(plot)
ggsave("boxplot_comparison_diga.png", plot = plot, width = 12, height = 8)

plot <- create_boxplot_comparison(df_dipa, scales_dipa, "Pre_Post", "Classification", c("Pre", "Post"))
print(plot)
ggsave("boxplot_comparison_dipa.png", plot = plot, width = 12, height = 8)

# Function for Mean and Confidence Interval Plot
create_mean_ci_plot <- function(data, variables, factor_column1, factor_column2, levels_factor1, legend_title) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  library(stringr) # For string wrapping
  
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(
      Variable = ifelse(Variable %in% names(scale_name_mapping), scale_name_mapping[Variable], Variable),
      !!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1),
      !!sym(factor_column2) := factor(ifelse(!!sym(factor_column2) %in% names(group_name_mapping), group_name_mapping[!!sym(factor_column2)], !!sym(factor_column2)))
    ) %>%
    group_by(Variable, !!sym(factor_column1), !!sym(factor_column2)) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      N = n(),
      SD = sd(Value, na.rm = TRUE),
      SEM = SD / sqrt(N),
      .groups = 'drop'
    ) %>%
    mutate(
      CI = qt(0.975, df = N-1) * SEM,
      Lower = Mean - CI,
      Upper = Mean + CI
    )
  
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Mean, color = !!sym(factor_column2))) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = Lower, ymax = Upper), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y", labeller = labeller(Variable = label_wrap_gen(25))) +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mittelwert und 95%-Konfidenzintervall Plot", 
      x = "Endpunkte",  
      y = "Wert",
      color = legend_title
    ) +
    scale_color_discrete(name = legend_title)
  
  return(p)
}

plot <- create_mean_ci_plot(df_diga, scales_diga, "Pre_Post", "Classification", c("Pre", "Post"), "ADELE Nutzung")
print(plot)
ggsave("mean_ci_plot_diga.png", plot = plot, width = 12, height = 8)

plot <- create_mean_ci_plot(df_dipa, scales_dipa, "Pre_Post", "Classification", c("Pre", "Post"), "ADELE Nutzung")
print(plot)
ggsave("mean_ci_plot_dipa.png", plot = plot, width = 12, height = 8)

# Update scale mapping for differences
scale_name_mapping <- c(
  "six_min_walk_test_distance_in_meter_difference" = "Six Minuten gehtest MGT",
  "development_1_difference" = "Patienten Souveraenitaet PS19",
  "development_2_difference" = "Patienten Souveraenitaet PS8D",
  "development_3_difference" = "Patienten Souveraenitaet Sicht Pflegender PS19 P",
  "development_4_difference" = "Patienten Souveraenitaet Sicht Pflegender PS8D P",
  "development_5_difference" = "Effekt Pflegender ICG",
  "HBP_MEAN_difference" = "HBP MEAN",
  "PCS_12_difference" = "PCS 12",
  "MCS_12_difference" = "MCS 12",
  "HCS_MEAN_difference" = "HCS MEAN",
  "DIGA_HELM_Evaluation_Overall_difference" = "DIGA HELM",
  "HLS_SUM_difference" = "HLS SUM",
  "Blood.pressure._.at.rest._systolic" = "BD systolisch",
  "Blood.pressure._.at.rest._diastolic_" = "BD diastolisch",
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe"
)

# Update the Classification column in the difference dataframes
df_diga_differences <- df_diga_differences %>%
  mutate(Classification = ifelse(Classification %in% names(group_name_mapping), 
                                 group_name_mapping[Classification], 
                                 Classification))

df_dipa_differences <- df_dipa_differences %>%
  mutate(Classification = ifelse(Classification %in% names(group_name_mapping), 
                                 group_name_mapping[Classification], 
                                 Classification))

create_mean_ci_plot_no_secondary <- function(data, variables, factor_column1, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  library(stringr)  # For string wrapping
  
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    mutate(
      Variable = ifelse(Variable %in% names(scale_name_mapping), scale_name_mapping[Variable], Variable),
      !!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1)
    ) %>%
    group_by(Variable, !!sym(factor_column1)) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      N = n(),
      SE = SD / sqrt(N),
      CI_lower = Mean - 1.96 * SE,
      CI_upper = Mean + 1.96 * SE,
      .groups = 'drop'
    )
  
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Mean)) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = CI_lower, ymax = CI_upper), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y", labeller = label_parsed) +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mittelwert- und 95% Konfidenzintervall-Diagramm", 
      x = "Primärer Faktor",  
      y = "Wert"
    )
  
  return(p)
}

factor1_levels <- c("Kontrollgruppe", "Interventionsgruppe")
factor_column1 <- "Classification"

scales_diga_diff <- c("six_min_walk_test_distance_in_meter_difference", "development_1_difference", "development_2_difference", "development_3_difference", "development_4_difference", "development_5_difference", "HBP_MEAN_difference", "PCS_12_difference", "MCS_12_difference", "HCS_MEAN_difference", "DIGA_HELM_Evaluation_Overall_difference", "HLS_SUM_difference", "Blood.pressure._.at.rest._systolic", "Blood.pressure._.at.rest._diastolic_")
scales_dipa_diff <- scales_diga_diff

plot <- create_mean_ci_plot_no_secondary(df_diga_differences, scales_diga_diff, factor_column1, factor1_levels)
print(plot)
ggsave("mean_ci_plot_diga_diff.png", plot = plot, width = 12, height = 8)

plot <- create_mean_ci_plot_no_secondary(df_dipa_differences, scales_dipa_diff, factor_column1, factor1_levels)
print(plot)
ggsave("mean_ci_plot_dipa_diff.png", plot = plot, width = 12, height = 8)

create_boxplot_comparison_no_secondary <- function(data, variables, factor_column1, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  library(stringr)  # For string wrapping
  
  long_data <- data %>%
    pivot_longer(cols = variables, names_to = "Variable", values_to = "Value") %>%
    mutate(
      Variable = ifelse(Variable %in% names(scale_name_mapping), scale_name_mapping[Variable], Variable),
      !!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1)
    )
  
  p <- ggplot(long_data, aes(x = !!sym(factor_column1), y = Value)) +
    geom_boxplot() +
    facet_wrap(~ Variable, scales = "free_y", labeller = label_parsed) +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Boxplot-Vergleich", 
      x = "Primärer Faktor",  
      y = "Wert"
    )
  
  return(p)
}

plot <- create_boxplot_comparison_no_secondary(df_diga_differences, scales_diga_diff, factor_column1, factor1_levels)
print(plot)
ggsave("boxplot_comparison_diga_diff.png", plot = plot, width = 12, height = 8)

plot <- create_boxplot_comparison_no_secondary(df_dipa_differences, scales_dipa_diff, factor_column1, factor1_levels)
print(plot)
ggsave("boxplot_comparison_dipa_diff.png", plot = plot, width = 12, height = 8)

# Function for Generating and Saving Plots (Histograms and QQ Plots)
generate_and_save_plots <- function(data, variables, filename, plot_type) {
  library(ggplot2)
  library(gridExtra)
  library(stringr) # For string wrapping
  
  plots <- lapply(variables, function(var) {
    title <- ifelse(var %in% names(scale_name_mapping), str_wrap(scale_name_mapping[var], width = 20), var)
    
    if (any(!is.na(data[[var]]))) {
      data_filtered <- data[!is.na(data[[var]]), ]
      
      if (plot_type == "histogram") {
        ggplot(data_filtered, aes_string(x = var)) +
          geom_histogram(bins = 30, fill = "blue", color = "black") +
          theme_minimal() +
          ggtitle(title)
      } else if (plot_type == "qq") {
        ggplot(data_filtered, aes_string(sample = var)) +
          stat_qq() +
          stat_qq_line() +
          theme_minimal() +
          ggtitle(title)
      }
    } else {
      NULL
    }
  })
  
  plots <- Filter(Negate(is.null), plots)
  
  if (length(plots) > 0) {
    combined_plot <- do.call(grid.arrange, c(plots, ncol = 4, nrow = ceiling(length(plots) / 4)))
    ggsave(filename, plot = combined_plot, width = 20, height = 15)
  } else {
    message("No plots were generated due to data issues.")
  }
}

# For histograms
generate_and_save_plots(df_diga_differences, scales_diga_diff, "histograms_diga_diff.png", "histogram")
generate_and_save_plots(df_dipa_differences, scales_dipa_diff, "histograms_dipa_diff.png", "histogram")

# For QQ plots
generate_and_save_plots(df_dipa_differences, scales_dipa_diff, "qq_plots_dipa_diff.png", "qq")
generate_and_save_plots(df_diga_differences, scales_diga_diff, "qq_plots_diga_diff.png", "qq")

# For histograms and QQ plots for the main datasets
generate_and_save_plots(df_diga, scales_diga, "histograms_diga.png", "histogram")
generate_and_save_plots(df_dipa, scales_dipa, "histograms_dipa.png", "histogram")
generate_and_save_plots(df_diga, scales_diga, "qq_plots_diga.png", "qq")
generate_and_save_plots(df_dipa, scales_dipa, "qq_plots_dipa.png", "qq")
