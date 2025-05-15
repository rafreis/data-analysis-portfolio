# Create a named vector for scale mapping
scale_name_mapping <- c(
  "six_min_walk_test_distance_in_meter" = "6-Minuten gehtest (MGT)",
  "development_1" = "Patienten Souverantitaet 19 Items (PS19)",
  "development_2" = "Patienten Souverantitaet 8 Dimensionen (PS8D)",
  "development_3" = "Patienten Souverantitaet Sicht Pflegender 19 Items (PS19-P)",
  "development_4" = "Patienten Souverantitaet Sicht Pflegender 8 Dimensionen (PS8D-P)",
  "development_5" = "Effekt Pflegender (ICG)",
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe"
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

## UPDATE SCALE MAPPING FOR DIFFERENCES

# Create a named vector for scale mapping
scale_name_mapping <- c(
  "six_min_walk_test_distance_in_meter_difference" = "Six_Minuten_gehtest_MGT",
  "development_1_difference" = "Patienten_Souveraenitaet_PS19",
  "development_2_difference" = "Patienten_Souveraenitaet_PS8D",
  "development_3_difference" = "Patienten_Souveraenitaet_Sicht_Pflegender_PS19_P",
  "development_4_difference" = "Patienten_Souveraenitaet_Sicht_Pflegender_PS8D_P",
  "development_5_difference" = "Effekt_Pflegender_ICG",
  "HBP_MEAN_difference" = "HBP_MEAN",
  "PCS_12_difference" = "PCS_12",
  "MCS_12_difference" = "MCS_12",
  "HCS_MEAN_difference" = "HCS_MEAN",
  "DIGA_HELM_Evaluation_Overall_difference" = "DIGA_HELM_Eval_Overall",
  "HLS_SUM_difference" = "HLS_SUM",
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe"
)

# For Control and Treatment groups
group_name_mapping <- c(
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe"
)


df_diga_differences <- df_diga_differences %>%
  mutate(Classification = ifelse(Classification %in% names(group_name_mapping), 
                                 group_name_mapping[Classification], 
                                 Classification))

# Update the Classification column in df_dipa_differences
df_dipa_differences <- df_dipa_differences %>%
  mutate(Classification = ifelse(Classification %in% names(group_name_mapping), 
                                 group_name_mapping[Classification], 
                                 Classification))




create_mean_ci_plot_no_secondary <- function(data, variables, factor_column1, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  library(stringr)  # For string wrapping
  
  # Pivot the data to long format
  long_data <- data %>%
    pivot_longer(cols = variables, names_to = "Variable", values_to = "Value")
  
  # Apply scale name mapping
  long_data$Variable <- ifelse(long_data$Variable %in% names(scale_name_mapping), scale_name_mapping[long_data$Variable], long_data$Variable)
  
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

plot <- create_mean_ci_plot_no_secondary(df_diga_differences, scales_diga_diff, factor_column1,factor1_levels)
print(plot)
ggsave("mean_ci_plot_diga_diff.png", plot = plot, width = 12, height = 8)

plot <- create_mean_ci_plot_no_secondary(df_dipa_differences, scales_dipa_diff, factor_column1,factor1_levels)
print(plot)
ggsave("mean_ci_plot_dipa_diff.png", plot = plot, width = 12, height = 8)



create_boxplot_comparison_no_secondary <- function(data, variables, factor_column1, levels_factor1) {
  library(ggplot2)
  library(dplyr)
  library(tidyr)
  library(stringr)  # For string wrapping
  
  # Pivot the data to long format
  long_data <- data %>%
    pivot_longer(cols = variables, names_to = "Variable", values_to = "Value") %>%
    mutate(
      Variable = ifelse(Variable %in% names(scale_name_mapping), scale_name_mapping[Variable], Variable),
      !!sym(factor_column1) := factor(!!sym(factor_column1), levels = levels_factor1)
    )
  
  # Create the boxplot
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
  


# Create a named vector for scale mapping
scale_name_mapping <- c(
  "six_min_walk_test_distance_in_meter_difference" = "Six Minuten gehtest MGT",
  "development_1_difference" = "Patienten Souveraenitaet PS19",
  "development_2_difference" = "Patienten Souveraenitaet PS8D",
  "development_3_difference" = "Patienten Souveraenitaet Sicht_Pflegender PS19 P",
  "development_4_difference" = "Patienten Souveraenitaet Sicht_Pflegender PS8D P",
  "development_5_difference" = "Effekt Pflegender ICG",
  "HBP_MEAN_difference" = "HBP MEAN",
  "PCS_12_difference" = "PCS 12",
  "MCS_12_difference" = "MCS 12",
  "HCS_MEAN_difference" = "HCS MEAN",
  "DIGA_HELM_Evaluation_Overall_difference" = "DIGA HELM",
  "HLS_SUM_difference" = "HLS SUM",
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe"
)


# Function for Generating and Saving Plots (Histograms and QQ Plots)
generate_and_save_plots <- function(data, variables, filename, plot_type) {
  library(ggplot2)
  library(gridExtra)
  library(stringr) # For string wrapping
  
  plots <- lapply(variables, function(var) {
    # Apply scale name mapping if exists
    title <- ifelse(var %in% names(scale_name_mapping), str_wrap(scale_name_mapping[var], width = 20), var)
    
    # Check if the variable column has any non-NA values
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
      NULL # Return NULL if all values are NA
    }
  })
  
  # Filter out NULL elements in case some variables had only NA values
  plots <- Filter(Negate(is.null), plots)
  
  # Generate combined plot only if there are plots to combine
  if (length(plots) > 0) {
    combined_plot <- do.call(grid.arrange, c(plots, ncol = 4, nrow = ceiling(length(plots) / 4)))
    ggsave(filename, plot = combined_plot, width = 20, height = 15)
  } else {
    message("No plots were generated due to data issues.")
  }
}


# For histograms
generate_and_save_plots(df_diga_differences, scales_diga_diff, "histograms_diga_diff.png","histogram")
generate_and_save_plots(df_dipa_differences, scales_dipa_diff, "histograms_dipa_diff.png","histogram")

# For QQ plots
generate_and_save_plots(df_dipa_differences,scales_dipa_diff, "qq_plots_dipa_diff.png","qq")
generate_and_save_plots(df_diga_differences,scales_diga_diff, "qq_plots_diga_diff.png","qq")


# Create a named vector for scale mapping
scale_name_mapping <- c(
  "six_min_walk_test_distance_in_meter" = "6-Minuten gehtest (MGT)",
  "development_1" = "Patienten Souverantitaet 19 Items (PS19)",
  "development_2" = "Patienten Souverantitaet 8 Dimensionen (PS8D)",
  "development_3" = "Patienten Souverantitaet Sicht Pflegender 19 Items (PS19-P)",
  "development_4" = "Patienten Souverantitaet Sicht Pflegender 8 Dimensionen (PS8D-P)",
  "development_5" = "Effekt Pflegender (ICG)",
  "HBP_MEAN" = "HBP MEAN",
  "PCS_12" = "PCS 12",
  "MCS_12" = "MCS 12",
  "HCS_MEAN" = "HCS MEAN",
  "DIGA_HELM_Evaluation_Overall" = "DIGA HELM",
  "HLS_SUM" = "HLS SUM",
  "Control" = "Kontrollgruppe",
  "Treatment" = "Interventionsgruppe"
)



# For histograms and QQ plots
generate_and_save_plots(df_diga, scales_diga, "histograms_diga.png", "histogram")
generate_and_save_plots(df_dipa, scales_dipa, "histograms_dipa.png", "histogram")
generate_and_save_plots(df_diga, scales_diga, "qq_plots_diga.png", "qq")
generate_and_save_plots(df_dipa, scales_dipa, "qq_plots_dipa.png", "qq")
