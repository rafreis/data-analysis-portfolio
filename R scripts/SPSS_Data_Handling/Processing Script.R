setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/chrislover")

library(haven)
df <- read_sav("Dataset (BLINDED) GAD7 checked 14.05.2025.sav")

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

colnames(df_renamed)


library(ggplot2)
library(dplyr)
library(tidyr)

# Function to create dot-and-whisker plots with mean and save them
create_mean_sd_plot <- function(data, variables, factor_column, event_column, save_path = NULL) {
  # Define the specific levels for redcap_event_name
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  # Filter and format the data
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    filter(!!sym(event_column) %in% event_levels) %>%  # Keep only specific event levels
    mutate(!!sym(event_column) := factor(!!sym(event_column), levels = event_levels)) %>%  # Set the correct order
    group_by(!!sym(event_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      .groups = 'drop'
    )
  
  # Clean variable names for better readability in the titles
  long_data$Variable <- gsub("_", " ", long_data$Variable)
  
  # Create the plot with mean points and error bars for SD
  p <- ggplot(long_data, aes(x = !!sym(event_column), y = Mean)) +
    geom_point(aes(color = !!sym(event_column)), size = 3) +
    geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD, color = !!sym(event_column)), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y", labeller = label_wrap_gen()) +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Mean and Standard Deviation Across Periods", 
      x = "Event Name",
      y = "Value",
      color = "Event Name"
    ) +
    scale_color_discrete(name = "Event Name")
  
  # Print the plot
  print(p)
  
  # Save the plot if a save path is provided
  if (!is.null(save_path)) {
    ggsave(filename = save_path, plot = p, width = 8, height = 6)
  }
  
  # Return the plot object
  return(p)
}

# Example usage:
# Assuming you have a dataframe `df` with the required variables and columns
plot <- create_mean_sd_plot(df, 
                            variables = c("COPM_TotalPerform", 
                                          "ACS3Global", 
                                          "GeneralisedAnxietyDisorder7GAD7Scale_complete",
                                          "BarthelIndexofActivitiesofDailyLiving_complete",
                                          "ParkinsonsDiseaseQualityofLifeQuestionnairePDQ39_compl",
                                          "EQ5D_complete"), 
                            factor_column = "Adele_Use", 
                            event_column = "redcap_event_name", 
                            save_path = "mean_sd_plot.png")




library(ggplot2)
library(dplyr)
library(tidyr)

# Function to create dot-and-whisker plots with median and save them
create_median_sd_plot <- function(data, variables, factor_column, event_column, save_path = NULL) {
  # Define the specific levels for redcap_event_name
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  # Filter and format the data
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    filter(!!sym(event_column) %in% event_levels) %>%  # Keep only specific event levels
    mutate(!!sym(event_column) := factor(!!sym(event_column), levels = event_levels)) %>%  # Set the correct order
    group_by(!!sym(event_column), Variable) %>%
    summarise(
      Median = median(Value, na.rm = TRUE),
      IQR = IQR(Value, na.rm = TRUE),  # Interquartile Range for error bars
      .groups = 'drop'
    )
  
  # Clean variable names for better readability in the titles
  long_data$Variable <- gsub("_", " ", long_data$Variable)
  
  # Create the plot with median points and error bars for IQR
  p <- ggplot(long_data, aes(x = !!sym(event_column), y = Median)) +
    geom_point(aes(color = !!sym(event_column)), size = 3) +
    geom_errorbar(aes(ymin = Median - IQR/2, ymax = Median + IQR/2, color = !!sym(event_column)), width = 0.2) +
    facet_wrap(~ Variable, scales = "free_y", labeller = label_wrap_gen()) +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)) +
    labs(
      title = "Median and Interquartile Range Across Periods", 
      x = "Event Name",
      y = "Value",
      color = "Event Name"
    ) +
    scale_color_discrete(name = "Event Name")
  
  # Print the plot
  print(p)
  
  # Save the plot if a save path is provided
  if (!is.null(save_path)) {
    ggsave(filename = save_path, plot = p, width = 8, height = 6)
  }
  
  # Return the plot object
  return(p)
}

# Example usage:
# Assuming you have a dataframe `df` with the required variables and columns
plot <- create_median_sd_plot(df, 
                              variables = c("COPM_TotalPerform", 
                                            "ACS3Global", 
                                            "GeneralisedAnxietyDisorder7GAD7Scale_complete",
                                            "BarthelIndexofActivitiesofDailyLiving_complete",
                                            "ParkinsonsDiseaseQualityofLifeQuestionnairePDQ39_compl",
                                            "EQ5D_complete"), 
                              factor_column = "Adele_Use", 
                              event_column = "redcap_event_name", 
                              save_path = "median_iqr_plot.png")






library(dplyr)

# Assuming df is your dataframe
df_renamed <- df %>%
  rename(
    "COPM Total Performance" = COPM_TotalPerform,
    "ACS 3 Global" = ACS3Global,
    "GAD-7 Complete" = GeneralisedAnxietyDisorder7GAD7Scale_complete,
    "Barthel Index ADL Complete" = BarthelIndexofActivitiesofDailyLiving_complete,
    "PDQ-39 Complete" = ParkinsonsDiseaseQualityofLifeQuestionnairePDQ39_compl,
    "EQ-5D Complete" = EQ5D_complete
  )

# After renaming, you can call the plotting functions as usual
plot_mean <- create_mean_sd_plot(df_renamed, 
                                 variables = c("COPM Total Performance", 
                                               "ACS 3 Global", 
                                               "GAD-7 Complete",
                                               "Barthel Index ADL Complete",
                                               "PDQ-39 Complete",
                                               "EQ-5D Complete"), 
                                 factor_column = "Adele_Use", 
                                 event_column = "redcap_event_name", 
                                 save_path = "mean_sd_plot.png")

plot_median <- create_median_sd_plot(df_renamed, 
                                     variables = c("COPM Total Performance", 
                                                   "ACS 3 Global", 
                                                   "GAD-7 Complete",
                                                   "Barthel Index ADL Complete",
                                                   "PDQ-39 Complete",
                                                   "EQ-5D Complete"), 
                                     factor_column = "Adele_Use", 
                                     event_column = "redcap_event_name", 
                                     save_path = "median_iqr_plot.png")







create_mean_sd_plot_separate <- function(data, variables, factor_column, event_column, save_path_prefix = NULL) {
  # Define the specific levels for redcap_event_name
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  # Filter and format the data
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    filter(!!sym(event_column) %in% event_levels) %>%  # Keep only specific event levels
    mutate(!!sym(event_column) := factor(!!sym(event_column), levels = event_levels)) %>%  # Set the correct order
    group_by(!!sym(event_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      .groups = 'drop'
    )
  
  # Loop through each variable and create individual plots
  for (var in variables) {
    subset_data <- long_data %>% filter(Variable == gsub("_", " ", var))
    
    # Create the plot with mean points and error bars for SD
    p <- ggplot(subset_data, aes(x = !!sym(event_column), y = Mean)) +
      geom_point(aes(color = !!sym(event_column)), size = 3) +
      geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD, color = !!sym(event_column)), width = 0.2) +
      labs(
        title = paste("Mean and Standard Deviation -", var), 
        x = "Event Name",
        y = "Value",
        color = "Event Name"
      ) +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1))
    
    # Save the plot if a save path prefix is provided
    if (!is.null(save_path_prefix)) {
      ggsave(filename = paste0(save_path_prefix, "_", var, ".png"), plot = p, width = 8, height = 6)
    }
  }
}




create_median_sd_plot_separate <- function(data, variables, factor_column, event_column, save_path_prefix = NULL) {
  # Define the specific levels for redcap_event_name
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  # Filter and format the data
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    filter(!!sym(event_column) %in% event_levels) %>%  # Keep only specific event levels
    mutate(!!sym(event_column) := factor(!!sym(event_column), levels = event_levels)) %>%  # Set the correct order
    group_by(!!sym(event_column), Variable) %>%
    summarise(
      Median = median(Value, na.rm = TRUE),
      IQR = IQR(Value, na.rm = TRUE),  # Interquartile Range for error bars
      .groups = 'drop'
    )
  
  # Loop through each variable and create individual plots
  for (var in variables) {
    subset_data <- long_data %>% filter(Variable == gsub("_", " ", var))
    
    # Create the plot with median points and error bars for IQR
    p <- ggplot(subset_data, aes(x = !!sym(event_column), y = Median)) +
      geom_point(aes(color = !!sym(event_column)), size = 3) +
      geom_errorbar(aes(ymin = Median - IQR/2, ymax = Median + IQR/2, color = !!sym(event_column)), width = 0.2) +
      labs(
        title = paste("Median and Interquartile Range -", var), 
        x = "Event Name",
        y = "Value",
        color = "Event Name"
      ) +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1))
    
    # Save the plot if a save path prefix is provided
    if (!is.null(save_path_prefix)) {
      ggsave(filename = paste0(save_path_prefix, "_", var, ".png"), plot = p, width = 8, height = 6)
    }
  }
}


create_mean_sd_plot_separate(df_renamed, 
                             variables = c("COPM Total Performance", 
                                           "ACS 3 Global", 
                                           "GAD-7 Complete",
                                           "Barthel Index ADL Complete",
                                           "PDQ-39 Complete",
                                           "EQ-5D Complete"), 
                             factor_column = "Adele_Use", 
                             event_column = "redcap_event_name", 
                             save_path_prefix = "mean_sd_plot")


create_median_sd_plot_separate(df_renamed, 
                               variables = c("COPM Total Performance", 
                                             "ACS 3 Global", 
                                             "GAD-7 Complete",
                                             "Barthel Index ADL Complete",
                                             "PDQ-39 Complete",
                                             "EQ-5D Complete"), 
                               factor_column = "Adele_Use", 
                               event_column = "redcap_event_name", 
                               save_path_prefix = "median_iqr_plot")







library(ggplot2)
library(dplyr)
library(tidyr)
# Separate by group

df_renamed <- df_renamed %>%
  rename(`Team allocation` = CRTAllocation_Blinded)


create_mean_sd_plot_separate <- function(data, variables, factor_column, event_column, save_path_prefix = NULL) {
  # Define the specific levels for event_column
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  # Filter and format the data
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    filter(!!sym(event_column) %in% event_levels) %>%
    mutate(!!sym(event_column) := factor(!!sym(event_column), levels = event_levels)) %>%
    group_by(!!sym(factor_column), !!sym(event_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      SD = sd(Value, na.rm = TRUE),
      .groups = 'drop'
    )
  
  # Create plots for each group
  for (var in variables) {
    subset_data <- long_data %>% filter(Variable == gsub("_", " ", var))
    
    p <- ggplot(subset_data, aes(x = !!sym(event_column), y = Mean, color = !!sym(factor_column))) +
      geom_point(size = 3) +
      geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD), width = 0.2) +
      facet_wrap(~Variable, scales = "free_y") +
      labs(title = "Mean and Standard Deviation by Group", x = "Time Point", y = "Value") +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
            plot.title = element_text(hjust = 0.5))  # Centering the title
    
    if (!is.null(save_path_prefix)) {
      ggsave(filename = paste0(save_path_prefix, "_mean_sd_", var, ".png"), plot = p, width = 8, height = 6)
    }
  }
}

create_median_sd_plot_separate <- function(data, variables, factor_column, event_column, save_path_prefix = NULL) {
  # Define the specific levels for event_column
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  # Filter and format the data
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    filter(!!sym(event_column) %in% event_levels) %>%
    mutate(!!sym(event_column) := factor(!!sym(event_column), levels = event_levels)) %>%
    group_by(!!sym(factor_column), !!sym(event_column), Variable) %>%
    summarise(
      Median = median(Value, na.rm = TRUE),
      IQR = IQR(Value, na.rm = TRUE),
      .groups = 'drop'
    )
  
  # Create plots for each group
  for (var in variables) {
    subset_data <- long_data %>% filter(Variable == gsub("_", " ", var))
    
    p <- ggplot(subset_data, aes(x = !!sym(event_column), y = Median, color = !!sym(factor_column))) +
      geom_point(size = 3) +
      geom_errorbar(aes(ymin = Median - IQR/2, ymax = Median + IQR/2), width = 0.2) +
      facet_wrap(~Variable, scales = "free_y") +
      labs(title = "Median and Interquartile Range by Group", x = "Time Point", y = "Value") +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
            plot.title = element_text(hjust = 0.5))  # Centering the title
    
    if (!is.null(save_path_prefix)) {
      ggsave(filename = paste0(save_path_prefix, "_median_iqr_", var, ".png"), plot = p, width = 8, height = 6)
    }
  }
}

df_renamed <- df_renamed %>%
  rename(`COPM Total Satisfaction` = COPMTotalSatis)


create_mean_sd_plot_separate(df_renamed, 
                             variables = c("COPM Total Performance", 
                                           "ACS 3 Global", 
                                           "GAD-7 Complete",
                                           "Barthel Index ADL Complete",
                                           "PDQ-39 Complete",
                                           "EQ-5D Complete",
                                           "COPM Total Satisfaction"), 
                             factor_column = "Team allocation", 
                             event_column = "redcap_event_name", 
                             save_path_prefix = "mean_sd_plot")


create_median_sd_plot_separate(df_renamed, 
                               variables = c("COPM Total Performance", 
                                             "ACS 3 Global", 
                                             "GAD-7 Complete",
                                             "Barthel Index ADL Complete",
                                             "PDQ-39 Complete",
                                             "EQ-5D Complete",
                                             "COPM Total Satisfaction"), 
                               factor_column = "Team allocation", 
                               event_column = "redcap_event_name", 
                               save_path_prefix = "median_iqr_plot")





library(ggplot2)
library(dplyr)
library(tidyr)

# Define the plotting function for one variable
create_copm_satisfaction_plot <- function(data, variable_name, group_column = "team_allocation", 
                                          event_column = "redcap_event_name", save_path = NULL,
                                          summary_stat = c("mean", "median")) {
  summary_stat <- match.arg(summary_stat)
  
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  summary_data <- data %>%
    filter(!!sym(event_column) %in% event_levels) %>%
    mutate(!!sym(event_column) := factor(!!sym(event_column), levels = event_levels)) %>%
    group_by(!!sym(group_column), !!sym(event_column)) %>%
    summarise(
      Center = if (summary_stat == "mean") mean(!!sym(variable_name), na.rm = TRUE) else median(!!sym(variable_name), na.rm = TRUE),
      Spread = if (summary_stat == "mean") sd(!!sym(variable_name), na.rm = TRUE) else IQR(!!sym(variable_name), na.rm = TRUE),
      .groups = 'drop'
    )
  
  # Rename for consistent plotting
  colnames(summary_data) <- c("Group", "Event", "Center", "Spread")
  
  p <- ggplot(summary_data, aes(x = Event, y = Center, color = Group)) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = Center - Spread/2, ymax = Center + Spread/2), width = 0.2) +
    facet_wrap(~"COPM Total Satisfaction") +
    labs(
      title = "Mean and Standard Deviation by Group",
      x = "Time Point",
      y = "Value",
      color = "Team allocation"
    ) +
    theme_minimal() +
    theme(
      axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
      plot.title = element_text(hjust = 0.5)
    )
  
  print(p)
  
  if (!is.null(save_path)) {
    ggsave(filename = save_path, plot = p, width = 8, height = 6)
  }
  
  return(p)
}



colnames(df_renamed)



library(ggplot2)
library(dplyr)

create_individual_dot_plots <- function(data, variables, group_column, event_column, save_path_prefix = NULL) {
  event_levels <- c("Baseline with GAD7", "12 weeks", "24 weeks")
  
  data <- data %>%
    mutate(`Team allocation` = recode(`Team allocation`,
                                      "A" = "A- Usual care occupational therapy",
                                      "B" = "B- OBtAIN-PD")) %>%
    filter(.data[[event_column]] %in% event_levels) %>%
    mutate(!!sym(event_column) := factor(.data[[event_column]], levels = event_levels))
  
  for (var in variables) {
    p <- ggplot(data, aes(x = .data[[event_column]], y = .data[[var]], color = .data[[group_column]])) +
      geom_jitter(width = 0.2, height = 0, alpha = 0.7, size = 2) +
      labs(
        title = paste("Individual Scores -", var),
        x = "Time Point",
        y = "Score",
        color = "Group"
      ) +
      theme_minimal() +
      theme(
        axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1),
        plot.title = element_text(hjust = 0.5)
      )
    
    print(p)
    
    if (!is.null(save_path_prefix)) {
      ggsave(filename = paste0(save_path_prefix, "_", gsub(" ", "_", var), "_individual_plot.png"), plot = p, width = 8, height = 6)
    }
  }
}


create_individual_dot_plots(df_renamed,
                            variables = c("COPM Total Performance",
                                          "COPM Total Satisfaction",
                                          "ACS 3 Global",
                                          "GAD-7 Complete"),
                            group_column = "Team allocation",
                            event_column = "redcap_event_name",
                            save_path_prefix = "individual_scores")



summary_table <- df_renamed %>%
  filter(redcap_event_name %in% c("Baseline with GAD7", "12 weeks", "24 weeks")) %>%
  mutate(`Team allocation` = recode(`Team allocation`,
                                    "A" = "A- Usual care occupational therapy",
                                    "B" = "B- OBtAIN-PD")) %>%
  group_by(`Team allocation`, redcap_event_name) %>%
  summarise(across(c(`COPM Total Performance`, 
                     `COPM Total Satisfaction`,
                     `ACS 3 Global`,
                     `GAD-7 Complete`),
                   list(mean = ~mean(.x, na.rm = TRUE),
                        sd = ~sd(.x, na.rm = TRUE)),
                   .names = "{.col}_{.fn}"),
            .groups = "drop")

print(summary_table)

library(openxlsx)


write.xlsx(summary_table, file = "summary_table.xlsx", rowNames = FALSE)


create_individual_dot_plots(df_renamed,
                            variables = c("COPM Total Performance",
                                          "COPM Total Satisfaction",
                                          "ACS 3 Global",
                                          "GADSCORE"),   # <-- alterado aqui
                            group_column = "Team allocation",
                            event_column = "redcap_event_name",
                            save_path_prefix = "individual_scores")




summary_table <- df_renamed %>%
  filter(redcap_event_name %in% c("Baseline with GAD7", "12 weeks", "24 weeks")) %>%
  mutate(`Team allocation` = recode(`Team allocation`,
                                    "A" = "A- Usual care occupational therapy",
                                    "B" = "B- OBtAIN-PD")) %>%
  group_by(`Team allocation`, redcap_event_name) %>%
  summarise(across(c(`COPM Total Performance`, 
                     `COPM Total Satisfaction`,
                     `ACS 3 Global`,
                     GADSCORE),   # <-- alterado aqui
                   list(mean = ~mean(.x, na.rm = TRUE),
                        sd = ~sd(.x, na.rm = TRUE)),
                   .names = "{.col}_{.fn}"),
            .groups = "drop")

write.xlsx(summary_table, file = "summary_table2.xlsx", rowNames = FALSE)

