library(readxl)

file_path <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/andrewharley/Data.xlsx"


# List all sheet names
sheet_names <- excel_sheets(file_path)

# Read each sheet into a separate dataframe
df_chemistry <- read_excel(file_path, sheet = sheet_names[1])
df_biomass <- read_excel(file_path, sheet = sheet_names[2])
df_texture <- read_excel(file_path, sheet = sheet_names[3])


# Normality Check

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


# Outlier Screening

library(dplyr)

calculate_z_scores_and_transform <- function(data, vars, id_var, z_threshold = 3) {
  # Calculate z-scores for each variable
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~ (.-mean(.))/sd(.), .names = "z_{.col}"))
  
  # Flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  # Create a new dataframe with the ID column
  log_transformed_data <- data %>% select(id_var)
  
  # Log-transform the original variables where outliers are found
  for (var in vars) {
    flag_var <- paste0("flag_z_", var)
    if (any(z_score_data[[flag_var]] == "Outlier")) {
      log_var_name <- paste0("log_", var)
      log_transformed_data[[log_var_name]] <- log1p(data[[var]]) # log1p is used to avoid issues with log(0)
    }
  }
  
  return(log_transformed_data)
}




numeric_vars_chemistry <- names(df_chemistry)[sapply(df_chemistry, is.numeric)]
numeric_vars_biomass <- names(df_biomass)[sapply(df_biomass, is.numeric)]
numeric_vars_texture <- names(df_texture)[sapply(df_texture, is.numeric)]


# Chemistry dataframe
normality_results_chemistry <- calculate_stats(df_chemistry, numeric_vars_chemistry)

# Biomass dataframe
normality_results_biomass <- calculate_stats(df_biomass, numeric_vars_biomass)

# Texture dataframe
normality_results_texture <- calculate_stats(df_texture, numeric_vars_texture)

# Apply the function to each dataframe to create the log-transformed data
log_chemistry <- calculate_z_scores_and_transform(df_chemistry, numeric_vars_chemistry, id_var = "Site")
log_biomass <- calculate_z_scores_and_transform(df_biomass, numeric_vars_biomass, id_var = "Site")
log_texture <- calculate_z_scores_and_transform(df_texture, numeric_vars_texture, id_var = "Site")

# Merge the log-transformed columns back to the original dataframes by "Site"
df_chemistry <- left_join(df_chemistry, log_chemistry, by = "Site")
df_biomass <- left_join(df_biomass, log_biomass, by = "Site")
df_texture <- left_join(df_texture, log_texture, by = "Site")

numeric_vars_chemistry_transf <- names(df_chemistry)[sapply(df_chemistry, is.numeric)]
numeric_vars_biomass_transf <- names(df_biomass)[sapply(df_biomass, is.numeric)]
numeric_vars_texture_transf <- names(df_texture)[sapply(df_texture, is.numeric)]



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

# Calculate descriptive statistics for each dataframe
desc_stats_chemistry <- calculate_descriptive_stats(df_chemistry, numeric_vars_chemistry_transf)
desc_stats_biomass <- calculate_descriptive_stats(df_biomass, numeric_vars_biomass_transf)
desc_stats_texture <- calculate_descriptive_stats(df_texture, numeric_vars_texture_transf)



# Disaggregated Analysis


# Create a mapping of sites to their respective groups
site_group_mapping <- data.frame(
  Site = c(
    "AMF-01", "AMF-02", "AMF-03", "DMR-01", "DMR-02", "MPF-01", "MPF-02", "MPF-04", "MPF-05", "RGF-01", "RGF-02", "RGF-03", "RGF-04", "RGF-05", "SF-01", "SF-02", "SF-03", "SF-04", "SF-05", 
    "AMG-01", "AMG-02", "AMG-03", "AMG-04", "AMG-05", "AMG-06", "DMC-04", "DMC-05", "DMC-06", "MPG-01", "MPG-02", "MPG-03", "MPG-04", "MPG-05", "RGG-01", "RGG-02", "RGG-03", "RGG-04", "RGG-05", "SG-01", "SG-02", "SG-03", "SG-04", "SG-05", 
    "AMP-01", "AMP-02", "AMP-03", "AMP-04", "AMP-05", "AMP-06", "DMC-01", "DMC-02", "DMC-03", "DMC-07", "MPG-05", "MPP-01", "MPP-02", "MPP-04", "MPP-05", "RGP-01", "RGP-02", "RGP-03", "RGP-04", "RGP-05", "SP-01", "SP-02", "SP-03", "SP-04", "SP-05"
  ),
  Group = c(
    rep("Reference", 19),
    rep("Good Cover", 23),
    rep("Poor Cover", 26)
  )
)

# Merge the mapping with each dataframe
df_chemistry <- merge(df_chemistry, site_group_mapping, by = "Site")
df_biomass <- merge(df_biomass, site_group_mapping, by = "Site")
df_texture <- merge(df_texture, site_group_mapping, by = "Site")

# Descriptive statistics for Chemistry dataframe
desc_stats_chemistry_grouped <- df_chemistry %>%
  group_by(Group) %>%
  summarise(across(all_of(numeric_vars_chemistry_transf), list(
    Mean = ~ mean(.x, na.rm = TRUE),
    Median = ~ median(.x, na.rm = TRUE),
    SEM = ~ sd(.x, na.rm = TRUE) / sqrt(length(na.omit(.x))),
    SD = ~ sd(.x, na.rm = TRUE),
    Skewness = ~ skewness(.x, na.rm = TRUE),
    Kurtosis = ~ kurtosis(.x, na.rm = TRUE)
  )))

# Descriptive statistics for Biomass dataframe
desc_stats_biomass_grouped <- df_biomass %>%
  group_by(Group) %>%
  summarise(across(all_of(numeric_vars_biomass_transf), list(
    Mean = ~ mean(.x, na.rm = TRUE),
    Median = ~ median(.x, na.rm = TRUE),
    SEM = ~ sd(.x, na.rm = TRUE) / sqrt(length(na.omit(.x))),
    SD = ~ sd(.x, na.rm = TRUE),
    Skewness = ~ skewness(.x, na.rm = TRUE),
    Kurtosis = ~ kurtosis(.x, na.rm = TRUE)
  )))

# Descriptive statistics for Texture dataframe
desc_stats_texture_grouped <- df_texture %>%
  group_by(Group) %>%
  summarise(across(all_of(numeric_vars_texture_transf), list(
    Mean = ~ mean(.x, na.rm = TRUE),
    Median = ~ median(.x, na.rm = TRUE),
    SEM = ~ sd(.x, na.rm = TRUE) / sqrt(length(na.omit(.x))),
    SD = ~ sd(.x, na.rm = TRUE),
    Skewness = ~ skewness(.x, na.rm = TRUE),
    Kurtosis = ~ kurtosis(.x, na.rm = TRUE)
  )))


library(ggplot2)
library(dplyr)
library(tidyr)

# Function to create mean and SD plots divided by prefixes of the Site column
create_facet_mean_sd_plot <- function(data, variables, group_column, save_path_prefix = "mean_sd_facet_plot") {
  # Split the variables into chunks of 9
  chunks <- split(variables, ceiling(seq_along(variables) / 6))
  
  # Iterate over each chunk and create a plot
  for (i in seq_along(chunks)) {
    chunk_vars <- chunks[[i]]
    
    # Create a long format of the data for ggplot
    long_data <- data %>%
      pivot_longer(cols = all_of(chunk_vars), names_to = "Variable", values_to = "Value") %>%
      group_by(!!sym(group_column), Variable) %>%
      summarise(
        Mean = mean(Value, na.rm = TRUE),
        SD = sd(Value, na.rm = TRUE),
        .groups = 'drop'
      )
    
    # Plot with facet wrap
    p <- ggplot(long_data, aes(x = !!sym(group_column), y = Mean)) +
      geom_point(size = 3) +
      geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD), width = 0.2) +
      facet_wrap(~ Variable, scales = "free_y", ncol = 3, nrow = 2) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
      labs(x = "Group", y = "Mean ± SD")
    
    # Save the plot
    ggsave(filename = paste0(save_path_prefix, "_part", i, ".png"), plot = p, width = 12, height = 8)
    
    # Print the plot
    print(p)
  }
}

# Example usage for Chemistry dataframe
create_facet_mean_sd_plot(df_chemistry, numeric_vars_chemistry_transf, group_column = "Group", save_path_prefix = "mean_sd_facet_plot_chemistry")

# Mean and SD plots for Biomass dataframe
create_facet_mean_sd_plot(df_biomass, numeric_vars_biomass_transf, group_column = "Group", save_path_prefix = "mean_sd_plot_biomass_grouped")

# Mean and SD plots for Texture dataframe
create_facet_mean_sd_plot(df_texture, numeric_vars_texture_transf, group_column = "Group", save_path_prefix = "mean_sd_plot_texture_grouped")


create_facet_boxplots <- function(df, vars, group_column, save_path_prefix = "boxplot_facet_plot") {
  # Split the variables into chunks of 9
  chunks <- split(vars, ceiling(seq_along(vars) / 6))
  
  # Iterate over each chunk and create a plot
  for (i in seq_along(chunks)) {
    chunk_vars <- chunks[[i]]
    
    # Create a long format of the data for ggplot
    long_data <- df %>%
      pivot_longer(cols = all_of(chunk_vars), names_to = "Variable", values_to = "Value")
    
    # Plot with facet wrap
    p <- ggplot(long_data, aes_string(x = group_column, y = "Value")) +
      geom_boxplot() +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
      facet_wrap(~ Variable, scales = "free_y", ncol = 3, nrow = 2) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
      labs(x = "Group", y = "Value")
    
    # Save the plot
    ggsave(filename = paste0(save_path_prefix, "_part", i, ".png"), plot = p, width = 12, height = 8)
    
    # Print the plot
    print(p)
  }
}

# Boxplot facet plots for Chemistry dataframe
create_facet_boxplots(df_chemistry, numeric_vars_chemistry_transf, group_column = "Group", save_path_prefix = "boxplot_facet_plot_chemistry")

# Boxplot facet plots for Biomass dataframe
create_facet_boxplots(df_biomass, numeric_vars_biomass_transf, group_column = "Group", save_path_prefix = "boxplot_facet_plot_biomass")

# Boxplot facet plots for Texture dataframe
create_facet_boxplots(df_texture, numeric_vars_texture_transf, group_column = "Group", save_path_prefix = "boxplot_facet_plot_texture")


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

# Consolidate all data
export_data_list <- list(
  "Stats - Chemistry Grouped" = desc_stats_chemistry_grouped,
  "Stats - Biomass Grouped" = desc_stats_biomass_grouped,
  "Stats - Texture Grouped" = desc_stats_texture_grouped,
  "Descriptive Stats - Chemistry" = desc_stats_chemistry,
  "Descriptive Stats - Biomass" = desc_stats_biomass,
  "Descriptive Stats - Texture" = desc_stats_texture,
  "Normality Results - Chemistry" = normality_results_chemistry,
  "Normality Results - Biomass" = normality_results_biomass,
  "Normality Results - Texture" = normality_results_texture,
  "Z-Scores - Chemistry" = z_scores_chemistry,
  "Z-Scores - Biomass" = z_scores_biomass,
  "Z-Scores - Texture" = z_scores_texture,
  "Chemistry Data" = df_chemistry,
  "Biomass Data" = df_biomass,
  "Texture Data" = df_texture
)


# Export the consolidated data to an Excel file
save_apa_formatted_excel(export_data_list, "All_Data_Analysis.xlsx")



# New grouping scheme

# Define the new site groups
new_site_group_mapping <- data.frame(
  Site = c(
    "AMG-01", "AMG-02", "AMG-03", "AMP-01", "AMP-02", "AMP-03", "AMG-04", "AMG-05", "AMG-06", "AMP-04", "AMP-05", "AMP-06", 
    "AMF-01", "AMF-02", "AMF-03", "DMC-04", "DMC-05", "DMC-06", "DMC-01", "DMC-02", "DMC-03", "DMC-07", "DMR-01", "DMR-02",
    "MPG-01", "MPG-02", "MPG-03", "MPG-04", "MPG-05", "MPP-01", "MPP-02", "MPP-03", "MPP-04", "MPP-05",
    "MPF-01", "MPF-02", "MPF-03", "MPF-04", "MPF-05", "RGG-01", "RGG-02", "RGG-03", "RGG-04", "RGG-05", "RGP-01", "RGP-02", "RGP-03", "RGP-04", "RGP-05",
    "RGF-01", "RGF-02", "RGF-03", "RGF-04", "RGF-05", "SG-01", "SG-02", "SG-03", "SG-04", "SG-05", "SP-01", "SP-02", "SP-03", "SP-04", "SP-05",
    "SF-01", "SF-02", "SF-03", "SF-04", "SF-05", "DMT-01", "DMT-02", "DMT-03", "DMT-04", "DMT-05", "DMT-06", "RGF-03", "MPF-01", "MPF-02", "MPF-03", "MPF-04", "MPF-05", "AMF-02", "AMF-03", "DMR-02",
    "RGF-02", "RGF-05", "AMF-01", "DMR-01"
  ),
  Group = c(
    rep("Atlas Grey GC", 3), rep("Atlas Grey PC", 3), rep("Atlas Yellow GC", 3), rep("Atlas Yellow PC", 3),
    rep("Atlas R", 3), rep("Doctor GC", 3), rep("Doctor PC", 4), rep("Doctor R", 2),
    rep("Mineral Park GC", 5), rep("Mineral Park PC", 5), rep("Mineral Park R", 5),
    rep("Rawley Gulch GC", 5), rep("Rawley Gulch PC", 5), rep("Rawley Gulch R", 5),
    rep("Snowden GC", 5), rep("Snowden PC", 5), rep("Snowden R", 5),
    rep("Doctor Base", 3), rep("Doctor Alluvial", 3),
    rep("Conifer", 9), rep("Aspen", 2), rep("Willow", 2)
  )
)

# Merge the new grouping with each dataframe
df_chemistry_newgroup <- df_chemistry %>%
  select(-Group) %>%
  merge(new_site_group_mapping, by = "Site", all.x = TRUE)

df_biomass_newgroup <- df_biomass %>%
  select(-Group) %>%
  merge(new_site_group_mapping, by = "Site", all.x = TRUE)

df_texture_newgroup <- df_texture %>%
  select(-Group) %>%
  merge(new_site_group_mapping, by = "Site", all.x = TRUE)


# Descriptive statistics for Chemistry dataframe
desc_stats_chemistry_grouped2 <- df_chemistry_newgroup %>%
  group_by(Group) %>%
  summarise(across(all_of(numeric_vars_chemistry_transf), list(
    Mean = ~ mean(.x, na.rm = TRUE),
    Median = ~ median(.x, na.rm = TRUE),
    SEM = ~ sd(.x, na.rm = TRUE) / sqrt(length(na.omit(.x))),
    SD = ~ sd(.x, na.rm = TRUE),
    Skewness = ~ skewness(.x, na.rm = TRUE),
    Kurtosis = ~ kurtosis(.x, na.rm = TRUE)
  )))

# Descriptive statistics for Biomass dataframe
desc_stats_biomass_grouped2 <- df_biomass_newgroup %>%
  group_by(Group) %>%
  summarise(across(all_of(numeric_vars_biomass_transf), list(
    Mean = ~ mean(.x, na.rm = TRUE),
    Median = ~ median(.x, na.rm = TRUE),
    SEM = ~ sd(.x, na.rm = TRUE) / sqrt(length(na.omit(.x))),
    SD = ~ sd(.x, na.rm = TRUE),
    Skewness = ~ skewness(.x, na.rm = TRUE),
    Kurtosis = ~ kurtosis(.x, na.rm = TRUE)
  )))

# Descriptive statistics for Texture dataframe
desc_stats_texture_grouped2 <- df_texture_newgroup %>%
  group_by(Group) %>%
  summarise(across(all_of(numeric_vars_texture_transf), list(
    Mean = ~ mean(.x, na.rm = TRUE),
    Median = ~ median(.x, na.rm = TRUE),
    SEM = ~ sd(.x, na.rm = TRUE) / sqrt(length(na.omit(.x))),
    SD = ~ sd(.x, na.rm = TRUE),
    Skewness = ~ skewness(.x, na.rm = TRUE),
    Kurtosis = ~ kurtosis(.x, na.rm = TRUE)
  )))



create_facet_mean_sd_plot <- function(data, variables, group_column, save_path_prefix = "mean_sd_facet_plot") {
  # Iterate over each variable and create a plot
  for (var in variables) {
    # Sanitize the variable name for use in the filename
    sanitized_var <- gsub("[^A-Za-z0-9_]", "_", var)  # Replace non-alphanumeric characters with underscore
    
    # Create a long format of the data for ggplot
    long_data <- data %>%
      pivot_longer(cols = var, names_to = "Variable", values_to = "Value") %>%
      group_by(!!sym(group_column), Variable) %>%
      summarise(
        Mean = mean(Value, na.rm = TRUE),
        SD = sd(Value, na.rm = TRUE),
        .groups = 'drop'
      )
    
    # Plot
    p <- ggplot(long_data, aes(x = !!sym(group_column), y = Mean)) +
      geom_point(size = 3) +
      geom_errorbar(aes(ymin = Mean - SD, ymax = Mean + SD), width = 0.2) +
      theme_minimal() +
      theme(plot.title = element_text(hjust = 0.5), axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
      labs(x = "Group", y = "Mean ± SD", title = var)
    
    # Save the plot with the sanitized variable name in the filename
    ggsave(filename = paste0(save_path_prefix, "_", sanitized_var, ".png"), plot = p, width = 12, height = 8)
    
    # Print the plot
    print(p)
  }
}


# Example usage for Chemistry dataframe
create_facet_mean_sd_plot(df_chemistry_newgroup, numeric_vars_chemistry_transf, group_column = "Group", save_path_prefix = "mean_sd_facet_plot_chemistry_newgroup")

# Mean and SD plots for Biomass dataframe
create_facet_mean_sd_plot(df_biomass_newgroup, numeric_vars_biomass_transf, group_column = "Group", save_path_prefix = "mean_sd_plot_biomass_newgroup")

# Mean and SD plots for Texture dataframe
create_facet_mean_sd_plot(df_texture_newgroup, numeric_vars_texture_transf, group_column = "Group", save_path_prefix = "mean_sd_plot_texture_newgroup")


create_facet_boxplots <- function(df, vars, group_column, save_path_prefix = "boxplot_facet_plot") {
  # Split the variables into chunks of 9
  chunks <- split(vars, ceiling(seq_along(vars) / 1))
  
  # Iterate over each chunk and create a plot
  for (i in seq_along(chunks)) {
    chunk_vars <- chunks[[i]]
    
    # Create a long format of the data for ggplot
    long_data <- df %>%
      pivot_longer(cols = all_of(chunk_vars), names_to = "Variable", values_to = "Value")
    
    # Plot with facet wrap
    p <- ggplot(long_data, aes_string(x = group_column, y = "Value")) +
      geom_boxplot() +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
      facet_wrap(~ Variable, scales = "free_y", ncol = 1, nrow = 1) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1)) +
      labs(x = "Group", y = "Value")
    
    # Save the plot
    ggsave(filename = paste0(save_path_prefix, "_part", i, ".png"), plot = p, width = 12, height = 8)
    
    # Print the plot
    print(p)
  }
}

# Boxplot facet plots for Chemistry dataframe
create_facet_boxplots(df_chemistry_newgroup, numeric_vars_chemistry_transf, group_column = "Group", save_path_prefix = "boxplot_facet_plot_chemistry_newgroup")

# Boxplot facet plots for Biomass dataframe
create_facet_boxplots(df_biomass_newgroup, numeric_vars_biomass_transf, group_column = "Group", save_path_prefix = "boxplot_facet_plot_biomass_newgroup")

# Boxplot facet plots for Texture dataframe
create_facet_boxplots(df_texture_newgroup, numeric_vars_texture_transf, group_column = "Group", save_path_prefix = "boxplot_facet_plot_texture_newgroup")


# Consolidate all data
export_data_list <- list(
  "Stats - Chemistry New Group" = desc_stats_chemistry_grouped2,
  "Stats - Biomass New Group" = desc_stats_biomass_grouped2,
  "Stats - Texture New Group" = desc_stats_texture_grouped2
  
)


# Export the consolidated data to an Excel file
save_apa_formatted_excel(export_data_list, "All_Data_Analysis_NewGroups.xlsx")
