# ---- Data Loading ----
# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/jeffsmith")

library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Scenario Counts_clean.xlsx")


clean_dataframe <- function(df) {
  # Clean column names
  names(df) <- gsub(" ", "_", trimws(names(df)))
  names(df) <- gsub("\\s+", "_", names(df))
  names(df) <- gsub("\\(", "_", names(df))
  names(df) <- gsub("\\)", "_", names(df))
  names(df) <- gsub("-", "_", names(df))
  names(df) <- gsub("/", "_", names(df))
  names(df) <- gsub("\\\\", "_", names(df)) 
  names(df) <- gsub("\\?", "", names(df))
  names(df) <- gsub("'", "", names(df))
  names(df) <- gsub(",", "_", names(df))
  names(df) <- gsub("\\$", "", names(df))
  names(df) <- gsub("\\+", "_", names(df))
  
  # Clean character values
  df <- data.frame(lapply(df, function(x) {
    if (is.character(x)) {
      x <- tolower(trimws(x))
    }
    return(x)
  }), stringsAsFactors = FALSE)
  
  return(df)
}

df <- clean_dataframe(df)

str(df)
options(max.print = 10000)  # or any large number you need
colnames(df)


library(dplyr)
library(tidyr)
library(stringr)

df[, -1] <- lapply(df[, -1], as.character)

df_long <- df %>%
  pivot_longer(
    cols = -Personality.Traits,
    names_to = "scenario_option",
    values_to = "value"
  ) %>%
  filter(!is.na(value) & value == 1) %>% # keep only mapped traits
  mutate(
    scenario = str_extract(scenario_option, "\\d+"),
    option = str_extract(scenario_option, "[a-d]"),
    scenario = as.integer(scenario)
  ) %>%
  select(trait = Personality.Traits, scenario, option)

# View result
head(df_long)


library(dplyr)

# 1. Total number of unique scenarios
n_scenarios <- df_long %>%
  distinct(scenario) %>%
  nrow()
print(paste("Total unique scenarios:", n_scenarios))

# 2. Traits per answer option (scenario + option)
traits_per_option <- df_long %>%
  group_by(scenario, option) %>%
  summarise(traits_count = n(), .groups = "drop")

# Summary statistics
summary(traits_per_option$traits_count)


# Identify the option with the maximum number of traits
outlier_option <- traits_per_option %>%
  filter(traits_count == max(traits_count))

print(outlier_option)

# View all traits linked to this outlier
df_long %>%
  filter(scenario == outlier_option$scenario, option == outlier_option$option) %>%
  arrange(trait)

# Clean df_long of invalid rows
df_long <- df_long %>%
  filter(!is.na(scenario))

# Recalculate traits_per_option after cleaning
traits_per_option <- df_long %>%
  group_by(scenario, option) %>%
  summarise(traits_count = n(), .groups = "drop")

# Verify fix
summary(traits_per_option$traits_count)


library(tidyr)
library(dplyr)

# Create a unique identifier for each answer option
df_long <- df_long %>%
  mutate(answer_id = paste0(scenario, option))

# Create the wide binary matrix
coverage_matrix <- df_long %>%
  select(answer_id, trait) %>%
  mutate(value = 1) %>%
  pivot_wider(
    names_from = trait,
    values_from = value,
    values_fill = list(value = 0)
  )

# View result
print(dim(coverage_matrix))
head(coverage_matrix[, 1:6]) # just preview first few traits

str(coverage_matrix)


library(tidyr)

# ---- Functions ----
get_minimal_coverage <- function(coverage_matrix, traits_to_cover = NULL) {
  answer_ids <- coverage_matrix$answer_id
  trait_matrix <- as.matrix(coverage_matrix[, -1])
  
  if (is.null(traits_to_cover)) {
    traits_to_cover <- colnames(trait_matrix)
  }
  
  trait_matrix <- trait_matrix[, traits_to_cover, drop = FALSE]
  covered_traits <- rep(FALSE, length(traits_to_cover))
  names(covered_traits) <- traits_to_cover
  selected_answers <- c()
  
  while (!all(covered_traits)) {
    new_coverage <- apply(trait_matrix, 1, function(row) {
      sum(row & !covered_traits)
    })
    
    best_answer_index <- which.max(new_coverage)
    selected_answers <- c(selected_answers, answer_ids[best_answer_index])
    covered_traits <- covered_traits | trait_matrix[best_answer_index, ] == 1
    if (all(covered_traits)) break
  }
  return(selected_answers)
}

get_max_profession_coverage <- function(coverage_matrix, df) {
  trait_weights <- df %>%
    select(Personality.Traits, Count.of.Professions.with.this.trait) %>%
    mutate(Personality.Traits = as.character(Personality.Traits)) %>%
    { setNames(as.numeric(.$Count.of.Professions.with.this.trait), .$Personality.Traits) }
  
  common_traits <- intersect(names(coverage_matrix)[-1], names(trait_weights))
  answer_ids <- coverage_matrix$answer_id
  trait_matrix <- coverage_matrix[, common_traits] %>%
    apply(2, function(x) as.numeric(as.character(x))) %>%
    as.matrix()
  
  weights <- trait_weights[common_traits]
  covered_traits <- rep(FALSE, length(common_traits))
  names(covered_traits) <- common_traits
  selected_answers <- c()
  
  while (!all(covered_traits)) {
    new_coverage <- apply(trait_matrix, 1, function(row) {
      row <- as.numeric(row)
      sum(weights * (row == 1) * (!covered_traits))
    })
    if (max(new_coverage) == 0) break
    best_answer_index <- which.max(new_coverage)
    selected_answers <- c(selected_answers, answer_ids[best_answer_index])
    covered_traits <- covered_traits | trait_matrix[best_answer_index, ] == 1
  }
  return(selected_answers)
}

get_professions_covered <- function(selected_answers, coverage_matrix, df) {
  selected_traits <- coverage_matrix %>%
    filter(answer_id %in% selected_answers) %>%
    select(-answer_id) %>%
    summarise_all(max) %>%
    pivot_longer(cols = everything(), names_to = "trait", values_to = "covered") %>%
    filter(covered == 1) %>%
    pull(trait)
  
  covered_professions <- df %>%
    filter(Personality.Traits %in% selected_traits) %>%
    pull(Count.of.Professions.with.this.trait) %>%
    as.numeric() %>%
    sum(na.rm = TRUE)
  
  return(covered_professions)
}

# ---- Main Analysis ----
total_professions <- df$Count.of.Professions.with.this.trait %>%
  as.numeric() %>%
  sum(na.rm = TRUE)

standard_set <- get_minimal_coverage(coverage_matrix)
standard_scenarios <- df_long %>% filter(answer_id %in% standard_set) %>%
  distinct(scenario) %>% nrow()
standard_professions <- get_professions_covered(standard_set, coverage_matrix, df)

weighted_set <- get_max_profession_coverage(coverage_matrix, df)
weighted_scenarios <- df_long %>% filter(answer_id %in% weighted_set) %>%
  distinct(scenario) %>% nrow()
weighted_professions <- get_professions_covered(weighted_set, coverage_matrix, df)

comparison_table <- data.frame(
  Method = c("Standard Trait Coverage", "Weighted Profession Coverage"),
  Answer_Options_Selected = c(length(standard_set), length(weighted_set)),
  Scenarios_Selected = c(standard_scenarios, weighted_scenarios),
  Percent_Reduction = round((1 - c(length(standard_set), length(weighted_set)) / 1685) * 100, 1),
  Professions_Covered = c(standard_professions, weighted_professions),
  Percent_Professions_Covered = round(c(standard_professions, weighted_professions) / total_professions * 100, 1)
)
print(comparison_table)

# ---- Simulation with added Professions Covered ----
simulate_coverage_reduction <- function(coverage_matrix, df_long, df, thresholds = c(1.0, 0.9, 0.8, 0.7, 0.6, 0.5)) {
  all_traits <- colnames(coverage_matrix)[-1]
  total_traits <- length(all_traits)
  total_professions <- sum(as.numeric(df$Count.of.Professions.with.this.trait), na.rm = TRUE)
  
  results <- data.frame()
  
  for (threshold in thresholds) {
    n_traits_required <- ceiling(total_traits * threshold)
    if (threshold < 1.0) {
      set.seed(123)
      traits_to_cover <- sample(all_traits, n_traits_required)
    } else {
      traits_to_cover <- all_traits
    }
    
    selected <- get_minimal_coverage(coverage_matrix, traits_to_cover)
    n_scenarios <- df_long %>%
      filter(answer_id %in% selected) %>%
      distinct(scenario) %>%
      nrow()
    professions_covered <- get_professions_covered(selected, coverage_matrix, df)
    
    results <- rbind(results, data.frame(
      Coverage_Percentage = threshold,
      Traits_Required = n_traits_required,
      Answers_Selected = length(selected),
      Scenarios_Selected = n_scenarios,
      Professions_Covered = professions_covered,
      Percent_Professions_Covered = round(professions_covered / total_professions * 100, 1),
      Percent_Reduction = round((1 - length(selected) / 1685) * 100, 1)
    ))
  }
  
  return(results)
}

results_table <- simulate_coverage_reduction(coverage_matrix, df_long, df)
print(results_table)

# ---- Add weighted to simulation table ----
weighted_row <- data.frame(
  Coverage_Percentage = "Weighted Strategy",
  Traits_Required = NA,
  Answers_Selected = length(weighted_set),
  Scenarios_Selected = weighted_scenarios,
  Professions_Covered = weighted_professions,
  Percent_Professions_Covered = round(weighted_professions / total_professions * 100, 1),
  Percent_Reduction = round((1 - length(weighted_set) / 1685) * 100, 1)
)
results_table$Coverage_Percentage <- as.character(results_table$Coverage_Percentage)
results_table_combined <- rbind(results_table, weighted_row)
print(results_table_combined)



library(ggplot2)

results_table_numeric <- results_table %>%
  filter(Coverage_Percentage != "Weighted Strategy") %>%
  mutate(Coverage_Percentage = as.numeric(Coverage_Percentage))


ggplot(results_table_numeric, aes(x = Coverage_Percentage * 100, y = Answers_Selected)) +
  geom_line(color = "darkred", size = 1.2) +
  geom_point(size = 3, color = "darkred") +
  scale_x_continuous(breaks = seq(50, 100, 10)) +
  labs(
    title = "Trade-off Between Trait Coverage and Number of Answer Options",
    x = "Trait Coverage Target (%)",
    y = "Number of Answer Options Selected"
  ) +
  theme_minimal()

ggplot(results_table_numeric, aes(x = Coverage_Percentage * 100, y = Percent_Professions_Covered)) +
  geom_line(color = "blue", size = 1.2) +
  geom_point(size = 3, color = "blue") +
  scale_x_continuous(breaks = seq(50, 100, 10)) +
  labs(
    title = "Trait Coverage vs Professions Covered",
    x = "Trait Coverage Target (%)",
    y = "Percent of Professions Covered"
  ) +
  theme_minimal()

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

# Prepare the list of dataframes
data_list <- list(
  "Comparison Table" = comparison_table, 
  "Simulation Results" = results_table_combined
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "jeffsmith_coverage_analysis_tables.xlsx")

print(selected_weighted_answers)


# Order by numeric scenario and option
ordered_list <- selected_weighted_answers %>%
  tibble::tibble(answer_id = .) %>%
  dplyr::mutate(
    scenario = as.numeric(stringr::str_extract(answer_id, "\\d+")),
    option = stringr::str_extract(answer_id, "[a-d]")
  ) %>%
  dplyr::arrange(scenario, option) %>%
  dplyr::pull(answer_id)

print(ordered_list)
