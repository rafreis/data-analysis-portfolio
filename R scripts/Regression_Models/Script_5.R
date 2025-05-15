# ---- Data Loading ----
# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/johnelfers")

library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("ADC Grief STEs  Data.xlsx")


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

library(dplyr)
colnames(df)
glimpse(df)


# ---- Initial Setup ----
library(tidyverse)
library(effectsize)  # for power analysis
library(car)         # for ANOVA
library(broom)       # for tidy output
library(psych)       # optional for scoring or checking scale distributions

# Recode categorical variables as factors with meaningful labels
df <- df %>%
  mutate(
    SpiritPract = factor(SpiritPract, levels = c(1, 2, 3), labels = c("Yes", "Somewhat", "No")),
    ADC = factor(ADC, levels = c(1, 2), labels = c("Yes", "No")),
    OngoingConn = factor(OngoingConn, levels = c(1, 2, 3), labels = c("Mild", "Moderate", "Strong"))
  )

# ------------------------

# ---- Power Analysis for Correlation ----
library(pwr)

# Sample size
n <- nrow(df)

# Small-to-medium effect (r = 0.20)
pwr_small <- pwr.r.test(n = n, r = 0.20, sig.level = 0.05, alternative = "two.sided")

# Medium effect (r = 0.30)
pwr_medium <- pwr.r.test(n = n, r = 0.30, sig.level = 0.05, alternative = "two.sided")

# View results
pwr_small
pwr_medium

colnames(df)


# ---- ANOVA + Descriptive Stats (mean, SD, SEM) by Group ----

df <- df %>%
  mutate(
    Age = factor(
      Age,
      levels = 1:5,
      labels = c("18–24", "25–40", "41–55", "56–70", "70+"),
      ordered = TRUE
    ),
    Gender = factor(
      Gender,
      levels = c(1, 2, 3, 4, 5),
      labels = c("Man", "Nonbinary", "Woman", "Transgender", "NA")
    )
  )


dv_vars <- c("TOTAL.GMRI", "Total.Cont.B", "TotalTGS", "Total.Comp")

# Vector with race variable names
race_vars <- c("AmInd", "AsianPI", "Black", "Hispan", "White")

# Recode: 1 if not NA, 0 otherwise
df[race_vars] <- lapply(df[race_vars], function(x) ifelse(is.na(x), 0, 1))

group_vars <- c("ADC", "SpiritPract", "OngoingConn", "Gender", "Age", race_vars)


str(df)

anova_by <- function(data, dv, group) {
  df_clean <- data %>%
    select(all_of(c(dv, group))) %>%
    drop_na()
  
  # Levene's Test
  levene <- car::leveneTest(df_clean[[dv]] ~ as.factor(df_clean[[group]]))
  levene_F <- round(levene[1, "F value"], 3)
  levene_p <- round(levene[1, "Pr(>F)"], 3)
  
  # Verificar número mínimo de observações por grupo
  group_counts <- df_clean %>%
    count(.data[[group]]) %>%
    pull(n)
  
  has_min_per_group <- all(group_counts >= 2)
  
  # Escolher teste
  if (levene_p < 0.05 && has_min_per_group) {
    test <- oneway.test(df_clean[[dv]] ~ as.factor(df_clean[[group]]), var.equal = FALSE)
    F_val <- round(test$statistic[[1]], 3)
    p_val <- round(test$p.value, 3)
  } else {
    model <- aov(reformulate(group, dv), data = df_clean)
    anova_table <- summary(model)[[1]]
    F_val <- round(unname(anova_table[1, "F value"]), 3)
    p_val <- round(unname(anova_table[1, "Pr(>F)"]), 3)
  }
  
  # Descritivas
  desc <- df_clean %>%
    group_by(across(all_of(group))) %>%
    summarise(
      DV = dv,
      Group = group,
      Group_Level = as.character(first(.data[[group]])),
      N = n(),
      Mean = mean(.data[[dv]], na.rm = TRUE),
      SD = sd(.data[[dv]], na.rm = TRUE),
      SEM = SD / sqrt(N),
      .groups = "drop"
    ) %>%
    mutate(
      F = F_val,
      p = p_val,
      Levene_F = levene_F,
      Levene_p = levene_p
    ) %>%
    relocate(DV, Group, Group_Level, N, Mean, SD, SEM, F, p, Levene_F, Levene_p)
  
  return(desc)
}

df_anova_gender <- df %>%
  filter(Gender %in% c("Man", "Woman")) %>%
  droplevels()

anova_results <- cross_df(list(DV = dv_vars, Group = group_vars)) %>%
  mutate(
    results = map2(DV, Group, function(dv, grp) {
      df_input <- if (grp == "Gender") {
        df %>% filter(Gender %in% c("Man", "Woman")) %>% droplevels()
      } else {
        df
      }
      anova_by(df_input, dv, grp)
    })
  ) %>%
  unnest(results, names_repair = "unique")



library(ggplot2)
library(tidyr)

# ---- Prepare summary table with all groups ----

prepare_plot_data <- function(data, dv, group_vars) {
  map_dfr(group_vars, function(g) {
    data %>%
      select(dv = all_of(dv), group = all_of(g)) %>%
      drop_na() %>%
      group_by(group_level = as.character(group)) %>%
      summarise(
        DV = dv,
        Grouping = g,
        Group_Level = group_level,
        Mean = mean(dv, na.rm = TRUE),
        SEM = sd(dv, na.rm = TRUE) / sqrt(n()),
        .groups = "drop"
      )
  })
}

# ---- Generate and save one plot per DV ----

library(forcats)

for (dv in dv_vars) {
  plot_data <- prepare_plot_data(df, dv, group_vars)
  
  # Generate ordered x-axis by grouping variable and level
  plot_data <- plot_data %>%
    mutate(
      x_group = paste0(Grouping, ": ", Group_Level),
      Grouping = factor(Grouping, levels = group_vars),
      x_group = fct_inorder(x_group)
    ) %>%
    arrange(Grouping, x_group)
  
  p <- ggplot(plot_data, aes(x = x_group, y = Mean, color = Grouping)) +
    geom_point(size = 3) +
    geom_errorbar(aes(ymin = Mean - SEM, ymax = Mean + SEM), width = 0.2) +
    labs(
      title = paste0("Means ± SEM for ", dv, " across groups"),
      x = "Group",
      y = "Mean",
      color = "Grouping Variable"
    ) +
    theme_minimal() +
    theme(
      axis.text.x = element_text(angle = 45, hjust = 1),
      legend.position = "none"  # <- remove legenda
    )
  print(p)
  ggsave(paste0("plot_", dv, "_all_groups.png"), plot = p, width = 8, height = 5, dpi = 300)
}



# ---- Custom OLS Function ----

ols_by <- function(data, dv, formula_rhs) {
  # Build formula
  fmla <- as.formula(paste(dv, "~", formula_rhs))
  
  # Drop missing
  df_clean <- data %>%
    select(all_of(c(dv, all.vars(fmla[[3]])))) %>%
    drop_na()
  
  # Fit model
  model <- lm(fmla, data = df_clean)
  summary_model <- summary(model)
  
  # Extract coefficients
  out <- broom::tidy(model) %>%
    mutate(
      DV = dv,
      R2 = round(summary_model$r.squared, 3),
      Adj_R2 = round(summary_model$adj.r.squared, 3)
    ) %>%
    select(DV, term, estimate, std.error, statistic, p.value, R2, Adj_R2) %>%
    mutate(across(where(is.numeric), ~ round(.x, 3)))
  
  return(out)
}

dv_vars <- c("TOTAL.GMRI", "Total.Cont.B", "TotalTGS", "Total.Comp")
predictors <- c("ADC", "SpiritPract", "OngoingConn", "Gender", "Age", race_vars)
fmla <- paste(predictors, collapse = " + ")

df$Age <- factor(df$Age, levels = sort(unique(df$Age)), ordered = FALSE)

df$Gender <- factor(df$Gender, levels = sort(unique(df$Gender)), ordered = FALSE)

regression_results <- map_dfr(dv_vars, ~ ols_by(df, dv = .x, formula_rhs = fmla))

library(performance)

# Plot and save diagnostic panels
for (dv in dv_vars) {
  fmla <- as.formula(paste(dv, "~", fmla_rhs))
  model <- lm(fmla, data = df)
  
  png(filename = paste0("diagnostic_panel_", dv, ".png"), width = 1200, height = 1500, res = 150)
  print(check_model(model))
  dev.off()
}

df$log_Total.Comp <- log(df$Total.Comp)

# Fórmula com a variável transformada
fmla_log <- as.formula(paste("log_Total.Comp ~", fmla_rhs))

# Ajustar o modelo
model_log <- lm(fmla_log, data = df)

# Salvar diagnóstico
png("diagnostic_panel_log_Total.Comp.png", width = 1200, height = 1000, res = 150)
print(performance::check_model(model_log))
dev.off()


ols_log_comp <- ols_by(df, dv = "log_Total.Comp", formula_rhs = fmla)


# Recode race from mutually exclusive dummies into a single categorical variable
df$Race <- case_when(
  df$AmInd == 1   ~ "American Indian",
  df$AsianPI == 1 ~ "Asian/Pacific Islander",
  df$Black == 1   ~ "Black",
  df$Hispan == 1  ~ "Hispanic",
  df$White == 1   ~ "White",
  TRUE            ~ "Unspecified"
)

# Opcional: converter para fator com ordem
df$Race <- factor(df$Race, levels = c("White", "Black", "Asian/Pacific Islander", "Hispanic", "American Indian", "Unspecified"))



create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each category variable
  for (category in categories) {
    # Ensure the category variable is a factor
    data[[category]] <- factor(data[[category]])
    
    # Calculate counts
    counts <- table(data[[category]])
    
    # Create a dataframe for this category
    freq_table <- data.frame(
      "Category" = rep(category, length(counts)),
      "Level" = names(counts),
      "Count" = as.integer(counts),
      stringsAsFactors = FALSE
    )
    
    # Calculate and add percentages
    freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
    
    # Add the result to the list
    all_freq_tables[[category]] <- freq_table
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

predictors <- c("ADC", "SpiritPract", "OngoingConn", "Gender", "Age", "Race")

df_freq <- create_frequency_tables(df, predictors)




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

# Example usage with your dataframe and list of variables
df_descriptive_stats <- calculate_descriptive_stats(df, dv_vars)




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
  "Data" = df,
  "Frequency Table" = df_freq, 
  "Descriptives" = df_descriptive_stats,
  "ANOVA Results" = anova_results,
  "Regression" = regression_results,
  "Adj log Regression" = ols_log_comp
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")



