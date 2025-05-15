# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dorint2009")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("CleanData.xlsx")

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

colnames(df)

library(dplyr)

var_Trial <- ("Field")
var_Replication <- ("Replicates")

# Rename the existing 'Hybrid' variable to 'Hybrid_old'
df <- df %>% rename(Hybrid_old = Hybrid)

# Recode the 'Hybrid_old' variable into a new 'Hybrid' variable
df <- df %>% 
  mutate(Hybrid = case_when(
    tolower(Hybrid_old) %in% c("dkc6980", "dkc6808", "dkc6625", "dkc6503") ~ "Tall Hybrids(75cm)",
    tolower(Hybrid_old) %in% c("dkc6836", "dkc7236", "ex6658", "et6641") ~ "SSC Hybrids(75cm)",
    tolower(Hybrid_old) %in% c("dkc6836n", "dkc7236n", "ex6658n", "et6641n") ~ "SSC Hybrids 40cm",
    TRUE ~ Hybrid_old  # Retain original value if no match is found
  ))

var_factors <- c("Hybrid", "Irrigation")
var_ID <- "Entry.list.2way"

vars_1wayANOVA <- c("uniformity" , "Early.vigor" , "Average.Plant.Height" ,"Average.Ear.Height")

vars_2wayANOVA <- c("Stalk.ECB" , "Fusarium", "Stay.Green"  ,  "Harvesyed.plants"  ,   "Kg_str_15.")
  
# Renaming the columns
colnames(df) <- gsub("^gS1$", "gs1", colnames(df))
colnames(df) <- gsub("^gS2$", "gs2", colnames(df))
colnames(df) <- gsub("^gS3$", "gs3", colnames(df))
colnames(df) <- gsub("^Ε", "E", colnames(df))  # Replace any Greek 'E' (capital epsilon) with 'E'
colnames(df) <- gsub("^Α", "A", colnames(df))  # Replace any Greek 'A' (capital alpha) with 'A'
colnames(df) <- gsub("^gs", "gs", colnames(df), ignore.case = TRUE)  # Ensure all 'gs' are lowercase

colnames(df) <- trimws(colnames(df))


vars_firstmeas <- c("E1", "gs1", "A1")
vars_allmeas <- c(   "E2", "gs2", "A2" ,                                               
  "E3"   ,                                      "gs3"   ,                                     "A3"  ,                                      
  "E4"    ,                                     "gs4"     ,                                   "A4"   ,                                     
  "E5"     ,                                    "gs5"      ,                                  "A5"    ,                                    
  "E6", "gs6", "A6"    )


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

# Example usage with your data
df_normality_results <- calculate_stats(df, c(vars_1wayANOVA, vars_2wayANOVA, vars_firstmeas, vars_allmeas))



# Load necessary packages
library(dplyr)
library(broom)
library(agricolae)
library(tibble)

# Function to run one-way ANOVA with Duncan posthoc test
run_1way_ANOVA <- function(df, dependent_vars, independent_var) {
  results_list <- lapply(dependent_vars, function(dv) {
    # Perform one-way ANOVA
    formula <- as.formula(paste(dv, "~", independent_var))
    anova_result <- aov(formula, data = df)
    
    # Get tidy ANOVA summary
    anova_tidy <- tidy(anova_result)
    
    # Perform Duncan posthoc test
    duncan_test <- duncan.test(anova_result, independent_var)
    duncan_tidy <- as.data.frame(duncan_test$groups)
    duncan_tidy <- duncan_tidy %>%
      mutate(variable = dv, .before = 1) %>%
      rownames_to_column(var = "group")
    
    # Return the results as a list
    list(
      anova = anova_tidy %>% mutate(variable = dv),
      duncan = duncan_tidy
    )
  })
  
  # Combine all results into a tidy dataframe
  anova_df <- bind_rows(lapply(results_list, `[[`, "anova"))
  duncan_df <- bind_rows(lapply(results_list, `[[`, "duncan"))
  
  list(anova = anova_df, duncan = duncan_df)
}

# Run the one-way ANOVA for the specified variables
results_1way <- run_1way_ANOVA(df, vars_1wayANOVA, "Hybrid")
anova_1way_results <- results_1way$anova
duncan_1way_results <- results_1way$duncan

library(broom)

# Function to run two-way ANOVA for multiple dependent variables and Duncan post-hoc test
run_2way_ANOVA <- function(df, dependent_vars, independent_vars) {
  results_list <- lapply(dependent_vars, function(dv) {
    # Create formula for two-way ANOVA
    formula <- as.formula(paste(dv, "~", paste(independent_vars, collapse = " * ")))
    anova_result <- aov(formula, data = df)
    
    # Get tidy ANOVA summary
    anova_tidy <- tidy(anova_result) %>%
      mutate(variable = dv)
    
    # Perform Duncan post-hoc test for all independent variables
    duncan_results <- lapply(independent_vars, function(iv) {
      tryCatch({
        duncan_test <- duncan.test(anova_result, iv)  # Run Duncan's test for each IV
        duncan_tidy <- as.data.frame(duncan_test$groups) %>%
          rownames_to_column(var = "group") %>%
          mutate(variable = dv, independent_var = iv, test = "Duncan")  # Add identifier for Duncan test
      }, error = function(e) {
        # If Duncan test fails, return NA for this variable
        data.frame(group = NA, means = NA, variable = dv, independent_var = iv, test = "Duncan")
      })
    })
    
    # Combine all Duncan results into a single dataframe
    duncan_combined <- bind_rows(duncan_results)
    
    # Return a list containing both ANOVA and Duncan results
    list(anova = anova_tidy, duncan = duncan_combined)
  })
  
  # Combine all results into tidy dataframes
  anova_df <- bind_rows(lapply(results_list, `[[`, "anova"))
  duncan_df <- bind_rows(lapply(results_list, `[[`, "duncan"))
  
  # Return both ANOVA and Duncan's test results
  return(list(anova_results = anova_df, duncan_results = duncan_df))
}


# Run the two-way ANOVA for the specified variables
anova_2way <- run_2way_ANOVA(df, vars_2wayANOVA, var_factors)
anova_2way_results <- anova_2way$anova
duncan_2way_results <- anova_2way$duncan



# Import all measure files

# Load necessary libraries
library(readxl)
library(dplyr)

# Set the path to the folder containing the Excel files
folder_path <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dorint2009/All Measures"

# List all Excel files in the folder
file_list <- list.files(path = folder_path, pattern = "*.xlsx", full.names = TRUE)

# Initialize an empty list to store the dataframes
df_list <- list()

# Loop through each file
for (file in file_list) {
  # Get the file name without the full path for comparison
  file_name <- basename(file)
  
  # Condition: Do not skip rows for 'SSC1_1st', 'SSC2_1st', 'SSC3_1st'
  if (file_name %in% c("SSC1_1st.xlsx", "SSC2_1st.xlsx", "SSC3_1st.xlsx")) {
    # Read the file without skipping rows
    df2 <- read_excel(file)
  } else {
    # Read the file, skipping the first two rows, treating the third as headers
    df2 <- read_excel(file, skip = 2)
  }
  
  # Select the first three columns by position and columns named E, gs, and A
  df2 <- df2 %>%
    select(1:3, 
           matches("^E$", ignore.case = TRUE), 
           matches("^gs$", ignore.case = TRUE), 
           matches("^A$", ignore.case = TRUE))
  
  # Add the dataframe to the list
  df_list[[file]] <- df2
}

# Now, standardize data types outside the main function
df_list <- lapply(df_list, function(df2) {
  df2 %>%
    mutate(across(everything(), as.character))  # Convert all columns to character
})

# Combine all dataframes into one
final_df <- bind_rows(df_list, .id = "source_file")


# Assuming you have already read the data into `final_df`
# Now let's join the columns

final_df <- final_df %>%
  # Combine columns into 'Entry'
  mutate(Entry = coalesce(...3, entry, Entry, enrty)) %>%
  
  # Combine columns into 'Tier Row'
  mutate(`Tier Row` = coalesce(`...2`, `tier roe`, `tier row`, `tier/row`, `Tier/row`)) %>%
  
  # Remove the original columns that were combined
  select(-c(`...3`, entry, enrty, `...2`, `tier roe`, `tier row`, `tier/row`, `Tier/row`))


final_df <- final_df %>%
  mutate(source_file = gsub("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dorint2009/All Measures/", "", source_file)) %>%
  mutate(source_file = gsub(".xlsx", "", source_file))


# Convert E, gs, and A columns to numeric, with errors replaced by NA
final_df <- final_df %>%
  mutate(across(c(E, gs, A), ~ as.numeric(.)))

str(final_df)

# Evaluate Extreme values

# Create ID column

final_df <- cbind(ID = 1:nrow(final_df), final_df)


calculate_z_scores <- function(data, vars, id_var, z_threshold = 3) {
  z_score_data <- data %>%
    # Calculate z-scores, handling NAs by ignoring them in the mean and sd calculations
    mutate(across(all_of(vars), ~ (.-mean(., na.rm = TRUE))/sd(., na.rm = TRUE), .names = "z_{.col}")) %>%
    # Select original values, z-scores, and the ID column
    select(!!sym(id_var), all_of(vars), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(is.na(.), NA, ifelse(abs(.) > z_threshold, "Outlier", "Normal")), .names = "flag_{.col}"))
  
  return(z_score_data)
}

id_var <- "ID"
vars_DVs <- c("E", "gs", "A")
df_zscores <- calculate_z_scores(final_df, vars_DVs, "ID")

library(ggplot2)
library(tidyr)
library(stringr)

create_boxplots <- function(df, vars) {
  # Loop through each variable in vars and create a boxplot for each
  for (var in vars) {
    # Create a separate plot for each variable
    p <- ggplot(df, aes_string(x = "factor(1)", y = var)) +
      geom_boxplot() +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
      labs(title = paste("Boxplot of", var), x = var, y = "Value") +
      theme(axis.text.x = element_blank())  # Remove x-axis labels since they aren't needed
    
    # Print the plot
    print(p)
    
    # Save each plot as a separate file
    ggsave(filename = paste0("boxplot_", var, ".png"), plot = p, width = 6, height = 4)
  }
}

create_boxplots(final_df, vars_DVs)

df <- cbind(ID = 1:nrow(df), df)
id_var <- "ID"
df_zscores_total <- calculate_z_scores(df, c(vars_1wayANOVA, vars_2wayANOVA, vars_firstmeas, vars_allmeas), "ID")


colnames(df)

# Run the one-way ANOVA for the specified variables
results_1way_EAgs1 <- run_1way_ANOVA(df, vars_firstmeas, "Hybrid")
anova_1way_results_EAgs1 <- results_1way_EAgs1$anova
duncan_1way_results_EAgs1 <- results_1way_EAgs1$duncan

# Run the two-way ANOVA for the specified variables
anova_2way_results <- run_2way_ANOVA(df, vars_allmeas, var_factors)
anova_2way_results_EAgs <- anova_2way_results$anova
duncan_2way_results_EAgs <- anova_2way_results$duncan



colnames(df)
str(df)

# Function to run repeated measures ANOVA for E, gs, and A across time points

long_df <- df %>%
  mutate(ID = row_number()) %>%  # Add an ID column corresponding to each row
  pivot_longer(cols = all_of(c(vars_firstmeas, vars_allmeas)),
               names_to = c("Measure", "Time"),
               names_pattern = "(E|gs|A)(\\d)") %>%
  mutate(Time = factor(Time)) %>%  # Ensure 'Time' is a factor
  select(ID, Measure, Time, value)  # Keep only relevant columns (ID, Measure, Time, Value)



# Function to run repeated measures ANOVA for E, gs, and A across time points
run_repeated_measures_ANOVA <- function(df, measures) {
  
  # Remove rows with missing data for the current measure
  long_df <- df %>%
    drop_na()  # This removes any rows with missing values
  
  # Initialize an empty list to store ANOVA results
  anova_results_list <- list()
  
  # Loop through each measure (E, gs, A) and perform repeated measures ANOVA
  for (measure in measures) {
    # Filter the data for the current measure
    measure_data <- long_df %>% filter(Measure == measure)
    
    # Create the formula for repeated measures ANOVA
    formula <- as.formula("value ~ Time + Error(ID/Time)")
    
    # Run the ANOVA
    anova_result <- aov(formula, data = measure_data)
    
    # Get tidy ANOVA results
    anova_tidy <- tidy(anova_result) %>%
      mutate(measure = measure)
    
    # Append results to the list
    anova_results_list[[measure]] <- anova_tidy
  }
  
  # Combine all results into a single dataframe
  anova_df <- bind_rows(anova_results_list)
  
  return(anova_df)
}

# Measures to analyze
measures <- c("E", "gs", "A")

# Run the repeated measures ANOVA
anova_repeated_results <- run_repeated_measures_ANOVA(long_df, measures)

# View the results
View(anova_repeated_results)


# Load required libraries
library(ggplot2)
library(dplyr)

# Function to create line plots for the average value of each measure across IDs
create_avg_line_plots <- function(df, measures) {
  # Loop through each measure and create a line plot of the average values
  for (measure in measures) {
    # Filter the data for the current measure
    measure_data <- df %>%
      filter(Measure == measure) %>%
      group_by(Time) %>%             # Group by time points
      summarise(mean_value = mean(value, na.rm = TRUE))  # Calculate mean for each time point
    
    # Create the line plot for the average values
    p <- ggplot(measure_data, aes(x = Time, y = mean_value, group = 1)) +
      geom_line(color = "blue", size = 1) +  # Add line for the average
      geom_point(size = 3) +                 # Add points for each time point
      geom_text(aes(label = sprintf("%.2f", mean_value)),  # Add data labels with 2 decimal points
                vjust = -0.5, size = 3.5) +  # Adjust vertical position of the labels
      
      labs(title = paste("Average Line Plot for", measure),
           x = "Time",
           y = paste("Average Value of", measure)) +
      theme_minimal()
    
    # Print the plot
    print(p)
  }
}

# Measures to analyze
measures <- c("E", "gs", "A")

# Run the line plot function to show average values with data labels
create_avg_line_plots(long_df, measures)




# Results with log-transformed variables

# 1. Log-transform the dependent variables (adding small value to avoid log(0))
log_transform_vars <- function(df, vars) {
  df <- df %>%
    mutate(across(all_of(vars), ~ log(. + 1)))  # Log-transform (adding 1 to avoid log(0))
  return(df)
}

# Log-transforming the dependent variables for one-way ANOVA
vars_1wayANOVA_log <- paste0("log_", vars_1wayANOVA)
df <- df %>%
  mutate(across(all_of(vars_1wayANOVA), ~ log(. + 1), .names = "log_{.col}"))

# Log-transforming the dependent variables for two-way ANOVA
vars_2wayANOVA_log <- paste0("log_", vars_2wayANOVA)
df <- df %>%
  mutate(across(all_of(vars_2wayANOVA), ~ log(. + 1), .names = "log_{.col}"))

# Log-transforming the repeated measures variables (E, gs, A)
vars_firstmeas_log <- paste0("log_", vars_firstmeas)
vars_allmeas_log <- paste0("log_", vars_allmeas)
df <- df %>%
  mutate(across(all_of(vars_firstmeas), ~ log(. + 1), .names = "log_{.col}")) %>%
  mutate(across(all_of(vars_allmeas), ~ log(. + 1), .names = "log_{.col}"))

# 2. Re-run the one-way ANOVA with Duncan post-hoc on log-transformed variables
results_1way_log <- run_1way_ANOVA(df, vars_1wayANOVA_log, "Hybrid")
anova_1way_results_log <- results_1way_log$anova
duncan_1way_results_log <- results_1way_log$duncan

# 3. Re-run the two-way ANOVA with Duncan post-hoc on log-transformed variables
anova_2way_log <- run_2way_ANOVA(df, vars_2wayANOVA_log, var_factors)
anova_2way_results_log <- anova_2way_log$anova
duncan_2way_results_log <- anova_2way_log$duncan

# 4. Prepare the long dataframe for repeated measures ANOVA on log-transformed variables
long_df_log <- df %>%
  mutate(ID = row_number()) %>%  # Add an ID column corresponding to each row
  pivot_longer(cols = all_of(c(vars_firstmeas_log, vars_allmeas_log)),
               names_to = c("Measure", "Time"),
               names_pattern = "log_(E|gs|A)(\\d)") %>%
  mutate(Time = factor(Time)) %>%
  select(ID, Measure, Time, value)  # Keep only relevant columns (ID, Measure, Time, Value)

# 5. Re-run the repeated measures ANOVA on log-transformed variables
anova_repeated_results_log <- run_repeated_measures_ANOVA(long_df_log, c("E", "gs", "A"))

# View the results for one-way, two-way, and repeated measures ANOVA
View(anova_1way_results_log)
View(duncan_1way_results_log)
View(anova_2way_results_log)
View(duncan_2way_results_log)
View(anova_repeated_results_log)



# 1. Log-transform the first measurement variables (E1, gs1, A1)
vars_firstmeas_log <- paste0("log_", vars_firstmeas)
df <- df %>%
  mutate(across(all_of(vars_firstmeas), ~ log(. + 1), .names = "log_{.col}"))

# 2. Log-transform the remaining measurement variables (E2, gs2, A2, ... E6, gs6, A6)
vars_allmeas_log <- paste0("log_", vars_allmeas)
df <- df %>%
  mutate(across(all_of(vars_allmeas), ~ log(. + 1), .names = "log_{.col}"))

# 3. Re-run the one-way ANOVA with Duncan post-hoc on log-transformed first measurement variables
results_1way_log_EAgs1 <- run_1way_ANOVA(df, vars_firstmeas_log, "Hybrid")
anova_1way_results_log_EAgs1 <- results_1way_log_EAgs1$anova
duncan_1way_results_log_EAgs1 <- results_1way_log_EAgs1$duncan

# 4. Re-run the two-way ANOVA with Duncan post-hoc on log-transformed remaining measurement variables
results_2way_log_EAgs <- run_2way_ANOVA(df, vars_allmeas_log, var_factors)
anova_2way_results_log_EAgs <- results_2way_log_EAgs$anova
duncan_2way_results_log_EAgs <- results_2way_log_EAgs$duncan

# View the results
View(anova_1way_results_log_EAgs1)
View(duncan_1way_results_log_EAgs1)
View(anova_2way_results_log_EAgs)
View(duncan_2way_results_log_EAgs)



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

# Create a list with all relevant data frames
data_list <- list(
  "anova_1way_results" = anova_1way_results,
  "anova_1way_results_EAgs1" = anova_1way_results_EAgs1,

  
  "anova_2way_results" = anova_2way_results,
  "anova_2way_results_EAgs" = anova_2way_results_EAgs,

  
  "duncan_1way_results" = duncan_1way_results,
  "duncan_1way_results_EAgs1" = duncan_1way_results_EAgs1,

  
  "duncan_2way_results" = duncan_2way_results,
  "duncan_2way_results_EAgs" = duncan_2way_results_EAgs,

  
  "anova_repeated_results" = anova_repeated_results,

  
  "df_zscores" = df_zscores,
  "df_zscores_total" = df_zscores_total,
  "df_normality_results" = df_normality_results,
  "final_df" = final_df
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

data_list <- list(

  "anova_1way_results_log" = anova_1way_results_log,
  "anova_1way_results_log_EAgs1" = anova_1way_results_log_EAgs1,

  "anova_2way_results_log" = anova_2way_results_log,
  "anova_2way_results_log_EAgs" = anova_2way_results_log_EAgs,

  "duncan_1way_results_log" = duncan_1way_results_log,
  "duncan_1way_results_log_EAgs1" = duncan_1way_results_log_EAgs1,
  

  "duncan_2way_results_log" = duncan_2way_results_log,
  "duncan_2way_results_log_EAgs" = duncan_2way_results_log_EAgs,
  

  "anova_repeated_results_log" = anova_repeated_results_log

)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_logtransformed.xlsx")


data_list <- list(
  
  "Data" = df
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "CompleteData.xlsx")








# Additional Duncans significance

run_2way_ANOVA <- function(df, dependent_vars, independent_vars) {
  results_list <- lapply(dependent_vars, function(dv) {
    # Create formula for two-way ANOVA
    formula <- as.formula(paste(dv, "~", paste(independent_vars, collapse = " * ")))
    anova_result <- aov(formula, data = df)
    
    # Get tidy ANOVA summary
    anova_tidy <- tidy(anova_result) %>%
      mutate(variable = dv)
    
    # Perform Duncan post-hoc test for all independent variables
    duncan_results <- lapply(independent_vars, function(iv) {
      tryCatch({
        duncan_test <- duncan.test(anova_result, iv)  # Run Duncan's test for each IV
        # Duncan test provides grouping information
        duncan_tidy <- as.data.frame(duncan_test$groups) %>%
          rownames_to_column(var = "group") %>%
          mutate(
            variable = dv,
            independent_var = iv,
            test = "Duncan",
            # Adding the p-value flags for p < 0.10 and p < 0.15 based on the grouping
            p_lt_0.10 = ifelse(grepl("[ab]", groups), TRUE, FALSE),  # This flag could be more specific
            p_lt_0.15 = ifelse(grepl("[ab]", groups), TRUE, FALSE)   # Modify flag as per interpretation of groups
          )
      }, error = function(e) {
        # If Duncan test fails, return NA for this variable
        data.frame(group = NA, means = NA, variable = dv, independent_var = iv, test = "Duncan", p_lt_0.10 = NA, p_lt_0.15 = NA)
      })
    })
    
    # Combine all Duncan results into a single dataframe
    duncan_combined <- bind_rows(duncan_results)
    
    # Return a list containing both ANOVA and Duncan results
    list(anova = anova_tidy, duncan = duncan_combined)
  })
  
  # Combine all results into tidy dataframes
  anova_df <- bind_rows(lapply(results_list, `[[`, "anova"))
  duncan_df <- bind_rows(lapply(results_list, `[[`, "duncan"))
  
  # Return both ANOVA and Duncan's test results
  return(list(anova_results = anova_df, duncan_results = duncan_df))
}

# Run the two-way ANOVA for the specified variables
anova_2way <- run_2way_ANOVA(df, vars_allmeas, var_factors)
anova_2way_results_allmeas <- anova_2way$anova
duncan_2way_results_allmeas <- anova_2way$duncan


# Update the Hybrid factor
var_factors <- c("Hybrid_old", "Irrigation")

# Function to run two-way ANOVA for multiple dependent variables and Duncan post-hoc test
run_2way_ANOVA <- function(df, dependent_vars, independent_vars) {
  results_list <- lapply(dependent_vars, function(dv) {
    # Create formula for two-way ANOVA
    formula <- as.formula(paste(dv, "~", paste(independent_vars, collapse = " * ")))
    anova_result <- aov(formula, data = df)
    
    # Get tidy ANOVA summary
    anova_tidy <- tidy(anova_result) %>%
      mutate(variable = dv)
    
    # Perform Duncan post-hoc test for all independent variables
    duncan_results <- lapply(independent_vars, function(iv) {
      tryCatch({
        # Duncan test with alpha = 0.05 (default)
        duncan_test_05 <- duncan.test(anova_result, iv, alpha = 0.05)
        duncan_tidy_05 <- as.data.frame(duncan_test_05$groups) %>%
          rownames_to_column(var = "group") %>%
          mutate(groups_05 = duncan_test_05$groups$groups)
        
        # Duncan test with alpha = 0.10
        duncan_test_10 <- duncan.test(anova_result, iv, alpha = 0.10)
        duncan_tidy_10 <- as.data.frame(duncan_test_10$groups) %>%
          rownames_to_column(var = "group") %>%
          mutate(groups_10 = duncan_test_10$groups$groups)
        
        # Duncan test with alpha = 0.15
        duncan_test_15 <- duncan.test(anova_result, iv, alpha = 0.15)
        duncan_tidy_15 <- as.data.frame(duncan_test_15$groups) %>%
          rownames_to_column(var = "group") %>%
          mutate(groups_15 = duncan_test_15$groups$groups)
        
        # Combine all the results into one dataframe using group as the common column
        duncan_combined <- duncan_tidy_05 %>%
          left_join(duncan_tidy_10 %>% select(group, groups_10), by = "group") %>%
          left_join(duncan_tidy_15 %>% select(group, groups_15), by = "group") %>%
          mutate(variable = dv, independent_var = iv, test = "Duncan")
        
        return(duncan_combined)
      }, error = function(e) {
        # If Duncan test fails, return NA for this variable
        data.frame(group = NA, means = NA, groups_05 = NA, groups_10 = NA, groups_15 = NA, variable = dv, independent_var = iv, test = "Duncan")
      })
    })
    
    # Combine all Duncan results into a single dataframe
    duncan_combined <- bind_rows(duncan_results)
    
    # Return a list containing both ANOVA and Duncan results
    list(anova = anova_tidy, duncan = duncan_combined)
  })
  
  # Combine all results into tidy dataframes
  anova_df <- bind_rows(lapply(results_list, `[[`, "anova"))
  duncan_df <- bind_rows(lapply(results_list, `[[`, "duncan"))
  
  # Return both ANOVA and Duncan's test results
  return(list(anova_results = anova_df, duncan_results = duncan_df))
}


run_1way_ANOVA <- function(df, dependent_vars, independent_var) {
  results_list <- lapply(dependent_vars, function(dv) {
    # Perform one-way ANOVA
    formula <- as.formula(paste(dv, "~", independent_var))
    anova_result <- aov(formula, data = df)
    
    # Get tidy ANOVA summary
    anova_tidy <- tidy(anova_result)
    
    # Perform Duncan posthoc test with alpha = 0.05 (default)
    duncan_test_05 <- duncan.test(anova_result, independent_var, alpha = 0.05)
    duncan_tidy_05 <- as.data.frame(duncan_test_05$groups) %>%
      rownames_to_column(var = "group") %>%
      mutate(groups_05 = duncan_test_05$groups$groups)
    
    # Perform Duncan posthoc test with alpha = 0.10
    duncan_test_10 <- duncan.test(anova_result, independent_var, alpha = 0.10)
    duncan_tidy_10 <- as.data.frame(duncan_test_10$groups) %>%
      rownames_to_column(var = "group") %>%
      mutate(groups_10 = duncan_test_10$groups$groups)
    
    # Perform Duncan posthoc test with alpha = 0.15
    duncan_test_15 <- duncan.test(anova_result, independent_var, alpha = 0.15)
    duncan_tidy_15 <- as.data.frame(duncan_test_15$groups) %>%
      rownames_to_column(var = "group") %>%
      mutate(groups_15 = duncan_test_15$groups$groups)
    
    # Combine all the results into one dataframe using group as the common column
    duncan_combined <- duncan_tidy_05 %>%
      left_join(duncan_tidy_10 %>% select(group, groups_10), by = "group") %>%
      left_join(duncan_tidy_15 %>% select(group, groups_15), by = "group") %>%
      mutate(variable = dv, independent_var = independent_var, test = "Duncan")
    
    # Return the results as a list
    list(
      anova = anova_tidy %>% mutate(variable = dv),
      duncan = duncan_combined
    )
  })
  
  # Combine all results into a tidy dataframe
  anova_df <- bind_rows(lapply(results_list, `[[`, "anova"))
  duncan_df <- bind_rows(lapply(results_list, `[[`, "duncan"))
  
  # Return the results
  list(anova = anova_df, duncan = duncan_df)
}

# 2. Re-run the one-way ANOVA with Duncan post-hoc on log-transformed variables
results_1way_hybrids <- run_1way_ANOVA(df, vars_1wayANOVA, "Hybrid_old")
anova_1way_results_hybrids <- results_1way_hybrids$anova
duncan_1way_results_hybrids <- results_1way_hybrids$duncan

# 3. Re-run the two-way ANOVA with Duncan post-hoc on log-transformed variables
anova_2way_hybrids <- run_2way_ANOVA(df, vars_2wayANOVA, var_factors)
anova_2way_results_hybrids <- anova_2way_hybrids$anova
duncan_2way_results_hybrids <- anova_2way_hybrids$duncan

# 3. Re-run the one-way ANOVA with Duncan post-hoc on log-transformed first measurement variables
results_1way_EAgs1_hybrids <- run_1way_ANOVA(df, vars_firstmeas, "Hybrid_old")
anova_1way_results_EAgs1_hybrids <- results_1way_EAgs1_hybrids$anova
duncan_1way_results_EAgs1_hybrids <- results_1way_EAgs1_hybrids$duncan

# 4. Re-run the two-way ANOVA with Duncan post-hoc on log-transformed remaining measurement variables
results_2way_EAgs_hybrids <- run_2way_ANOVA(df, vars_allmeas, var_factors)
anova_2way_results_EAgs_hybrids <- results_2way_EAgs_hybrids$anova
duncan_2way_results_EAgs_hybrids <- results_2way_EAgs_hybrids$duncan


data_list <- list(
  "anova_1way_results" = anova_1way_results_hybrids,
  "anova_1way_results_EAgs1" = anova_1way_results_EAgs1_hybrids,
  "anova_2way_results" = anova_2way_results_hybrids,
  "anova_2way_results_EAgs" = anova_2way_results_EAgs_hybrids,
  "duncan_1way_results" = duncan_1way_results_hybrids,
  "duncan_1way_results_EAgs" = duncan_1way_results_EAgs1_hybrids,
  "duncan_2way_results" = duncan_2way_results_hybrids,
  "duncan_2way_results_EAgs" = duncan_2way_results_EAgs_hybrids
)

#save_apa_formatted_excel(data_list, "APA_Formatted_Tables_Hybrid2.xlsx")














library(dplyr)

# Function to run all the analyses for a given subset of the data
run_analyses_for_field <- function(subset_data) {
  results_1way_hybrids <- run_1way_ANOVA(subset_data, vars_1wayANOVA, "Hybrid_old")
  anova_1way_results_hybrids <- results_1way_hybrids$anova
  duncan_1way_results_hybrids <- results_1way_hybrids$duncan
  
  anova_2way_hybrids <- run_2way_ANOVA(subset_data, vars_2wayANOVA, var_factors)
  anova_2way_results_hybrids <- anova_2way_hybrids$anova
  duncan_2way_results_hybrids <- anova_2way_hybrids$duncan
  
  results_1way_EAgs1_hybrids <- run_1way_ANOVA(subset_data, vars_firstmeas, "Hybrid_old")
  anova_1way_results_EAgs1_hybrids <- results_1way_EAgs1_hybrids$anova
  duncan_1way_results_EAgs1_hybrids <- results_1way_EAgs1_hybrids$duncan
  
  results_2way_EAgs_hybrids <- run_2way_ANOVA(subset_data, vars_allmeas, var_factors)
  anova_2way_results_EAgs_hybrids <- results_2way_EAgs_hybrids$anova
  duncan_2way_results_EAgs_hybrids <- results_2way_EAgs_hybrids$duncan
  
  # Return the results as a list
  return(list(
    "anova_1way_results" = anova_1way_results_hybrids,
    "anova_1way_results_EAgs1" = anova_1way_results_EAgs1_hybrids,
    "anova_2way_results" = anova_2way_results_hybrids,
    "anova_2way_results_EAgs" = anova_2way_results_EAgs_hybrids,
    "duncan_1way_results" = duncan_1way_results_hybrids,
    "duncan_1way_results_EAgs1" = duncan_1way_results_EAgs1_hybrids,
    "duncan_2way_results" = duncan_2way_results_hybrids,
    "duncan_2way_results_EAgs" = duncan_2way_results_EAgs_hybrids
  ))
}

# Iterate over each subset of data by 'Field'
all_results <- df %>%
  group_by(Field) %>%
  group_map(~ run_analyses_for_field(.x))

# Save each subset's results into separate sheets in an Excel file
write_excel <- function(results_list, field_name, file_name) {
  for (i in seq_along(results_list)) {
    save_apa_formatted_excel(results_list[[i]], paste0(file_name, "_", field_name[i], ".xlsx"))
  }
}

# Generate list of field names
field_names <- unique(df$Field)

# Save results for all fields
write_excel(all_results, field_names, "APA_Formatted_Tables_byTrial")

library(dplyr)



long_df <- df %>%
  mutate(ID = row_number()) %>%  # Add an ID column corresponding to each row
  pivot_longer(cols = all_of(c(vars_firstmeas, vars_allmeas)),
               names_to = c("Measure", "Time"),
               names_pattern = "(E|gs|A)(\\d)") %>%
  mutate(Time = factor(Time)) %>%  # Ensure 'Time' is a factor
  select(ID, Field, Measure, Time, value)  # Keep 'ID', 'Field', 'Measure', 'Time', 'Value' columns

# Function to run the repeated measures ANOVA for a given subset of the data
run_repeated_for_field <- function(subset_data, measures) {
  return(run_repeated_measures_ANOVA(subset_data, measures))
}

# Check column names to verify that Field exists
if (!"Field" %in% colnames(long_df)) {
  stop("Field column not found in long_df")
}

# Verify data types to ensure everything is as expected
str(long_df)


grouped_data <- long_df %>%
  group_by(Field) %>%
  group_split()

# Apply the analysis to each group
all_repeated_results <- lapply(grouped_data, function(subset_data) {
  field_name <- unique(subset_data$Field)[1]  # Extract the Field value for the group
  result <- run_repeated_for_field(subset_data, measures)  # Apply your repeated measures analysis
  list(field_name = field_name, result = result)  # Return the field name and result
})

library(openxlsx)

library(openxlsx)

write_repeated_excel <- function(results_list, file_name) {
  wb <- createWorkbook()  # Create a new workbook
  
  # Loop through each data frame in the list and add as a new sheet
  for (i in seq_along(results_list)) {
    field_name <- results_list[[i]]$field_name  # Extract the Field level name
    
    # Ensure there's a valid field name, or use a fallback
    if (is.null(field_name) || field_name == "") {
      field_name <- paste0("Field_", i)  # Fallback name using index
    }
    
    result <- results_list[[i]]$result  # Extract the result (data frame)
    
    for (j in seq_along(result)) {  # Loop through the result data frames
      # Use the Field name and ensure unique sheet name
      sheet_name <- paste0(field_name, "_Measure_", j)
      sheet_name <- substr(sheet_name, 1, 31)  # Ensure the sheet name is no longer than 31 characters
      addWorksheet(wb, sheet_name)  # Add a new worksheet
      writeData(wb, sheet_name, result[[j]])  # Write the data frame to the sheet
    }
  }
  
  # Save the workbook
  saveWorkbook(wb, paste0(file_name, ".xlsx"), overwrite = TRUE)
}



# Generate list of field names
field_names <- unique(long_df$Field)

# Save results for all fields
# Save results for all fields
write_repeated_excel(all_repeated_results, "APA_Formatted_RepeatedMeasuresbyTrial2")




library(ggplot2)
library(dplyr)

# Function to create and save line plots for the average value of each measure across IDs for each Field
create_avg_line_plots_by_field <- function(df, measures, save_dir = "plots") {
  
  # Create a directory to save plots if it doesn't exist
  if (!dir.exists(save_dir)) {
    dir.create(save_dir)
  }
  
  # Split the data by 'Field'
  fields <- unique(df$Field)
  
  # Loop through each Field
  for (field in fields) {
    # Filter the data for the current Field
    field_data <- df %>% filter(Field == field)
    
    # Loop through each measure and create a line plot of the average values
    for (measure in measures) {
      # Filter the data for the current measure
      measure_data <- field_data %>%
        filter(Measure == measure) %>%
        group_by(Time) %>%  # Group by time points
        summarise(mean_value = mean(value, na.rm = TRUE))  # Calculate mean for each time point
      
      # Create the line plot for the average values
      p <- ggplot(measure_data, aes(x = Time, y = mean_value, group = 1)) +
        geom_line(color = "blue", size = 1) +  # Add line for the average
        geom_point(size = 3) +  # Add points for each time point
        geom_text(aes(label = sprintf("%.2f", mean_value)),  # Add data labels with 2 decimal points
                  vjust = -0.5, size = 3.5) +  # Adjust vertical position of the labels
        labs(title = paste("Average Line Plot for", measure, "in Field:", field),
             x = "Time",
             y = paste("Average Value of", measure)) +
        theme_minimal()
      
      # Display the plot
      print(p)
      
      # Save the plot to the specified directory
      filename <- paste0(save_dir, "/", gsub(" ", "_", field), "_", measure, "_line_plot.png")
      ggsave(filename, plot = p, width = 8, height = 6)
    }
  }
}

# Measures to analyze
measures <- c("E", "gs", "A")

# Run the line plot function to show and save average values with data labels for each Field
create_avg_line_plots_by_field(long_df, measures)










df <- df %>%
  rename(Hybrid_Group = Hybrid,   # Rename "Hybrid" to "Hybrid_Group"
         Hybrid = Hybrid_old)     # Rename "Hybrid_old" to "Hybrid"


# Required Libraries
library(ggplot2)
library(dplyr)

# Function to create bar graph for one-way ANOVA
plot_one_way_anova_bars <- function(df, dependent_var, independent_var, save_dir) {
  # Summarize the data
  df_summary <- df %>%
    group_by(!!sym(independent_var)) %>%
    summarize(mean_value = mean(!!sym(dependent_var), na.rm = TRUE),
              sd_value = sd(!!sym(dependent_var), na.rm = TRUE)) %>%
    ungroup()
  
  # Plot the bar graph with grayscale
  p <- ggplot(df_summary, aes(x = !!sym(independent_var), y = mean_value, fill = !!sym(independent_var))) +
    geom_bar(stat = "identity", position = position_dodge(), show.legend = FALSE) +
    geom_errorbar(aes(ymin = mean_value - sd_value, ymax = mean_value + sd_value),
                  position = position_dodge(0.9), width = 0.25) +
    scale_fill_grey(start = 0.3, end = 0.7) +  # Use grayscale fill
    labs(x = independent_var, y = dependent_var,
         title = paste("Effect of", independent_var, "on", dependent_var)) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 0, hjust = 0.5))
  
  # Save the plot
  file_name <- paste0(save_dir, "/", dependent_var, "_vs_", independent_var, "_one_way_anova.png")
  ggsave(file_name, plot = p, width = 8, height = 6)
  
  # Print the plot for visualization
  print(p)
}

# Example: Save and plot one-way ANOVA bars for all dependent variables
vars_1wayANOVA <- c("uniformity", "Early.vigor", "Average.Plant.Height", "Average.Ear.Height")
save_dir <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dorint2009/interactionplots/oneway" # Replace with the directory where you want to save the plots

for (dv in vars_1wayANOVA) {
  plot_one_way_anova_bars(df, dependent_var = dv, independent_var = "Hybrid", save_dir = save_dir)
}

for (dv in vars_allmeas) {
  plot_one_way_anova_bars(df, dependent_var = dv, independent_var = "Hybrid", save_dir = save_dir)
}

for (dv in vars_firstmeas) {
  plot_one_way_anova_bars(df, dependent_var = dv, independent_var = "Hybrid", save_dir = save_dir)
}


# Required Libraries
library(ggplot2)
library(dplyr)

# Function to create and save interaction bar graphs for two-way ANOVA with side-by-side plots for Irrigation
plot_two_way_anova_side_by_side <- function(df, dependent_var, factor1, factor2, save_dir) {
  # Check if the dependent variable exists in the dataframe
  if (!dependent_var %in% names(df)) {
    stop(paste("Error: Variable", dependent_var, "not found in the dataframe."))
  }
  
  # Ensure Irrigation is treated as a factor (categorical variable)
  df[[factor2]] <- as.factor(df[[factor2]])
  
  # Summarize the data, ignoring missing values (NA)
  df_summary <- df %>%
    group_by(!!sym(factor1), !!sym(factor2)) %>%
    summarize(mean_value = mean(!!sym(dependent_var), na.rm = TRUE),
              sd_value = sd(!!sym(dependent_var), na.rm = TRUE),
              .groups = 'drop')  # Ensure ungrouping after summarization
  
  # Check if there are no results after summarization (could happen if all values are NA)
  if (nrow(df_summary) == 0) {
    stop(paste("Error: All values for", dependent_var, "are missing or NA."))
  }
  
  # Plot the interaction bar graph with side-by-side bars and grayscale
  p <- ggplot(df_summary, aes(x = !!sym(factor1), y = mean_value, fill = !!sym(factor2))) +
    geom_bar(stat = "identity", position = position_dodge()) +
    geom_errorbar(aes(ymin = mean_value - sd_value, ymax = mean_value + sd_value),
                  position = position_dodge(0.9), width = 0.25) +
    scale_fill_grey(start = 0.3, end = 0.7) +  # Use grayscale fill
    labs(x = factor1, y = dependent_var, fill = factor2,
         title = paste("Interaction between", factor1, "and", factor2, "on", dependent_var)) +
    theme_minimal() +
    facet_wrap(as.formula(paste("~", factor2)), scales = "free") +  # Side-by-side for each Irrigation condition
    theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  # Save the plot
  file_name <- paste0(save_dir, "/", dependent_var, "_interaction_", factor1, "_", factor2, "_two_way_anova_side_by_side.png")
  ggsave(file_name, plot = p, width = 10, height = 6)
  
  # Print the plot for visualization
  print(p)
}

# Example: Save and plot two-way ANOVA side-by-side bars for all dependent variables
vars_2wayANOVA <- c("Stalk.ECB", "Fusarium", "Stay.Green", "Harvesyed.plants", "Kg_str_15.")
save_dir <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dorint2009/interactionplots/twoway" # Replace with the directory where you want to save the plots

for (dv in vars_2wayANOVA) {
  plot_two_way_anova_side_by_side(df, dependent_var = dv, factor1 = "Hybrid", factor2 = "Irrigation", save_dir = save_dir)
}

for (dv in vars_allmeas) {
  plot_two_way_anova_side_by_side(df, dependent_var = dv, factor1 = "Hybrid", factor2 = "Irrigation", save_dir = save_dir)
}





# Plots by Field

# Required Libraries
library(ggplot2)
library(dplyr)

# Function to create bar graph for one-way ANOVA by Field levels
plot_one_way_anova_bars <- function(df, dependent_var, independent_var, field_var, save_dir) {
  unique_fields <- unique(df[[field_var]])
  
  for (field_level in unique_fields) {
    # Filter by Field level
    df_field <- df %>% filter(!!sym(field_var) == field_level)
    
    # Summarize the data
    df_summary <- df_field %>%
      group_by(!!sym(independent_var)) %>%
      summarize(mean_value = mean(!!sym(dependent_var), na.rm = TRUE),
                sd_value = sd(!!sym(dependent_var), na.rm = TRUE)) %>%
      ungroup()
    
    # Plot the bar graph with grayscale
    p <- ggplot(df_summary, aes(x = !!sym(independent_var), y = mean_value, fill = !!sym(independent_var))) +
      geom_bar(stat = "identity", position = position_dodge(), show.legend = FALSE) +
      geom_errorbar(aes(ymin = mean_value - sd_value, ymax = mean_value + sd_value),
                    position = position_dodge(0.9), width = 0.25) +
      scale_fill_grey(start = 0.3, end = 0.7) + 
      labs(x = independent_var, y = dependent_var,
           title = paste("Effect of", independent_var, "on", dependent_var, "for Field:", field_level)) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 0, hjust = 0.5))
    
    # Save the plot
    file_name <- paste0(save_dir, "/", dependent_var, "_vs_", independent_var, "_", field_level, "_one_way_anova.png")
    ggsave(file_name, plot = p, width = 8, height = 6)
    
    # Print the plot for visualization
    print(p)
  }
}

# Function to create and save interaction bar graphs for two-way ANOVA with side-by-side plots for Irrigation by Field levels
plot_two_way_anova_side_by_side <- function(df, dependent_var, factor1, factor2, field_var, save_dir) {
  unique_fields <- unique(df[[field_var]])
  
  for (field_level in unique_fields) {
    # Filter by Field level
    df_field <- df %>% filter(!!sym(field_var) == field_level)
    
    # Convert factors to categorical variables to avoid continuous error
    df_field[[factor1]] <- as.factor(df_field[[factor1]])
    df_field[[factor2]] <- as.factor(df_field[[factor2]])
    
    # Summarize the data
    df_summary <- df_field %>%
      group_by(!!sym(factor1), !!sym(factor2)) %>%
      summarize(mean_value = mean(!!sym(dependent_var), na.rm = TRUE),
                sd_value = sd(!!sym(dependent_var), na.rm = TRUE),
                .groups = 'drop')
    
    # Plot the interaction bar graph
    p <- ggplot(df_summary, aes(x = !!sym(factor1), y = mean_value, fill = !!sym(factor2))) +
      geom_bar(stat = "identity", position = position_dodge()) +
      geom_errorbar(aes(ymin = mean_value - sd_value, ymax = mean_value + sd_value),
                    position = position_dodge(0.9), width = 0.25) +
      scale_fill_grey(start = 0.3, end = 0.7) + 
      labs(x = factor1, y = dependent_var, fill = factor2,
           title = paste("Interaction between", factor1, "and", factor2, "on", dependent_var, "for Field:", field_level)) +
      theme_minimal() +
      facet_wrap(as.formula(paste("~", factor2)), scales = "free") +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
    
    # Save the plot
    file_name <- paste0(save_dir, "/", dependent_var, "_interaction_", factor1, "_", factor2, "_", field_level, "_two_way_anova.png")
    ggsave(file_name, plot = p, width = 10, height = 6)
    
    # Print the plot for visualization
    print(p)
  }
}


# Example calls: Save and plot one-way ANOVA and two-way ANOVA for each dependent variable by levels of Field
vars_1wayANOVA <- c("uniformity", "Early.vigor", "Average.Plant.Height", "Average.Ear.Height")
vars_2wayANOVA <- c("Stalk.ECB", "Fusarium", "Stay.Green", "Harvesyed.plants", "Kg_str_15.")
save_dir_1way <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dorint2009/interactionplots/oneway"
save_dir_2way <- "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/dorint2009/interactionplots/twoway"

# One-way ANOVA plots
for (dv in vars_1wayANOVA) {
  plot_one_way_anova_bars(df, dependent_var = dv, independent_var = "Hybrid", field_var = "Field", save_dir = save_dir_1way)
}

# One-way ANOVA plots
for (dv in vars_allmeas) {
  plot_one_way_anova_bars(df, dependent_var = dv, independent_var = "Hybrid", field_var = "Field", save_dir = save_dir_1way)
}

# One-way ANOVA plots
for (dv in vars_firstmeas) {
  plot_one_way_anova_bars(df, dependent_var = dv, independent_var = "Hybrid", field_var = "Field", save_dir = save_dir_1way)
}

# Two-way ANOVA plots
for (dv in vars_2wayANOVA) {
  plot_two_way_anova_side_by_side(df, dependent_var = dv, factor1 = "Hybrid", factor2 = "Irrigation", field_var = "Field", save_dir = save_dir_2way)
}

for (dv in vars_allmeas) {
  plot_two_way_anova_side_by_side(df, dependent_var = dv, factor1 = "Hybrid", factor2 = "Irrigation", field_var = "Field", save_dir = save_dir_2way)
}









