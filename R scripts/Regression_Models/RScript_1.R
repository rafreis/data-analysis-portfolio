library(openxlsx)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/drpmouser")

df <- read.xlsx("Complete Geltor Clinical dataset in one spreadsheet for stats software.xlsx")

## DATA PROCESSING

library(tidyverse)

df <- df %>%
  pivot_longer(
    cols = c("T0", "T42", "T84"),
    names_to = "Timepoint",
    values_to = "Measurement"
  )

df <- df %>%
  pivot_wider(
    names_from = Parameter, 
    values_from = Measurement,
    id_cols = c(Group, Group.Name, Subject, Timepoint, Age)
  )

# Declare Dependent Variables

dvs <- c(
  "Corneometer forearm",
  "Elasticity Forearm",
  "Elasticity Face",
  "Firmness Forearm",
  "Firmness Face",
  "Wrinkles",
  "Smoothness Sesm Forearm",
  "VFV Forearm",
  "Smoothness Sesm Face",
  "VFV Face"
)

# Update column names to syntactically valid names
names(df) <- make.names(names(df))

# Also update the names in the 'dvs' vector
dvs <- make.names(dvs)

## OUTLIER INSPECTION

# Function to calculate Z-scores and identify outliers
find_outliers <- function(data, vars, threshold = 3) {
  outlier_df <- data.frame(Variable = character(), Outlier_Count = integer())
  outlier_values_df <- data.frame()
  
  for (var in vars) {
    x <- data[[var]]
    z_scores <- (x - mean(x, na.rm = TRUE)) / sd(x, na.rm = TRUE)
    
    # Identify outliers based on the Z-score threshold
    outliers <- which(abs(z_scores) > threshold)
    
    # Count outliers
    count_outliers <- length(outliers)
    
    # Store results in a summary dataframe
    outlier_df <- rbind(outlier_df, data.frame(Variable = var, Outlier_Count = count_outliers))
    
    # Store outlier values and Z-scores in a separate dataframe
    if (count_outliers > 0) {
      tmp_df <- data.frame(data[outliers, ], Z_Score = z_scores[outliers])
      tmp_df$Variable <- var
      outlier_values_df <- rbind(outlier_values_df, tmp_df)
      
      print(paste("Variable:", var, "has", count_outliers, "outliers with Z-scores greater than", threshold))
    }
  }
  
  return(list(summary = outlier_df, values = outlier_values_df))
}

# Execute the function
outlier_summary <- find_outliers(df, dvs, threshold = 3)

# Splitting the list into two separate dataframes
outlier_count <- outlier_summary$summary
outlier_values_df <- outlier_summary$values
  
# Check Boxplots
library(ggplot2)

# Generate separate boxplots for each dependent variable
for (dv in dvs) {
  plot <- ggplot(df, aes(x = Timepoint, y = !!sym(dv), fill = Group.Name)) +
    geom_boxplot() +
    labs(
      title = paste("Boxplot of", dv),
      x = "Timepoint",
      y = "Measurement Value"
    ) +
    theme(
      panel.background = element_rect(fill = "white"),
      panel.grid.major = element_line(colour = "grey92"),
      panel.grid.minor = element_line(colour = "grey92")
    )
  
  # Show plot
  print(plot)
  
  # Save plot to a file if needed
  ggsave(filename = paste0("boxplot_", dv, ".png"), plot = plot)
}

## DATA DISTRIBUTION INSPECTION

# Function to calculate normality statistics for each variable
library(moments)
library(stats)

calc_descriptive_stats_for_dvs <- function(data, dvs) {
  stats_df <- data.frame(Variable = character(), Skewness = numeric(), 
                         Kurtosis = numeric(), W = numeric(), P_Value = numeric())
  
  for (var in dvs) {
    x <- data[[var]]
    x <- x[!is.na(x)]  # Remove NA values if present
    
    if (length(x) < 3) {
      next  # Skip if less than 3 observations
    }
    
    skewness_value <- skewness(x)
    kurtosis_value <- kurtosis(x)
    shapiro_test <- shapiro.test(x)
    W_value <- shapiro_test$statistic
    p_value <- shapiro_test$p.value
    
    stats_df <- rbind(stats_df, data.frame(Variable = var, Skewness = skewness_value, 
                                           Kurtosis = kurtosis_value, W = W_value, P_Value = p_value))
  }
  
  return(stats_df)
}

# Use function to get stats for all dvs

stats_df <- calc_descriptive_stats_for_dvs(df, dvs)

# Write stats_df to an Excel file
write.xlsx(stats_df, "stats_df.xlsx")

# CHECK HOMOGENEITY OF VARIANCES

# Levene's test
library(car)

# Declare the factor variable
factor_var <- "Group.Name"

# Initialize an empty dataframe to store the results
levenes_results <- data.frame(
  Farm = character(),
  Dependent_Variable = character(),
  Factor = character(),
  F_value = numeric(),
  DF1 = numeric(),
  DF2 = numeric(),
  pValue = numeric(),
  stringsAsFactors = FALSE
)

# Loop through each dependent variable and apply Levene's test
for (var in dvs) {
  
  # Conduct Levene's Test
  levenes_test <- leveneTest(as.formula(paste(var, "~", factor_var)), data = df)
  
  # Extract relevant statistics
  F_value <- levenes_test[1, "F value"]
  DF1 <- levenes_test[1, "Df"]
  DF2 <- levenes_test[2, "Df"]
  pValue <- levenes_test[1, "Pr(>F)"]
  
  # Store the results in the dataframe
  levenes_results <- rbind(levenes_results, data.frame(
    Dependent_Variable = var,
    Factor = factor_var,
    F_value = F_value,
    DF1 = DF1,
    DF2 = DF2,
    pValue = pValue,
    stringsAsFactors = FALSE
  ))
}

# Write levenes_results to an Excel file
write.xlsx(levenes_results, "levenes_results.xlsx")

## DESCRIPTIVE STATISTICS

within_subject_factor <- "Timepoint"
between_subject_factor <- "Group.Name"

library(dplyr)

# Loop through each dependent variable
for (var in dvs) {
  
  # Calculate mean scores for each group and timepoint
  temp_mean_df <- df %>%
    group_by(!!sym(between_subject_factor), !!sym(within_subject_factor)) %>%
    summarise(Mean = mean(!!sym(var), na.rm = TRUE))
  
  # Create the line plot
  p <- ggplot(temp_mean_df, aes_string(x = within_subject_factor, y = "Mean", group = between_subject_factor, color = between_subject_factor)) +
    geom_line() +
    geom_point(size = 4) +
    labs(
      title = paste("Line Plot of", var),
      x = within_subject_factor,
      y = "Mean Value"
    ) + 
    theme(panel.background = element_rect(fill = "white"))  # Add white background
  
  # Display the plot
  print(p)
}

# Initialize an empty data frame to store the mean values
mean_values_df <- data.frame()

# Loop through each dependent variable
for (var in dvs) {
  
  # Calculate mean scores for each group and timepoint
  temp_df <- df %>%
    group_by(!!sym(between_subject_factor), !!sym(within_subject_factor)) %>%
    summarise(Mean_Value = mean(!!sym(var), na.rm = TRUE)) %>%
    spread(!!sym(within_subject_factor), Mean_Value) %>% # convert to wide format with timepoints as columns
    add_column(Variable = var, .before = 1) # add the dependent variable name
  
  # Add the mean values to the final data frame
  mean_values_df <- bind_rows(mean_values_df, temp_df)
}

# Display the mean_values_df
print(mean_values_df)


## LINEAR MIXED MODEL
# Include necessary libraries
library(nlme)
library(dplyr)

# Convert Group.Name and Timepoint to factors
df$Group.Name <- as.factor(df$Group.Name)
df$Timepoint <- as.factor(df$Timepoint)

# Initialize an empty dataframe to store model summary statistics
model_summary_df <- data.frame()

# Get unique group names
group_names <- unique(df$Group.Name)

# Loop through each dependent variable
for (var in dvs) {
  
  # Loop through each unique pair of groups
  for (i in 1:(length(group_names) - 1)) {
    for (j in (i + 1):length(group_names)) {
      
      # Filter data for the current pair of groups
      df_pair <- df %>% filter(Group.Name %in% c(group_names[i], group_names[j]))
      
      # Fit the linear mixed-effects model using nlme
      model_formula <- as.formula(paste(var, "~ Group.Name * Timepoint"))
      model <- lme(model_formula, random = ~1 | Subject, data = df_pair, method = "REML")
      
      # Get ANOVA table
      anova_table <- anova(model)
      
      # Create a temporary dataframe to store the results
      model_summary_temp <- data.frame(
        Dependent_Variable = rep(var, nrow(anova_table)),
        Group_Pair = paste(group_names[i], "-", group_names[j]),
        Effect = row.names(anova_table),
        numDF = anova_table$numDF,
        denDF = anova_table$denDF,
        F_value = anova_table$`F-value`,
        pValue = anova_table$`p-value`,
        stringsAsFactors = FALSE
      )
      
      # Append the results to the main summary dataframe
      model_summary_df <- rbind(model_summary_df, model_summary_temp)
    }
  }
}

# Same analysis in log-transformed data

# Create a copy of the original data frame
log_df <- df

# Loop through the list of dependent variables to log-transform them
for (var in dvs) {
  log_df[[var]] <- log(log_df[[var]])
}

# Use function to get stats for all dvs
stats_log_df <- calc_descriptive_stats_for_dvs(log_df, dvs)

# Initialize an empty dataframe to store model summary statistics
model_summary_log_df <- data.frame()

# Loop through each dependent variable
for (var in dvs) {
  
  # Loop through each unique pair of groups
  for (i in 1:(length(group_names) - 1)) {
    for (j in (i + 1):length(group_names)) {
      
      # Filter data for the current pair of groups
      df_pair <- log_df %>% filter(Group.Name %in% c(group_names[i], group_names[j]))
      
      # Fit the linear mixed-effects model using nlme
      model_formula <- as.formula(paste(var, "~ Group.Name * Timepoint"))
      model <- lme(model_formula, random = ~1 | Subject, data = df_pair, method = "REML", 
                   control = lmeControl(opt = "optim"))
      
      # Get ANOVA table
      anova_table <- anova(model)
      
      # Create a temporary dataframe to store the results
      model_summary_temp <- data.frame(
        Dependent_Variable = rep(var, nrow(anova_table)),
        Group_Pair = paste(group_names[i], "-", group_names[j]),
        Effect = row.names(anova_table),
        numDF = anova_table$numDF,
        denDF = anova_table$denDF,
        F_value = anova_table$`F-value`,
        pValue = anova_table$`p-value`,
        stringsAsFactors = FALSE
      )
      
      # Append the results to the main summary dataframe
      model_summary_log_df <- rbind(model_summary_df, model_summary_temp)
    }
  }
}



## ANALYSIS FOR AGE GROUPS

# Calculate mean and median age for each Group.Name
age_summary <- df %>%
  group_by(Group.Name) %>%
  summarise(
    mean_age = mean(Age, na.rm = TRUE),
    median_age = median(Age, na.rm = TRUE)
  )

# Create a boxplot of Age grouped by Group.Name
p <- ggplot(df, aes(x = Group.Name, y = Age)) +
  geom_boxplot() +
  ggtitle("Boxplot of Age by Group.Name") +
  xlab("Group Name") +
  ylab("Age")

# Display the plot
print(p)

# Filter data where Timepoint = "T84"
filtered_df_age <- df %>% filter(Timepoint == "T84")

# Loop through each dependent variable in 'dvs' to create plots
for (var in dvs) {
  p <- ggplot(filtered_df_age, aes_string(x = "Age", y = var)) +
    geom_point() +
    geom_smooth(method = "lm") +
    ggtitle(paste("Distribution of", var, "Against Age at Timepoint T84")) +
    xlab("Age") +
    ylab(var)
  
  # Display the plot
  print(p)
}


## RUN MODELS FOR DIFFERENT AGE CUTOFFS

# Initialize an empty list to store model summary data frames for each cutoff
model_summary_list <- list()

# Loop through the specified age cutoff points
for (cutoff_age in c(48, 50, 52, 54, 56, 58, 60)) {
  
  # Initialize an empty dataframe to store model summary statistics
  model_summary_df_age <- data.frame(Dependent_Variable=character(), 
                                     Effect=character(),
                                     numDF=numeric(),
                                     denDF=numeric(),
                                     F_value=numeric(),
                                     pValue=numeric(),
                                     Age_Group=character(),
                                     Group_Pair=character(),
                                     stringsAsFactors=FALSE)
  
  # Loop through age groups ("Cutoff Age or Older" and "Younger")
  for (age_group in c("Cutoff Age or Older", "Younger")) {
    
    # Filter data based on age cutoff and age group
    df_filtered <- df %>% filter(ifelse(Age >= cutoff_age, "Cutoff Age or Older", "Younger") == age_group)
    
    # Get unique group names
    group_names <- unique(df_filtered$Group.Name)
    
    # Loop through each unique pair of groups
    for (i in 1:(length(group_names) - 1)) {
      for (j in (i + 1):length(group_names)) {
        
        # Filter data for the current pair of groups
        df_pair <- df_filtered %>% filter(Group.Name %in% c(group_names[i], group_names[j]))
        
        # Loop through each dependent variable to fit the model and store results
        for (var in dvs) {
          
          # Fit the linear mixed-effects model using nlme
          model_formula <- as.formula(paste(var, "~ Group.Name * Timepoint"))
          model <- lme(model_formula, random = ~1 | Subject, data = df_pair, method = "REML")
          
          # Get ANOVA table
          anova_table <- anova(model)
          
          # Create a temporary dataframe to store the results
          model_summary_temp <- data.frame(
            Dependent_Variable = rep(var, nrow(anova_table)),
            Effect = row.names(anova_table),
            numDF = anova_table$numDF,
            denDF = anova_table$denDF,
            F_value = anova_table$`F-value`,
            pValue = anova_table$`p-value`,
            Age_Group = age_group,
            Group_Pair = paste(group_names[i], "-", group_names[j]),
            stringsAsFactors = FALSE
          )
          
          # Append the results to the main summary dataframe for this age cutoff
          model_summary_df_age <- rbind(model_summary_df_age, model_summary_temp)
        }
      }
    }
  }
  # Store the model summary data frame for this cutoff in the list
  model_summary_list[[as.character(cutoff_age)]] <- model_summary_df_age
}


library(openxlsx)

wb <- createWorkbook()

# Add a new worksheet with model_summary_df
addWorksheet(wb, "Model_Summary")
writeData(wb, "Model_Summary", model_summary_df)

# Add a new worksheet with model_summary_log_df
addWorksheet(wb, "Model_Summary_LOG")
writeData(wb, "Model_Summary_LOG", model_summary_log_df)

# Loop through each element in model_summary_list to add as new worksheets
for (cutoff_age in c(48, 50, 52, 54, 56, 58, 60)) {
  
  # Get the dataframe from the list
  current_df <- model_summary_list[[paste("Age_Cutoff", cutoff_age)]]
  
  # Add new worksheet with the current dataframe
  sheet_name <- paste("Model_Summary_Age ", cutoff_age)
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, current_df)
}

# Save the workbook
saveWorkbook(wb, "Model Results.xlsx", overwrite = TRUE)


## ALL AGE RESULTS COMBINED

# Initialize an empty data frame to store all the results
combined_df_age <- data.frame()

# Loop through each element in model_summary_list to add a column and concatenate
for (cutoff_age in c(48, 50, 52, 54, 56, 58, 60)) {
  
  # Get the dataframe from the list
  current_df <- model_summary_list[[as.character(cutoff_age)]]
  
  # Add a new column to identify the dataframe
  current_df$AgeCutoff <- cutoff_age
  
  # Concatenate the current data frame to the combined data frame
  combined_df_age <- rbind(combined_df_age, current_df)
}

# Create a new workbook
wb <- createWorkbook()

# Add a new worksheet with the combined data frame
addWorksheet(wb, "Combined_Model_Summaries")
writeData(wb, "Combined_Model_Summaries", combined_df_age)

# Save the workbook
saveWorkbook(wb, "Combined_AgeGroups_Model_Results.xlsx", overwrite = TRUE)

## OVERALL RESULTS COMBINED

# Add Analysis_Type column to model_summary_df and model_summary_log_df
model_summary_df$Analysis_Type <- "Original"
model_summary_log_df$Analysis_Type <- "Log_Transformed"

# Combine the two data frames into one
combined_summary_df <- rbind(model_summary_df, model_summary_log_df)

library(dplyr)

# Filter rows for the interaction term and where pValue < 0.10 for Overall dataframe
siginteraction_filtered_df <- combined_summary_df %>% 
  filter(Effect == "Group.Name:Timepoint" & pValue < 0.10)

# Filter rows for the interaction term and where pValue < 0.10
siginteraction_filtered_df_age <- combined_df_age %>% 
  filter(Effect == "Group.Name:Timepoint" & pValue < 0.10)

# Count the frequency for each AgeCutoff
interaction_frequency_df <- siginteraction_filtered_df_age %>% 
  group_by(AgeCutoff) %>% 
  summarise(Frequency = n())


## COMPARE MEAN SCORES OF SIGNIFICANT INTERACTIONS

# Create a list of cutoff points
cutoff_points <- c(48, 50, 52, 54, 56, 58, 60)

# Initialize an empty dataframe to store the mean scores
mean_scores_age_df <- data.frame()


# Loop through each cutoff point
for (cutoff in cutoff_points) {
  
  # Create a new column for age group based on the current cutoff
  df$Age_Group <- ifelse(df$Age >= cutoff, "Cutoff Age or Older", "Younger")
  
  # Loop through each dependent variable
  for (var in dvs) {
    
    # Calculate the mean scores for each DV, Timepoint, Group.Name, and Age_Group
    mean_scores_temp <- df %>%
      group_by(Timepoint, Group.Name, Age_Group) %>%
      summarise(Mean_Score = mean(!!sym(var), na.rm = TRUE),
                Age_Cutoff = cutoff,
                Dependent_Variable = var) %>%
      ungroup()
    
    # Append the temporary mean scores to the main dataframe
    mean_scores_age_df <- rbind(mean_scores_age_df, mean_scores_temp)
  }
}

# Define a color palette
color_palette <- c(
  "5g Primacoll" = "#D32F2F",  # Dark Red
  "Placebo" = "#1976D2",      # Dark Blue
  "10g Primacoll" = "#388E3C", # Dark Green
  "2.5g Primacoll" = "#8E24AA", # Purple
  "Vital Protein" = "#FBC02D",  # Bright Yellow
  "1g Primacolll" = "#8D6E63",  # Brown
  "Verisol" = "#424242"         # Dark Gray
)

# Initialize an empty ggplot object
p <- ggplot() +
  xlab("Timepoint") +
  ylab("Mean Score") +
  theme_minimal()

# Loop through unique combinations of Dependent_Variable, Age_Group, and AgeCutoff in siginteraction_filtered_df_age
unique_conditions <- unique(siginteraction_filtered_df_age[, c("Dependent_Variable", "Age_Group", "AgeCutoff")])

# Loop through each unique combination of Dependent_Variable, Age_Group, and AgeCutoff
for(i in 1:nrow(unique_conditions)) {
  
  # Extract the relevant information from the current row
  dv <- unique_conditions$Dependent_Variable[i]
  age_group <- unique_conditions$Age_Group[i]
  age_cutoff <- unique_conditions$AgeCutoff[i]
  
  # Get the relevant Group.Names from siginteraction_filtered_df_age
  relevant_groups <- siginteraction_filtered_df_age %>%
    filter(Dependent_Variable == dv,
           Age_Group == age_group,
           AgeCutoff == age_cutoff) %>%
    pull(Group_Pair) %>%
    unique()
  
  # Convert relevant_groups into a single character vector
  all_relevant_groups <- unique(unlist(strsplit(relevant_groups, " - ")))
  
  # Filter mean_scores_age_df based on the current combination's criteria and relevant groups
  filtered_data <- mean_scores_age_df %>%
    filter(Dependent_Variable == dv,
           Age_Group == age_group,
           Age_Cutoff == age_cutoff,
           Group.Name %in% all_relevant_groups)
  
  # Check if filtered_data is empty
  if(nrow(filtered_data) == 0) next
  
  # Generate the plot for the current set of criteria
  p <- ggplot(filtered_data, aes(x=Timepoint, y=Mean_Score, group=Group.Name, color=Group.Name)) +
    geom_line() +
    geom_point() +
    ggtitle(paste("Dependent Variable:", dv, "\nAge Group:", age_group, "\nCutoff:", age_cutoff)) +
    xlab("Timepoint") +
    ylab("Mean Score") +
    scale_color_manual(values = color_palette) +
    theme(
      panel.background = element_rect(fill = "white"),
      panel.grid.major = element_line(colour = "grey92"),
      panel.grid.minor = element_line(colour = "grey92")
    )
  
  
  # Show the plot
  print(p)
  
  # Save the figure
  ggsave(filename = paste0("figures/", dv, "_", age_group, "_Cutoff_", age_cutoff, ".png"), plot = p)
}



## COMPARE MEAN SCORES OF SIGNIFICANT INTERACTIONS - OVERALL

# Initialize an empty dataframe to store the mean scores
mean_scores_df <- data.frame()


  # Loop through each dependent variable
  for (var in dvs) {
    
    # Calculate the mean scores for each DV, Timepoint, Group.Name, and Age_Group
    mean_scores_temp <- df %>%
      group_by(Timepoint, Group.Name) %>%
      summarise(Mean_Score = mean(!!sym(var), na.rm = TRUE),
                Dependent_Variable = var) %>%
      ungroup()
    
    # Append the temporary mean scores to the main dataframe
    mean_scores_df <- rbind(mean_scores_df, mean_scores_temp)
  }

# Filter siginteraction_filtered_df for Analysis_type = "Original"
siginteraction_filtered_df_nonlog <- siginteraction_filtered_df %>%
  filter(Analysis_Type == "Original")

# Define a color palette
color_palette <- c(
  "5g Primacoll" = "#D32F2F",  # Dark Red
  "Placebo" = "#1976D2",      # Dark Blue
  "10g Primacoll" = "#388E3C", # Dark Green
  "2.5g Primacoll" = "#8E24AA", # Purple
  "Vital Protein" = "#FBC02D",  # Bright Yellow
  "1g Primacolll" = "#8D6E63",  # Brown
  "Verisol" = "#424242"         # Dark Gray
)


# Loop through each unique Dependent_Variable in siginteraction_filtered_df
for (dv in unique(siginteraction_filtered_df_nonlog$Dependent_Variable)) {
  
  # Get the relevant Group.Names from siginteraction_filtered_df
  relevant_groups <- siginteraction_filtered_df_nonlog %>%
    filter(Dependent_Variable == dv) %>%
    pull(Group_Pair) %>%
    unique()
  
  # Convert relevant_groups into a single character vector
  all_relevant_groups <- unique(unlist(strsplit(relevant_groups, " - ")))
  
  # Filter mean_scores_df based on the current Dependent_Variable and relevant groups
  filtered_data <- mean_scores_df %>%
    filter(Dependent_Variable == dv,
           Group.Name %in% all_relevant_groups)
  
  # Check if filtered_data is empty
  if (nrow(filtered_data) == 0) next
  
  # Generate the plot for the current Dependent_Variable
  p <- ggplot(filtered_data, aes(x = Timepoint, y = Mean_Score, group = Group.Name, color = Group.Name)) +
    geom_line() +
    geom_point() +
    ggtitle(paste("Dependent Variable:", dv)) +
    xlab("Timepoint") +
    ylab("Mean Score") +
    scale_color_manual(values = color_palette) +
    theme(
      panel.background = element_rect(fill = "white"),
      panel.grid.major = element_line(colour = "grey92"),
      panel.grid.minor = element_line(colour = "grey92")
    )
  
  # Show the plot
  print(p)
  
  # Save the figure
  ggsave(filename = paste0("figures_overall/figure_", dv, ".png"), plot = p)
}



## EXPORT SIGNIFICANT TABLES

#Adjust the table with means for age groups
library(tidyverse)

# Assuming mean_scores_age_df is your original data frame
mean_scores_age_df <- mean_scores_age_df %>%
  spread(key = Timepoint, value = Mean_Score)

## Join mean values to model table

# Function to fetch the mean score for a specific group and timepoint
get_mean_score <- function(group, dependent_variable, timepoint) {
  mean_score <- mean_values_df_long$Mean_Score[mean_values_df_long$Group.Name == group & 
                                                 mean_values_df_long$Variable == dependent_variable &
                                                 mean_values_df_long$Timepoint == timepoint]
  if (length(mean_score) == 0) {
    print(paste("No mean score found for Group:", group, "Variable:", dependent_variable, "Timepoint:", timepoint))
  }
  return(mean_score)
}

# First, split the Group_Pair into separate columns for easier joining
siginteraction_filtered_df <- siginteraction_filtered_df %>%
  separate(Group_Pair, c("Group1", "Group2"), sep = " - ")

# Reshape mean_values_df from wide to long format for easier joining
mean_values_long <- mean_values_df %>%
  gather(key = "Timepoint", value = "Mean_Score", T0, T42, T84)

# Function to fetch the mean score for a specific group and timepoint
get_mean_score <- function(group, dependent_variable, timepoint) {
  mean_score <- mean_values_df_long$Mean_Score[mean_values_df_long$Group.Name == group & 
                                            mean_values_df_long$Variable == dependent_variable &
                                              mean_values_df_long$Timepoint == timepoint]
  return(mean_score)
}

# Initialize new columns
siginteraction_filtered_df$Mean_Score_Group1_T0 <- NA
siginteraction_filtered_df$Mean_Score_Group1_T42 <- NA
siginteraction_filtered_df$Mean_Score_Group1_T84 <- NA
siginteraction_filtered_df$Mean_Score_Group2_T0 <- NA
siginteraction_filtered_df$Mean_Score_Group2_T42 <- NA
siginteraction_filtered_df$Mean_Score_Group2_T84 <- NA

# Populate new columns with mean scores
for (i in 1:nrow(siginteraction_filtered_df)) {
  row <- siginteraction_filtered_df[i, ]
  group1 <- row$Group1
  group2 <- row$Group2
  dependent_variable <- row$Dependent_Variable
  
  siginteraction_filtered_df$Mean_Score_Group1_T0[i] <- get_mean_score(group1, dependent_variable, 'T0')
  siginteraction_filtered_df$Mean_Score_Group1_T42[i] <- get_mean_score(group1, dependent_variable, 'T42')
  siginteraction_filtered_df$Mean_Score_Group1_T84[i] <- get_mean_score(group1, dependent_variable, 'T84')
  
  siginteraction_filtered_df$Mean_Score_Group2_T0[i] <- get_mean_score(group2, dependent_variable, 'T0')
  siginteraction_filtered_df$Mean_Score_Group2_T42[i] <- get_mean_score(group2, dependent_variable, 'T42')
  siginteraction_filtered_df$Mean_Score_Group2_T84[i] <- get_mean_score(group2, dependent_variable, 'T84')
}


# Function to calculate the percentage difference between two numbers
calculate_percentage_difference <- function(a, b) {
  if (is.na(a) || is.na(b) || b == 0) {
    return(NA)
  }
  return((a - b) / b)
}

# Initialize new columns for percentage differences between timepoints
siginteraction_filtered_df$Percentage_Difference_Group1_T0_T42 <- NA
siginteraction_filtered_df$Percentage_Difference_Group1_T42_T84 <- NA
siginteraction_filtered_df$Percentage_Difference_Group2_T0_T42 <- NA
siginteraction_filtered_df$Percentage_Difference_Group2_T42_T84 <- NA

# Populate the new columns with percentage differences
for (i in 1:nrow(siginteraction_filtered_df)) {
  row <- siginteraction_filtered_df[i, ]
  
  mean_group1_t0 <- row$Mean_Score_Group1_T0
  mean_group1_t42 <- row$Mean_Score_Group1_T42
  mean_group1_t84 <- row$Mean_Score_Group1_T84
  
  mean_group2_t0 <- row$Mean_Score_Group2_T0
  mean_group2_t42 <- row$Mean_Score_Group2_T42
  mean_group2_t84 <- row$Mean_Score_Group2_T84
  
  siginteraction_filtered_df$Percentage_Difference_Group1_T0_T42[i] <- calculate_percentage_difference(mean_group1_t42, mean_group1_t0)
  siginteraction_filtered_df$Percentage_Difference_Group1_T42_T84[i] <- calculate_percentage_difference(mean_group1_t84, mean_group1_t42)
  
  siginteraction_filtered_df$Percentage_Difference_Group2_T0_T42[i] <- calculate_percentage_difference(mean_group2_t42, mean_group2_t0)
  siginteraction_filtered_df$Percentage_Difference_Group2_T42_T84[i] <- calculate_percentage_difference(mean_group2_t84, mean_group2_t42)
}

## Same for AGE GROUPS

# Convert to long format
mean_scores_age_df_long <- mean_scores_age_df %>%
  gather(key = "Timepoint", value = "Mean_Score", T0, T42, T84)

# Split the Group_Pair into separate columns for Group1 and Group2
siginteraction_filtered_df_age <- siginteraction_filtered_df_age %>%
  separate(Group_Pair, c("Group1", "Group2"), sep = " - ")

# Function to fetch the mean score for a specific group, age group, and dependent variable
get_mean_score_age <- function(group, age_group, age_cutoff, dependent_variable, timepoint) {
  mean_score <- mean_scores_age_df_long$Mean_Score[mean_scores_age_df_long$Group.Name == group & 
                                                     mean_scores_age_df_long$Age_Group == age_group &
                                                     mean_scores_age_df_long$Age_Cutoff == age_cutoff &
                                                     mean_scores_age_df_long$Dependent_Variable == dependent_variable &
                                                     mean_scores_age_df_long$Timepoint == timepoint]
  return(mean_score)
}

# Initialize new columns
siginteraction_filtered_df_age$Mean_Score_Group1_T0 <- NA
siginteraction_filtered_df_age$Mean_Score_Group1_T42 <- NA
siginteraction_filtered_df_age$Mean_Score_Group1_T84 <- NA
siginteraction_filtered_df_age$Mean_Score_Group2_T0 <- NA
siginteraction_filtered_df_age$Mean_Score_Group2_T42 <- NA
siginteraction_filtered_df_age$Mean_Score_Group2_T84 <- NA

# Populate new columns with mean scores
for (i in 1:nrow(siginteraction_filtered_df_age)) {
  row <- siginteraction_filtered_df_age[i, ]
  group1 <- row$Group1
  group2 <- row$Group2
  dependent_variable <- row$Dependent_Variable
  age_group <- row$Age_Group
  age_cutoff <- row$AgeCutoff
  
  siginteraction_filtered_df_age$Mean_Score_Group1_T0[i] <- get_mean_score_age(group1, age_group, age_cutoff, dependent_variable, 'T0')
  siginteraction_filtered_df_age$Mean_Score_Group1_T42[i] <- get_mean_score_age(group1, age_group, age_cutoff, dependent_variable, 'T42')
  siginteraction_filtered_df_age$Mean_Score_Group1_T84[i] <- get_mean_score_age(group1, age_group, age_cutoff, dependent_variable, 'T84')
  
  siginteraction_filtered_df_age$Mean_Score_Group2_T0[i] <- get_mean_score_age(group2, age_group, age_cutoff, dependent_variable, 'T0')
  siginteraction_filtered_df_age$Mean_Score_Group2_T42[i] <- get_mean_score_age(group2, age_group, age_cutoff, dependent_variable, 'T42')
  siginteraction_filtered_df_age$Mean_Score_Group2_T84[i] <- get_mean_score_age(group2, age_group, age_cutoff, dependent_variable, 'T84')
}


# Calculate percentage change between timepoints for each group
siginteraction_filtered_df_age$Pct_Change_Group1_T0_T42 <- ((siginteraction_filtered_df_age$Mean_Score_Group1_T42 - siginteraction_filtered_df_age$Mean_Score_Group1_T0) / siginteraction_filtered_df_age$Mean_Score_Group1_T0)
siginteraction_filtered_df_age$Pct_Change_Group1_T42_T84 <- ((siginteraction_filtered_df_age$Mean_Score_Group1_T84 - siginteraction_filtered_df_age$Mean_Score_Group1_T42) / siginteraction_filtered_df_age$Mean_Score_Group1_T42)

siginteraction_filtered_df_age$Pct_Change_Group2_T0_T42 <- ((siginteraction_filtered_df_age$Mean_Score_Group2_T42 - siginteraction_filtered_df_age$Mean_Score_Group2_T0) / siginteraction_filtered_df_age$Mean_Score_Group2_T0)
siginteraction_filtered_df_age$Pct_Change_Group2_T42_T84 <- ((siginteraction_filtered_df_age$Mean_Score_Group2_T84 - siginteraction_filtered_df_age$Mean_Score_Group2_T42) / siginteraction_filtered_df_age$Mean_Score_Group2_T42)





# Include the openxlsx library
library(openxlsx)

# Create a new workbook
wb <- createWorkbook()

# Add a worksheet for siginteraction_filtered_df_age
addWorksheet(wb, "SigInteractions_Age")
writeData(wb, "SigInteractions_Age", siginteraction_filtered_df_age)

# Add a worksheet for siginteraction_filtered_df
addWorksheet(wb, "SigInteractions")
writeData(wb, "SigInteractions", siginteraction_filtered_df)

# Add a worksheet for mean_scores_df
addWorksheet(wb, "Mean_Scores")
writeData(wb, "Mean_Scores", mean_values_df)

# Add a worksheet for mean_scores_age_df
addWorksheet(wb, "Mean_Scores_Age")
writeData(wb, "Mean_Scores_Age", mean_scores_age_df)

# Save the workbook
saveWorkbook(wb, "Significant_Interactions_and_Means_v2.xlsx", overwrite = TRUE)


# Count Age groups
# Initialize a dataframe to store the counts
age_group_count_df <- data.frame(Cutoff_Age=numeric(), 
                                 Age_Group=character(),
                                 Group_Name=character(), 
                                 Count=numeric(),
                                 stringsAsFactors=FALSE)

# Loop through specified age cutoff points
for (cutoff_age in c(48, 50, 52, 54, 56, 58, 60)) {
  
  # Loop through age groups ("Cutoff Age or Older" and "Younger")
  for (age_group in c("Cutoff Age or Older", "Younger")) {
    
    # Filter data based on age cutoff, age group, and timepoint
    df_filtered <- df %>% 
      filter(Timepoint == "T0") %>%
      filter(ifelse(Age >= cutoff_age, "Cutoff Age or Older", "Younger") == age_group)
    
    # Count by Group.Name
    group_counts <- df_filtered %>% 
      group_by(Group.Name) %>%
      summarise(Count = n())
    
    # Append the counts, cutoff age, and age group to the dataframe
    for (i in 1:nrow(group_counts)) {
      age_group_count_df <- rbind(age_group_count_df, 
                                  data.frame(Cutoff_Age=cutoff_age, 
                                             Age_Group=age_group, 
                                             Group_Name=group_counts$Group.Name[i], 
                                             Count=group_counts$Count[i]))
    }
  }
}
