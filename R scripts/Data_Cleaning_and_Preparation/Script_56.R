library(googlesheets4)


#gs4_auth()  # This will open a browser window for you to authenticate

# Specify the Google Sheet URL
sheet_url <- "https://docs.google.com/spreadsheets/d/1gRG7F7NrReKwrCjrMsvptPYPUxtD5DUJGW7AmxZx2_o/edit#gid=1068712586"

# Get all sheet names
sheet_names <- sheet_names(sheet_url)

# Remove the last three sheet names
sheet_names <- sheet_names[1:(length(sheet_names) - 3)]
sheet_names
# List to store problematic rows
problematic_rows <- list()

# Function to read data from a single sheet and add the sheet name as a column
read_sheet_with_name <- function(sheet_name, sheet_url) {
  tryCatch({
    data <- read_sheet(sheet_url, sheet = sheet_name)
    data$SheetName <- sheet_name
    return(data)
  }, error = function(e) {
    message(paste("Error reading sheet:", sheet_name, "-", e$message))
    problematic_rows <<- append(problematic_rows, list(data.frame(SheetName = sheet_name, Error = e$message)))
    return(NULL)
  })
}

# Read data from all sheets except the last three and combine into a single dataframe
all_data_list <- lapply(sheet_names, read_sheet_with_name, sheet_url = sheet_url)

# Filter out NULLs and standardize columns to character type
all_data_list <- all_data_list[!sapply(all_data_list, is.null)]
all_data_list <- lapply(all_data_list, function(df) {
  df[] <- lapply(df, as.character)
  return(df)
})

# Combine all dataframes into one
all_data <- do.call(rbind, all_data_list)


# Create a named vector for each question to map text responses to numerical codes
question_mappings <- list(
  `How does this leader contribute to the growth of the company and prioritize the company's interests over their team's interests?` = c(
    "Always Prioritizes Company - Consistently places company goals and interests above personal and team considerations, actively seeking outcomes that benefit the broader organization." = 4,
    "Usually Prioritizes Company - Often aligns with company interests, though occasionally balances or prioritizes team or personal goals when beneficial." = 3,
    "Sometimes Prioritizes Company - Balances between company, team, and personal interests, with variable alignment to the company's overall objectives." = 2,
    "Rarely Prioritizes Company - Frequently prioritizes personal or team interests over the company’s goals, often at the expense of broader organizational objectives." = 1
  ),
  `Are they able to build, manage, and lead a team effectively to reach their goals?` = c(
    "Exceptionally Well - Exemplifies leadership, consistently inspiring and effectively managing the team to achieve top performance." = 4,
    "Effectively - Leads and manages the team well, meeting goals and overcoming challenges competently." = 3,
    "Adequately - Manages the team with moderate success, though some areas need improvement." = 2,
    "Poorly - Struggles with leadership and management, resulting in underperformance or team dissatisfaction." = 1
  ),
  `Do they create a mindset in their team / area of excellence, continuous improvement  and innovation?` = c(
    "Highly Innovative - Continuously drives growth and innovation within the team, setting a benchmark in the organization." = 4,
    "Consistently Improving - Regularly introduces new ideas and processes that lead to team growth and improvement." = 3,
    "Occasionally Innovative - Sometimes pushes the team towards growth and innovation, but lacks consistency." = 2,
    "Rarely Innovative - Struggles to implement growth or innovation initiatives within the team." = 1
  ),
  `Does this leader actively develop and promote their own people?` = c(
    "Proactively Develops - Regularly invests in the development and promotion of team members, with clear outcomes and advancements." = 4,
    "Generally Supports Development - Supports team members' development and career progression, with some successes." = 3,
    "Inconsistently Develops - Shows sporadic effort in developing and promoting team members, with mixed results." = 2,
    "Neglects Development - Rarely takes action to develop or promote the growth of team members." = 1
  ),
  `How effective is this person in influencing others and gaining buy-in their area?` = c(
    "Highly effective - Regularly influences others and successfully gains support for ideas." = 4,
    "Effectively - Often persuades others and gets buy-in, with some exceptions." = 3,
    "Moderately effective - Sometimes struggles but can occasionally sway opinions." = 2,
    "Ineffectively - Rarely convinces others or gains support for their ideas." = 1
  ),
  `How infectiously passionate is this leader when it comes to Clearwater and driving growth within their team?` = c(
    "Highly engaged - Shows exceptional dedication and enthusiasm for their work, team and the company." = 4,
    "Engaged - Regularly involved and committed, with moments of high enthusiasm." = 3,
    "Somewhat engaged - Engages with work adequately but lacks consistent enthusiasm." = 2,
    "Minimally engaged - Shows limited engagement and lackluster involvement in organizational activities." = 1
  ),
  `As you think about this person and when you need support, are they:` = c(
    "Always willing to support/get involved - Readily available and actively offers assistance." = 4,
    "Usually too busy - Often preoccupied and hard to engage for support." = 3,
    "Supports when possible - Available to assist occasionally, but not consistently." = 2,
    "Rarely available to assist - Infrequently accessible and tends to delegate support to others." = 1
  ),
  `Does this leader demonstrate the ability to adapt in pursuit of company goals or to new situations, e.g., Gen AI/ML, GO initiatives?` = c(
    "Highly adaptable - Regularly adjusts tactics and strategies to meet evolving organizational needs." = 4,
    "Mostly adaptable - Willing to change but prefers sticking to tried and tested methods." = 3,
    "Seldom adaptable - Resists change and infrequently aligns with new directions." = 2,
    "Not adaptable - Sticks to traditional approaches, ignoring new opportunities or requirements." = 1
  ),
  `When thinking about this person taking on additional leadership responsibilities to actively support growth, are they:` = c(
    "Ready now - Fully prepared and capable of taking on more leadership roles to drive growth and achieve Clearwater's potential"= 4,
    "Ready in 6-12 months - Will be prepared for additional responsibilities soon with some development." = 3,
    "Ready in 12+ months - Needs time and/or training to be ready for more leadership opportunites related to driving growth" = 2,
    "Not capable - Currently not suitable for additional leadership responsibilities." = 1
  ),
  `Does this person actively contribute to finding and developing strategic opportunities for company growth?` = c(
    "Significantly - Always at the forefront of identifying and developing growth opportunities." = 4,
    "Considerably - Frequently contributes valuable ideas for strategic growth." = 3,
    "Occasionally - From time to time, contributes ideas for growth." = 2,
    "Rarely - Seldom involved in strategic planning or growth initiatives." = 1
  ),
  `What has been the leader's impact on their team and organizational goals around driving growth?` = c(
    "Transformative Impact - Has significantly and positively impacted both team and organizational goals, driving substantial advancements." = 4,
    "Positive Impact - Consistently contributes to meeting and occasionally exceeding team and organizational goals." = 3,
    "Minimal Impact - Has a limited but noticeable impact on achieving team and organizational goals." = 2,
    "No Impact or Negative Impact - Has had little to no positive impact, or has negatively impacted the team and organizational goals." = 1
  )
)

# Function to map text responses to numerical codes
map_to_codes <- function(data, mappings) {
  for (question in names(mappings)) {
    # Ensure the column is a character vector
    data[[question]] <- as.character(data[[question]])
    
    # Trim whitespace and preprocess the column to match the mapping keys
    data[[question]] <- trimws(data[[question]])
    
    # Apply the mapping
    tryCatch({
      data[[question]] <- as.numeric(mappings[[question]][data[[question]]])
    }, error = function(e) {
      print(paste("Error mapping question:", question))
      print(paste("Data:", toString(data[[question]])))
      print(paste("Mapping:", toString(mappings[[question]])))
      stop(e)
    })
  }
  return(data)
}

# Apply the mapping function to the data
mapped_data <- map_to_codes(all_data, question_mappings)

mapped_data <- mapped_data[!is.na(mapped_data$Token), ]

# Rename SheetName to Individual
names(mapped_data)[names(mapped_data) == "SheetName"] <- "Leader"
names(all_data)[names(all_data) == "SheetName"] <- "Leader"

## RELIABILITY ANALYSIS

# Scale Construction

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Added for item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val <- alpha_results$total$raw_alpha
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"] # Get ITC for the item
      item_mean <- mean(item_data, na.rm = TRUE)
      item_sem <- sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data)))
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = item_mean,
        SEM = item_sem,
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
        Alpha = NA
      ))
    }
    
    # Calculate the sum score for the scale and add as a new column
    scale_sum <- rowSums(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_sum
    scale_mean <- mean(scale_sum, na.rm = TRUE)
    scale_sem <- sd(scale_sum, na.rm = TRUE) / sqrt(sum(!is.na(scale_sum)))
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = scale_mean,
      SEM = scale_sem,
      StDev = sd(scale_sum, na.rm = TRUE),
      ITC = NA,  # ITC is not applicable for the total scale
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}



scales <- list(
  "Leadership Skills" = c("How does this leader contribute to the growth of the company and prioritize the company's interests over their team's interests?",
                          "Are they able to build, manage, and lead a team effectively to reach their goals?",
                          "Do they create a mindset in their team / area of excellence, continuous improvement  and innovation?",
                          "Does this leader actively develop and promote their own people?",
                          "How effective is this person in influencing others and gaining buy-in their area?"
                          ),
  
  "Mindset & Attitude" = c("How infectiously passionate is this leader when it comes to Clearwater and driving growth within their team?"    ,                    
                           "As you think about this person and when you need support, are they:"  ,                                                               
                           "Does this leader demonstrate the ability to adapt in pursuit of company goals or to new situations, e.g., Gen AI/ML, GO initiatives?"),
  
  "Growth / Potential" = c("When thinking about this person taking on additional leadership responsibilities to actively support growth, are they:" ,             
                           "Does this person actively contribute to finding and developing strategic opportunities for company growth?"           ,               
                           "What has been the leader's impact on their team and organizational goals around driving growth?"  ),
  
  
  "Skills, Mindset & Attitude" = c("How does this leader contribute to the growth of the company and prioritize the company's interests over their team's interests?",
                                   "Are they able to build, manage, and lead a team effectively to reach their goals?",
                                   "Do they create a mindset in their team / area of excellence, continuous improvement  and innovation?",
                                   "Does this leader actively develop and promote their own people?",
                                   "How effective is this person in influencing others and gaining buy-in their area?",
                                   "How infectiously passionate is this leader when it comes to Clearwater and driving growth within their team?"    ,                    
                                   "As you think about this person and when you need support, are they:"  ,                                                               
                                   "Does this leader demonstrate the ability to adapt in pursuit of company goals or to new situations, e.g., Gen AI/ML, GO initiatives?"
                                   ),
  
  "Leadership Effectiveness" = c("How does this leader contribute to the growth of the company and prioritize the company's interests over their team's interests?",
                                 "Are they able to build, manage, and lead a team effectively to reach their goals?",
                                 "Do they create a mindset in their team / area of excellence, continuous improvement  and innovation?",
                                 "Does this leader actively develop and promote their own people?",
                                 "How effective is this person in influencing others and gaining buy-in their area?",
                                 "How infectiously passionate is this leader when it comes to Clearwater and driving growth within their team?"    ,                    
                                 "As you think about this person and when you need support, are they:"  , 
                                 "Does this leader demonstrate the ability to adapt in pursuit of company goals or to new situations, e.g., Gen AI/ML, GO initiatives?",
                                 "When thinking about this person taking on additional leadership responsibilities to actively support growth, are they:" ,             
                                 "Does this person actively contribute to finding and developing strategic opportunities for company growth?"           ,               
                                 "What has been the leader's impact on their team and organizational goals around driving growth?"  )
                               
)


alpha_results <- reliability_analysis(mapped_data, scales)

df_recoded <- alpha_results$data_with_scales
df_descriptives <- alpha_results$statistics

factor_analysis_loadings <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Scale = character(),
    Variable = character(),
    FactorLoading = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[,scales[[scale]]]
    
    # Remove rows with any missing values in the subset data
    subset_data <- subset_data[complete.cases(subset_data), ]
    
    # Perform factor analysis; using 1 factor as it's the most common scenario
    fa_results <- factanal(subset_data, factors = 1, rotation = "varimax")
    
    # Extract factor loadings
    loadings <- fa_results$loadings[,1]
    
    # Append results for each item in the scale
    for (item in scales[[scale]]) {
      results <- rbind(results, data.frame(
        Scale = scale,
        Variable = item,
        FactorLoading = loadings[item]
      ))
    }
    # Calculate total variance explained
    variance_explained <- sum(loadings^2) / length(loadings)
    
    # Print summary of the results - Total Variance Explained
    cat("Scale:", scale, "\n")
    cat("Total Variance Explained by Factor:", variance_explained, "\n")
  }
  
  return(results)
}

df_validity <- factor_analysis_loadings(df_recoded, scales)

scales_scores <- c("Leadership Skills"  ,                                                                                                                 
                   "Mindset & Attitude"   ,         
                   "Skills, Mindset & Attitude",
                   "Growth / Potential",
                   "Leadership Effectiveness")

library(stringr)


# Adjust Leader names
df_recoded <- df_recoded %>%
  mutate(Leader = str_replace(Leader, "^LA  ", ""))

# Change to %
df_recoded <- df_recoded %>%
  mutate(
    `Leadership Skills` = ((`Leadership Skills` - 5) / (20 - 5)) * 100,
    `Mindset & Attitude` = ((`Mindset & Attitude` - 3) / (12 - 3)) * 100,
    `Skills, Mindset & Attitude` = ((`Skills, Mindset & Attitude` - 8) / (32 - 8)) * 100,
    `Growth / Potential` = ((`Growth / Potential` - 3) / (12 - 3)) * 100,
    `Leadership Effectiveness` = ((`Leadership Effectiveness` - 11) / (44 - 11)) * 100
  )

## BOXPLOTS 
library(ggplot2)
library(tidyr)
library(stringr)

create_boxplots <- function(df, leader_col, vars) {
  # Ensure that vars are in the dataframe
  df <- df %>% select(all_of(c(leader_col, vars)))
  
  # Reshape the data to a long format
  long_df <- df %>%
    gather(key = "Variable", value = "Value", -all_of(leader_col))
  
  # Ensure the variable order is as specified in vars
  long_df$Variable <- factor(long_df$Variable, levels = vars)
  
  # Create separate boxplots for each leader
  leader_levels <- unique(df[[leader_col]])
  for (leader in leader_levels) {
    leader_df <- long_df %>% filter(!!sym(leader_col) == leader)
    p <- ggplot(leader_df, aes(x = Variable, y = Value, fill = Variable)) +
      geom_boxplot(outlier.color = "black", outlier.shape = 16) +
      stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
      labs(title = paste("Boxplots for Leader:", leader), y = "Value") +
      theme(axis.text.x = element_text(vjust = 1)) +
      scale_x_discrete(labels = function(x) str_wrap(x, width = 10)) +
      scale_y_continuous(limits = c(0, 100), breaks = seq(0, 100, by = 20)) +
      scale_fill_brewer(palette = "Set3")
    
    # Print the plot
    print(p)
    
    # Save the plot
    ggsave(filename = paste0("boxplots_leader_", leader, ".png"), plot = p, width = 10, height = 6)
  }
}

create_boxplots(df_recoded, "Leader", scales_scores)

# Outlier Evaluation

library(dplyr)

calculate_z_scores <- function(data, vars, id_var, leader_col, z_threshold = 3) {
  z_score_data <- data %>%
    group_by(!!sym(leader_col)) %>%
    mutate(across(all_of(vars), ~ (.-mean(.))/sd(.), .names = "z_{.col}")) %>%
    ungroup() %>%
    select(!!sym(id_var), !!sym(leader_col), starts_with("z_"))
  
  # Add a column to flag outliers based on the specified z-score threshold
  z_score_data <- z_score_data %>%
    mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))
  
  return(z_score_data)
}

# Create ID column

df_recoded <- cbind(ID = 1:nrow(df_recoded ), df_recoded)

id_var <- "ID"
df_zscores <- calculate_z_scores(df_recoded, scales_scores, "ID", "Leader")

# Remove Outliers with Z <> +-3

colnames(df_zscores) <- gsub("\\.", "_", colnames(df_zscores))

# Function to remove values in df_recoded based on z-score table
remove_outliers_based_on_z_scores <- function(data, z_score_data, id_var, leader_col) {
  # Reshape the z-score data to long format
  z_score_data_long <- z_score_data %>%
    pivot_longer(cols = starts_with("z_"), names_to = "Variable", values_to = "z_value") %>%
    pivot_longer(cols = starts_with("flag_z_"), names_to = "Flag_Variable", values_to = "Flag_Value") %>%
    filter(str_replace(Flag_Variable, "flag_", "") == Variable) %>%
    select(!!sym(id_var), !!sym(leader_col), Variable, Flag_Value)
  
  # Remove rows where z-scores are flagged as outliers
  for (i in 1:nrow(z_score_data_long)) {
    if (z_score_data_long$Flag_Value[i] == "Outlier") {
      col_name <- gsub("z_", "", z_score_data_long$Variable[i])
      data <- data %>%
        mutate(!!sym(col_name) := ifelse(ID == z_score_data_long[[id_var]][i] & Leader == z_score_data_long[[leader_col]][i], NA, !!sym(col_name)))
    }
  }
  return(data)
}

df_recoded_nooutliers <- remove_outliers_based_on_z_scores(df_recoded, df_zscores, id_var, "Leader")


library(ggplot2)
library(dplyr)
library(tidyr)
library(broom)

# Function to create mean and 95% CI plot
create_mean_ci_plots <- function(data, variables, factor_column) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    group_by(!!sym(factor_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      N = n(),
      SD = sd(Value, na.rm = TRUE),
      SEM = SD / sqrt(N),
      .groups = 'drop'
    ) %>%
    mutate(
      CI = qt(0.975, df = N-1) * SEM,  # 95% confidence interval
      Lower = Mean - CI,
      Upper = Mean + CI
    )
  
  # Loop through each variable and create separate plots
  for (variable in variables) {
    plot_data <- long_data %>% filter(Variable == variable)
    p <- ggplot(plot_data, aes(x = !!sym(factor_column), y = Mean)) +
      geom_point(aes(color = !!sym(factor_column)), size = 3) +
      geom_errorbar(aes(ymin = Lower, ymax = Upper, color = !!sym(factor_column)), width = 0.2) +
      theme(
        axis.text.x = element_text(angle = 90, vjust = 0.5),
        legend.position = "none"  # Remove the legend
      ) +
      labs(
        title = paste("Mean and 95% Confidence Interval for", variable), 
        
        y = "Value"
      ) +
      scale_y_continuous(limits = c(0, 100), breaks = seq(0, 100, by = 20))
    
    # Print the plot
    print(p)
    
    # Save the plot
    ggsave(filename = paste0("mean_ci_plot_", variable, ".png"), plot = p, width = 10, height = 6)
  }
}

leader_col <- "Leader"  

# Define the scales in the desired order
scales_scores <- c(
  "Leadership Skills",
  "Mindset & Attitude",
  "Skills, Mindset & Attitude",
  "Growth / Potential",
  "Leadership Effectiveness"
)

# Create the mean and 95% CI plot
create_mean_ci_plots(df_recoded_nooutliers, scales_scores, leader_col)

# Column Plots

scales_scores2 <- c("Skills, Mindset & Attitude",
                   "Growth / Potential"
                   )


#Function for each leader

library(ggplot2)
library(dplyr)
library(tidyr)

# Function to create mean and 95% CI plot for a specific group
create_mean_ci_plots <- function(data, variables, factor_column, group_name, leaders) {
  # Filter the data for the specified leaders
  group_data <- data %>% filter(!!sym(factor_column) %in% leaders)
  
  # Create a long format of the data for ggplot
  long_data <- group_data %>%
    pivot_longer(cols = all_of(variables), names_to = "Variable", values_to = "Value") %>%
    group_by(!!sym(factor_column), Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      N = n(),
      SD = sd(Value, na.rm = TRUE),
      SEM = SD / sqrt(N),
      .groups = 'drop'
    ) %>%
    mutate(
      CI = qt(0.975, df = N-1) * SEM,  # 95% confidence interval
      Lower = Mean - CI,
      Upper = Mean + CI
    )
  
  # Loop through each variable and create separate plots
  for (variable in variables) {
    plot_data <- long_data %>% filter(Variable == variable)
    p <- ggplot(plot_data, aes(x = !!sym(factor_column), y = Mean)) +
      geom_point(aes(color = !!sym(factor_column)), size = 3) +
      geom_errorbar(aes(ymin = Lower, ymax = Upper, color = !!sym(factor_column)), width = 0.2) +
      theme(
        axis.text.x = element_text(angle = 90, vjust = 0.5),
        legend.position = "none"  # Remove the legend
      ) +
      labs(
        title = paste("Mean and 95% Confidence Interval for", variable, "-", group_name), 
        x = factor_column,
        y = "Value"
      ) +
      scale_y_continuous(limits = c(0, 100), breaks = seq(0, 100, by = 20))
    
    # Print the plot
    print(p)
    
    # Save the plot
    ggsave(filename = paste0("mean_ci_plot_", variable, "_", group_name, ".png"), plot = p, width = 10, height = 6)
  }
}

leader_col <- "Leader"

# Define the scales in the desired order
scales_scores <- c(
  "Leadership Skills",
  "Mindset & Attitude",
  "Skills, Mindset & Attitude",
  "Growth / Potential",
  "Leadership Effectiveness"
)

# Get all distinct values of the Leader column
distinct_leaders <- df_recoded_nooutliers %>% 
  distinct(Leader) %>%
  arrange(Leader)

# Print the distinct values
print(distinct_leaders)

df_recoded_nooutliers <- df_recoded_nooutliers %>%
  mutate(Leader = trimws(Leader))

# Define leader groups
engineering_leaders <- c(
  "Tim Ramey", "Shilpa Dabke", "Sai Naidu", "Anurag Singh", 
  "Amit Vasdev", "GeoffreyToustou", "Sam Evans"
)

product_leaders <- c(
  "Len Randazzo", "Ben Lattin", "Jonathan Flitt", "Jared Higley", 
  "Jimmy Jacob", "Trevor Headley", "Jonny Dittmer", "Summer Kisner",
  "Bharat Mabbu", "Samuel Hobbs", "Bryan Yip", "Sushant Jha", 
  "Shivanee Patel"
)

# Create the mean and 95% CI plots for each group
create_mean_ci_plots(df_recoded_nooutliers, scales_scores, leader_col, "Engineering and Info Sec", engineering_leaders)
create_mean_ci_plots(df_recoded_nooutliers, scales_scores, leader_col, "Product", product_leaders)




# Function to create column plots for each factor
create_column_plots <- function(data, scales_scores, factor_column) {
  # Create a long format of the data for ggplot
  long_data <- data %>%
    pivot_longer(cols = all_of(scales_scores), names_to = "Variable", values_to = "Value") %>%
    group_by(!!sym(factor_column), Variable) %>%
    summarise(
      Mean = round(mean(Value, na.rm = TRUE), 1),
      .groups = 'drop'
    )
  
  # Get unique levels of the factor column
  factor_levels <- unique(data[[factor_column]])
  
  # Loop through each factor level and create separate plots
  for (factor_level in factor_levels) {
    plot_data <- long_data %>% filter(!!sym(factor_column) == factor_level)
    
    p <- ggplot(plot_data, aes(x = Variable, y = Mean, fill = Variable)) +
      geom_col() +
      geom_text(aes(label = Mean), vjust = -0.3, size = 3.5) +
      labs(
        title = paste("Mean of Scales Scores for", factor_level), 
        
        y = "Mean"
      ) +
      theme(
        axis.text.x = element_text(vjust = 1),
        legend.position = "none"
      ) +
      scale_y_continuous(limits = c(0, 100), breaks = seq(0, 100, by = 20)) +
      scale_fill_brewer(palette = "Set3")
    
    # Print the plot
    print(p)
    
    # Save the plot with a short width
    ggsave(filename = paste0("column_plot_", factor_level, ".png"), plot = p, width = 6, height = 4)
  }
}

create_column_plots(df_recoded_nooutliers, scales_scores2, leader_col)



## DESCRIPTIVE TABLES

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    Lower_CI = numeric(), # Lower bound of the 95% CI
    Upper_CI = numeric(), # Upper bound of the 95% CI
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    ci_multiplier <- qt(0.975, df = length(na.omit(variable_data)) - 1) # 95% CI multiplier
    lower_ci <- mean_val - ci_multiplier * sem_val
    upper_ci <- mean_val + ci_multiplier * sem_val
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val,
      Lower_CI = lower_ci,
      Upper_CI = upper_ci
    ))
  }
  
  return(results)
}

df_descriptive_stats <- calculate_descriptive_stats(df_recoded_nooutliers, scales_scores)


calculate_means_and_sds_by_factors <- function(data, variable_names, factor_columns) {
  # Create an empty dataframe to store results
  results <- data.frame()
  
  # Create a new interaction factor that combines all factor columns
  data$interaction_factor <- interaction(data[factor_columns], drop = TRUE)
  
  # Iterate over each variable name
  for (variable_name in variable_names) {
    # Calculate means and SDs for each combination of factors for the variable
    aggregated_data <- aggregate(data[[variable_name]], 
                                 by = list(data$interaction_factor), 
                                 FUN = function(x) {
                                   n <- length(na.omit(x))
                                   mean_val <- mean(x, na.rm = TRUE)
                                   sd_val <- sd(x, na.rm = TRUE)
                                   sem_val <- sd_val / sqrt(n)
                                   ci_multiplier <- qt(0.975, df = n - 1)
                                   lower_ci <- mean_val - ci_multiplier * sem_val
                                   upper_ci <- mean_val + ci_multiplier * sem_val
                                   return(c(Mean = mean_val, SD = sd_val, Lower_CI = lower_ci, Upper_CI = upper_ci))
                                 })
    
    # Rename the columns and reshape the aggregated data
    names(aggregated_data)[1] <- "Factors"
    variable_results <- do.call(rbind, lapply(1:nrow(aggregated_data), function(i) {
      tibble(
        Factors = as.character(aggregated_data$Factors[i]),
        Variable = variable_name,
        Mean = aggregated_data$x[i, "Mean"],
        SD = aggregated_data$x[i, "SD"],
        Lower_CI = aggregated_data$x[i, "Lower_CI"],
        Upper_CI = aggregated_data$x[i, "Upper_CI"]
      )
    }))
    
    # Append results
    results <- bind_rows(results, variable_results)
  }
  
  # Split the interaction factor back into the original factors
  results <- results %>% 
    separate(Factors, into = factor_columns, sep = "\\.", remove = FALSE)
  
  return(results)
}

df_descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df_recoded_nooutliers, scales_scores, "Leader")


# HeatMap of leaders

library(ggplot2)
library(reshape2)

# Calculate means for each leader and scale
means <- df_recoded_nooutliers %>%
  group_by(Leader) %>%
  summarise(across(starts_with("Leadership Skills"):starts_with("Leadership Effectiveness"), mean, na.rm = TRUE))

# Reshape data for heatmap
means_melt <- melt(means, id.vars = "Leader")

# Create heatmap
ggplot(means_melt, aes(x = variable, y = Leader, fill = value)) +
  geom_tile() +
  geom_text(aes(label = round(value, 1)), color = "black", size = 3) +
  scale_fill_gradient(low = "white", high = "darkgreen") +
  labs(title = "Heatmap of Leadership Effectiveness Scores", x = "Scale", y = "Leader") +
  theme(axis.text.x = element_text(size = 10),
        axis.ticks.length = unit(-0.00, "cm"),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.text.y = element_text(size = 10),
        axis.title.y = element_text(size = 12))


# By Group

# Function to create heatmap

library(ggplot2)
library(reshape2)
library(dplyr)

create_heatmap <- function(data, group_name, leaders) {
  # Filter data for the specified leaders
  group_data <- data %>% filter(Leader %in% leaders)
  
  # Calculate means for each leader and scale
  means <- group_data %>%
    group_by(Leader) %>%
    summarise(across(starts_with("Leadership Skills"):starts_with("Leadership Effectiveness"), mean, na.rm = TRUE))
  
  # Reshape data for heatmap
  means_melt <- melt(means, id.vars = "Leader")
  
  # Create heatmap
  p <- ggplot(means_melt, aes(x = variable, y = Leader, fill = value)) +
    geom_tile() +
    geom_text(aes(label = round(value, 1)), color = "black", size = 3) +
    scale_fill_gradient(low = "white", high = "darkgreen") +
    labs(title = paste("Heatmap of Leadership Effectiveness Scores -", group_name), x = "Scale", y = "Leader") +
    theme(axis.text.x = element_text(size = 10),
          axis.ticks.length = unit(-0.00, "cm"),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          axis.text.y = element_text(size = 10),
          axis.title.y = element_text(size = 12))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = paste0("heatmap_", group_name, ".png"), plot = p, width = 10, height = 6)
}

# Create heatmaps for each group
create_heatmap(df_recoded_nooutliers, "Engineering and Info Sec", engineering_leaders)
create_heatmap(df_recoded_nooutliers, "Product", product_leaders)


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

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/trackrecord")

# Example usage
data_list <- list(
  "Cleaned Data" = df_recoded,
  "Reliability" = df_descriptives,
  "Validity" = df_validity,
  "Overall Descriptives" = df_descriptive_stats,
  "Descriptives - By Leader" = df_descriptive_stats_bygroup
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")



#Rename columns for radar charts
library(dplyr)

df_renamed <- df_recoded %>% 
  rename(
    `Company v self interest` = `How does this leader contribute to the growth of the company and prioritize the company's interests over their team's interests?`,
    `Effective team builder` = `Are they able to build, manage, and lead a team effectively to reach their goals?`,
    `Team demonstrates excellence and innovation` = `Do they create a mindset in their team / area of excellence, continuous improvement  and innovation?`,
    `Active development and promotion` = `Does this leader actively develop and promote their own people?`,
    `Effective at influencing and buy-in` = `How effective is this person in influencing others and gaining buy-in their area?`,
    `Infectiously passionate` = `How infectiously passionate is this leader when it comes to Clearwater and driving growth within their team?`,
    `Supportive` = `As you think about this person and when you need support, are they:`,
    `Ability to Adapt` = `Does this leader demonstrate the ability to adapt in pursuit of company goals or to new situations, e.g., Gen AI/ML, GO initiatives?`,
    `Ready for new responsibilities` = `When thinking about this person taking on additional leadership responsibilities to actively support growth, are they:`,
    `Contributes for opportunities for growth` = `Does this person actively contribute to finding and developing strategic opportunities for company growth?`,
    `Impact on team and organization` = `What has been the leader's impact on their team and organizational goals around driving growth?`,
    `Leader` = `Leader`
  )

# Define relevant columns
relevant_columns <- c(
  "Leader", 
  "Company v self interest", 
  "Effective team builder", 
  "Team demonstrates excellence and innovation", 
  "Active development and promotion", 
  "Effective at influencing and buy-in", 
  "Infectiously passionate", 
  "Supportive", 
  "Ability to Adapt", 
  "Ready for new responsibilities", 
  "Contributes for opportunities for growth", 
  "Impact on team and organization"
)

# Subset the dataframe to include only relevant columns
df_relevant <- df_renamed %>% select(all_of(relevant_columns))

# Rescale the values from 1-4 to 0-100
rescale_to_100 <- function(x) {
  return ((x - 1) / 3 * 100)
}

df_scaled <- df_relevant %>% mutate(across(-Leader, rescale_to_100))


# Calculate the average for each leader
df_avg <- df_scaled %>% group_by(Leader) %>% summarize(across(everything(), ~mean(.x, na.rm = TRUE)))


data_list <- list(
  "Radar Chart Data" = df_avg
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "DataRadarChart.xlsx")




# Extra datacut for Leadership Effectiveness


library(openxlsx)
df_employee_2 <- read.xlsx("Org Effectiveness and Succession Planning_Round 1 5.15.24 2 (1).xlsx", sheet = "Leader Eval Email v2")

df_recoded <- df_recoded %>%
  mutate(`Please add your email so that we know you have completed the survey. Your answers will remain confidential.` = 
           trimws(`Please add your email so that we know you have completed the survey. Your answers will remain confidential.`))

df_employee_2 <- df_employee_2 %>%
  mutate(Email = trimws(Email))

df_final_2 <- left_join(df_recoded, df_employee_2, by = c("Please add your email so that we know you have completed the survey. Your answers will remain confidential." = "Email"))




# Split the dataframe by the 'Area' column
df_split_by_area <- split(df_final_2, df_final_2$Area)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Sample_Size = integer(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    Lower_CI = numeric(), # Lower bound of the 95% CI
    Upper_CI = numeric(), # Upper bound of the 95% CI
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    n <- length(na.omit(variable_data))
    mean_val <- mean(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(n)
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    ci_multiplier <- qt(0.975, df = n - 1) # 95% CI multiplier
    lower_ci <- mean_val - ci_multiplier * sem_val
    upper_ci <- mean_val + ci_multiplier * sem_val
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Sample_Size = n,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val,
      Lower_CI = lower_ci,
      Upper_CI = upper_ci
    ))
  }
  
  return(results)
}

# Function to calculate means and SDs by factors
calculate_means_and_sds_by_factors <- function(data, variable_names, factor_columns) {
  # Create an empty dataframe to store results
  results <- data.frame()
  
  # Create a new interaction factor that combines all factor columns
  data$interaction_factor <- interaction(data[factor_columns], drop = TRUE)
  
  # Iterate over each variable name
  for (variable_name in variable_names) {
    # Calculate means and SDs for each combination of factors for the variable
    aggregated_data <- aggregate(data[[variable_name]], 
                                 by = list(data$interaction_factor), 
                                 FUN = function(x) {
                                   n <- length(na.omit(x))
                                   mean_val <- mean(x, na.rm = TRUE)
                                   sd_val <- sd(x, na.rm = TRUE)
                                   sem_val <- sd_val / sqrt(n)
                                   ci_multiplier <- qt(0.975, df = n - 1)
                                   lower_ci <- mean_val - ci_multiplier * sem_val
                                   upper_ci <- mean_val + ci_multiplier * sem_val
                                   return(c(Mean = mean_val, SD = sd_val, Lower_CI = lower_ci, Upper_CI = upper_ci, Sample_Size = n))
                                 })
    
    # Rename the columns and reshape the aggregated data
    names(aggregated_data)[1] <- "Factors"
    variable_results <- do.call(rbind, lapply(1:nrow(aggregated_data), function(i) {
      tibble(
        Factors = as.character(aggregated_data$Factors[i]),
        Variable = variable_name,
        Mean = aggregated_data$x[i, "Mean"],
        SD = aggregated_data$x[i, "SD"],
        Lower_CI = aggregated_data$x[i, "Lower_CI"],
        Upper_CI = aggregated_data$x[i, "Upper_CI"],
        Sample_Size = aggregated_data$x[i, "Sample_Size"]
      )
    }))
    
    # Append results
    results <- bind_rows(results, variable_results)
  }
  
  # Split the interaction factor back into the original factors
  results <- results %>% 
    separate(Factors, into = factor_columns, sep = "\\.", remove = FALSE)
  
  return(results)
}

# Apply the functions to each subset
desc_stats_by_area <- lapply(df_split_by_area, calculate_descriptive_stats, desc_vars = scales_scores)
means_sds_by_area <- lapply(df_split_by_area, calculate_means_and_sds_by_factors, variable_names = scales_scores, factor_columns = "Leader")

# Combine results into dataframes
df_desc_stats_combined <- bind_rows(lapply(names(desc_stats_by_area), function(area) {
  cbind(Area = area, desc_stats_by_area[[area]])
}))

df_means_sds_combined <- bind_rows(lapply(names(means_sds_by_area), function(area) {
  cbind(Area = area, means_sds_by_area[[area]])
}))


data_list <- list(
  "Split by Area - Overall" = df_desc_stats_combined,
  "Split by Area - by Leader" = df_means_sds_combined
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Leadership_SplitbyArea.xlsx")



# Split the dataframe by the 'Area' column
df_split_by_function <- split(df_final_2, df_final_2$Function)

# Apply the functions to each subset
desc_stats_by_area <- lapply(df_split_by_function, calculate_descriptive_stats, desc_vars = scales_scores)
means_sds_by_area <- lapply(df_split_by_function, calculate_means_and_sds_by_factors, variable_names = scales_scores, factor_columns = "Leader")

# Combine results into dataframes
df_desc_stats_combined_function <- bind_rows(lapply(names(desc_stats_by_area), function(area) {
  cbind(Area = area, desc_stats_by_area[[area]])
}))

df_means_sds_combined_function <- bind_rows(lapply(names(means_sds_by_area), function(area) {
  cbind(Area = area, means_sds_by_area[[area]])
}))

data_list <- list(
  "Split by Function - Overall" = df_desc_stats_combined_function,
  "Split by Function - by Leader" = df_means_sds_combined_function
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Leadership_SplitbyFunction.xlsx")