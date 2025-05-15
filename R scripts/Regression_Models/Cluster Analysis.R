setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/sulafabadi/Quarta Ordem")

library(readxl)
df <- read_xlsx('Dataset-4 market segmentation.xlsx', sheet = 'Clean Data')

# Get rid of special characters

names(df) <- gsub(" ", "_", names(df))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))

df$index <- 1:nrow(df)

# Recode Categorical Variables

# Load necessary library
library(dplyr)
library(tidyr)

# Recoding the variables
df_recoded <- df %>%
  mutate(
    Residency = case_when(
      `Where_do_you_live` == 1 ~ "Abu Dhabi",
      `Where_do_you_live` == 2 ~ "Dubai",
      `Where_do_you_live` == 3 ~ "Sharjah",
      `Where_do_you_live` == 4 ~ "Ajman",
      `Where_do_you_live` == 5 ~ "Umm Al Quwain",
      `Where_do_you_live` == 6 ~ "Ras Al Khaimah",
      `Where_do_you_live` == 7 ~ "Fujairah",
      TRUE ~ NA_character_
    ),
    Gender = case_when(
      `What_is_your_gender` == 1 ~ "Male",
      `What_is_your_gender` == 2 ~ "Female",
      TRUE ~ NA_character_
    ),
    Status_UAE = case_when(
      `What_is_your_current_residency_status_in_the_UAE` == 1 ~ "Local Resident (UAE citizen)",
      `What_is_your_current_residency_status_in_the_UAE` == 2 ~ "Expatriate Resident (non-UAE citizen)",
      `What_is_your_current_residency_status_in_the_UAE` == 3 ~ "Tourist",
      TRUE ~ NA_character_
    ),
    Occupation = case_when(
      `What_is_your_occupation` == 1 ~ "Student",
      `What_is_your_occupation` == 2 ~ "Working (full-time)",
      `What_is_your_occupation` == 3 ~ "Working (part-time)",
      `What_is_your_occupation` == 4 ~ "Self-employed/ Business owner",
      `What_is_your_occupation` == 5 ~ "Unemployed",
      `What_is_your_occupation` == 6 ~ "Retired",
      TRUE ~ NA_character_
    ),
    Age = case_when(
      `What_is_your_age` == 1 ~ "Under 20",
      `What_is_your_age` == 2 ~ "20–29",
      `What_is_your_age` == 3 ~ "30–39",
      `What_is_your_age` == 4 ~ "40–49",
      `What_is_your_age` == 5 ~ "50–59",
      `What_is_your_age` == 6 ~ "60 and above",
      TRUE ~ NA_character_
    ),
    Education = case_when(
      `What_is_the_highest_level_of_education_that_you_have_completed` == 1 ~ "Less than high school",
      `What_is_the_highest_level_of_education_that_you_have_completed` == 2 ~ "High school diploma or equivalent",
      `What_is_the_highest_level_of_education_that_you_have_completed` == 3 ~ "Bachelor’s degree",
      `What_is_the_highest_level_of_education_that_you_have_completed` == 4 ~ "Master’s degree",
      `What_is_the_highest_level_of_education_that_you_have_completed` == 5 ~ "Doctoral degree",
      TRUE ~ NA_character_
    ),
    Household_Income = case_when(
      `What_is_your_households_annual_income` == 1 ~ "Less than 50,000 AED",
      `What_is_your_households_annual_income` == 2 ~ "51,000 – 100,000 AED",
      `What_is_your_households_annual_income` == 3 ~ "101,000 – 150,000 AED",
      `What_is_your_households_annual_income` == 4 ~ "151,000 – 200,000 AED",
      `What_is_your_households_annual_income` == 5 ~ "201,000 – 250,000 AED",
      `What_is_your_households_annual_income` == 6 ~ "More than 250,000 AED",
      TRUE ~ NA_character_
    ),
    Car_Ownership = case_when(
      `How_many_cars_do_you_own_in_your_household` == 1 ~ "No car",
      `How_many_cars_do_you_own_in_your_household` == 2 ~ "1 car",
      `How_many_cars_do_you_own_in_your_household` == 3 ~ "2 cars",
      `How_many_cars_do_you_own_in_your_household` == 4 ~ "3 cars",
      `How_many_cars_do_you_own_in_your_household` == 5 ~ "4 cars or more")
  )


# Scale Construction

library(psych)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SE_of_the_Mean = numeric(),
    StDev = numeric(),
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val <- alpha_results$total$raw_alpha
    
    # Saving item-total correlations in the environment
    itc_name <- paste("itc_", scale, sep = "")
    assign(itc_name, alpha_results$item.stats[,"raw.r"], envir = .GlobalEnv)
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = mean(item_data, na.rm = TRUE),
        SE_of_the_Mean = sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data))),
        StDev = sd(item_data, na.rm = TRUE),
        Alpha = NA
      ))
    }
    
    # Calculate the mean score for the scale and add as a new column
    scale_mean <- rowMeans(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_mean
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = mean(scale_mean, na.rm = TRUE),
      SE_of_the_Mean = sd(scale_mean, na.rm = TRUE) / sqrt(sum(!is.na(scale_mean))),
      StDev = sd(scale_mean, na.rm = TRUE),
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}


scales <- list(
  "Pro-technology" = c("PRT1", "PRT2", "PRT3", "PRT4"),
  "Pro-environment" = c("PRE1", "PRE2", "PRE3", "PRE4"),
  "Public transport attitude" = c("PTA1", "PTA2"),
  "Performance expectancy" = c("PE1", "PE2", "PE3", "PE4", "PE5"),
  "Effort expectancy" = c("EE1", "EE2", "EE3", "EE4"),
  "Social Influence" = c("SI1", "SI2", "SI3"),
  "Facilitating Conditions" = c("FC1", "FC2", "FC3", "FC4"),
  "Hedonic Motivation" = c("HM1", "HM2", "HM3"),
  "Price value" = c("PV1", "PV2", "PV3", "PV4"),
  "Perceived risks" = c("PR1", "PR2", "PR3", "PR4"),
  "Behavioural Intention" = c("BI1", "BI2", "BI3", "BI4")
)


alpha_results <- reliability_analysis(df_recoded, scales)

df_recoded <- alpha_results$data_with_scales
df_descriptives <- alpha_results$statistics


categorical_vars_cluster <- c("Residency", "Gender", "Status_UAE", "Occupation")
numerical_vars <- c("Pro-technology",                                                             
                    "Pro-environment" ,                                                           
                    "Public transport attitude"  ,                                                
                    "Performance expectancy"   ,                                                  
                    "Effort expectancy"       ,                                                   
                    "Social Influence"        ,                                                   
                    "Facilitating Conditions" ,                                                   
                    "Hedonic Motivation"      ,                                                   
                    "Price value"             ,                                                   
                    "Perceived risks"         ,                                                   
                    "Behavioural Intention"  )

# Convert to factors

for (var in categorical_vars_cluster) {
  # Check if the variable exists in the dataframe
  if (var %in% names(df_recoded)) {
    # Convert the variable to a factor, keeping NA as a level if they exist
    df_recoded[[var]] <- factor((df_recoded)[[var]], exclude = NULL)
  } else {
    # Print a message if the variable is not found in the dataframe
    message(sprintf("Variable '%s' not found in the dataframe.", var))
  }
}

## BOXPLOTS OF MULTIPLE VARIABLES
library(ggplot2)
library(tidyr)
library(stringr)

create_boxplots <- function(df, vars) {
  # Ensure that vars are in the dataframe
  df <- df[, vars, drop = FALSE]
  
  # Reshape the data to a long format
  long_df <- df %>%
    gather(key = "Variable", value = "Value")
  
  # Create side-by-side boxplots for each variable
  p <- ggplot(long_df, aes(x = Variable, y = Value)) +
    geom_boxplot() +
    stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = "Side-by-Side Boxplots", x = "Variable", y = "Value") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1)) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 10))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = "side_by_side_boxplots.png", plot = p, width = 10, height = 6)
}

create_boxplots(df_recoded, numerical_vars)


# Frequency Table

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

categorical_vars_freqtable <- c("Residency", "Gender", "Status_UAE", "Occupation",
                                "Age"   ,                                                                     
                                "Education"   ,                                                               
                                "Household_Income"    ,                                                       
                                "Car_Ownership")

freq_df <- create_frequency_tables(df_recoded, categorical_vars_freqtable)

#Recode categories

df_recoded <- df_recoded %>%
  mutate(Occupation = case_when(
    Occupation == 'Retired' ~ 'Retired or Student',  # Code for 'Retired'
    Occupation == 'Student' ~ 'Retired or Student',  # Code for 'Student'
    TRUE ~ Occupation     # Keep other categories as they are
  ))



freq_df_2 <- create_frequency_tables(df_recoded, categorical_vars_freqtable)


## CLUSTERING ALGORITHM - HIERARCHICAL PHASE

library(cluster)
library(factoextra)
library(dplyr)
library(ggplot2)

perform_hierarchical_clustering <- function(df, vars) {
  # Select variables
  data_selected <- df %>% select(all_of(vars))
  
  # Handling mixed data types
  dissimilarity_matrix <- daisy(data_selected)
  
  # Hierarchical clustering
  hc <- hclust(dissimilarity_matrix, method = "ward.D2")
  
  # Initialize lists to store results
  elbow_results <- list()
  sil_results <- list()
  gap_results <- list()
  
  for (k in 2:10) {
    # Cut the dendrogram to get cluster assignments
    cluster_assignments <- cutree(hc, k)
    
    # Calculate centroids
    centroids <- aggregate(data_selected, by = list(cluster_assignments), FUN = mean)
    centroids <- centroids[, -1]  # Remove the cluster assignment column
    
    # Check if the number of centroids matches k
    if (nrow(centroids) != k) {
      stop("The number of centroids does not match the number of clusters.")
    }
    
    # Elbow Method
    wss <- sum(kmeans(data_selected, centers = centroids, nstart = 1, iter.max = 50)$withinss)
    
    # Store results for Elbow Method
    elbow_results[[k]] <- data.frame(Number_of_Clusters = k, Total_withinss = wss)
    
    # Silhouette analysis
    sil_width <- mean(silhouette(cluster_assignments, dissimilarity_matrix)[, 'sil_width'])
    
    # Store results for Silhouette Width
    sil_results[[k]] <- data.frame(Number_of_Clusters = k, Average_Silhouette_Width = sil_width)
    
    # Gap Statistic
    set.seed(123)
    gap_stat <- clusGap(data_selected, FUN = kmeans, K.max = 10, B = 50)
    
    # Store results for Gap Statistic
    gap_results[[k]] <- data.frame(Number_of_Clusters = 1:10, Gap_Stat = gap_stat$Tab[, "gap"])
  }
  
  # Combine results into dataframes
  elbow_df <- do.call(rbind, elbow_results)
  sil_df <- do.call(rbind, sil_results)
  gap_df <- do.call(rbind, gap_results)
  
  # Elbow plot
  elbow_plot <- ggplot(elbow_df, aes(x = Number_of_Clusters, y = Total_withinss)) +
    geom_line() +
    geom_point() +
    theme_minimal() +
    labs(title = "Elbow Method for Optimal Number of Clusters",
         x = "Number of Clusters", y = "Total Within-Cluster Sum of Squares")
  
  # Silhouette plot
  sil_plot <- ggplot(sil_df, aes(x = Number_of_Clusters, y = Average_Silhouette_Width)) +
    geom_line() +
    geom_point() +
    theme_minimal() +
    labs(title = "Average Silhouette Widths for Different Numbers of Clusters",
         x = "Number of Clusters", y = "Average Silhouette Width")
  
  # Gap Statistic plot
  gap_plot <- fviz_gap_stat(gap_stat)
  
  # Return results
  return(list(
    Hierarchical_Clustering = hc,
    Elbow_Plot = elbow_plot,
    Silhouette_Plot = sil_plot,
    Gap_Stat_Plot = gap_plot,
    Elbow_DataFrame = elbow_df,
    Silhouette_DataFrame = sil_df,
    Gap_Stat_DataFrame = gap_df
  ))
}

    
#Dummy categorical columns

library(dplyr)
library(tidyr)
library(fastDummies)

clustering_vars <- c(                                               
  "Gender_Female"   ,                                                           
  "Gender_Male"      ,                                                          
  "Status_UAE_Expatriate_Resident__non_UAE_citizen_" ,                          
  "Status_UAE_Local_Resident__UAE_citizen_"        ,                            
  "Status_UAE_Tourist"                              ,                           
  "Status_UAE_NA"                                    ,                          
  "Occupation_Self_employed__Business_owner"           ,                        
  "Occupation_Unemployed"                               ,                       
  "Occupation_Working__full_time_"                       ,                      
  "Occupation_Working__part_time_",
  "What_is_the_highest_level_of_education_that_you_have_completed",
  "What_is_your_households_annual_income"                   ,                 
  "How_many_cars_do_you_own_in_your_household"                ,                
  "How_frequently_do_you_go_out_for_shopping_activities_per_month" ,           
  "How_frequently_do_you_go_out_for_leisure_and_fun_activities_per_month"  ,   
  "How_frequently_do_you_go_out_for_social_visits_per_month"                ,  
  "How_frequently_do_you_walk__cycle__or_use_a_scooter_per_month"            , 
  "How_frequently_do_you_use_your_private_car_per_month"                      ,
  "How_frequently_do_you_use_public_transport__bus__tram__or_metro__per_month",
  "How_frequently_do_you_use_a_taxi_per_month")

df_recoded <- dummy_cols(df_recoded, select_columns = categorical_vars_cluster)

# Function to normalize (scale) variables
normalize_variables_min_max <- function(df, variables) {
  df %>%
    mutate(across(all_of(variables), 
                  ~ (.-min(.))/(max(.) - min(.)), 
                  .names = "{.col}_norm"))
}

vars_to_normalize <- c("What_is_the_highest_level_of_education_that_you_have_completed",
                       "What_is_your_households_annual_income",
                       "How_many_cars_do_you_own_in_your_household",
                       "How_frequently_do_you_go_out_for_shopping_activities_per_month"   ,         
                       "How_frequently_do_you_go_out_for_leisure_and_fun_activities_per_month"   ,  
                       "How_frequently_do_you_go_out_for_social_visits_per_month"  ,                
                       "How_frequently_do_you_walk__cycle__or_use_a_scooter_per_month"  ,           
                       "How_frequently_do_you_use_your_private_car_per_month"     ,                 
                       "How_frequently_do_you_use_public_transport__bus__tram__or_metro__per_month",
                       "How_frequently_do_you_use_a_taxi_per_month" )


normalized_df <- normalize_variables_min_max(df_recoded, vars_to_normalize)

names(normalized_df) <- gsub(" ", "_", names(normalized_df))
names(normalized_df) <- gsub("\\(", "_", names(normalized_df))
names(normalized_df) <- gsub("\\)", "_", names(normalized_df))
names(normalized_df) <- gsub("\\-", "_", names(normalized_df))
names(normalized_df) <- gsub("/", "_", names(normalized_df))
names(normalized_df) <- gsub("\\\\", "_", names(normalized_df)) 
names(normalized_df) <- gsub("\\?", "", names(normalized_df))
names(normalized_df) <- gsub("\\'", "", names(normalized_df))
names(normalized_df) <- gsub("\\,", "_", names(normalized_df))

clustering_vars <- c(                                               
  "Gender_Female"   ,                                                           
  "Gender_Male"      ,                                                          
  "Status_UAE_Expatriate_Resident__non_UAE_citizen_" ,                          
  "Status_UAE_Local_Resident__UAE_citizen_"        ,                            
  "Status_UAE_Tourist"                              ,                           
  "Status_UAE_NA"                                    ,                          
  "Occupation_Self_employed__Business_owner"           ,                        
  "Occupation_Unemployed"                               ,                       
  "Occupation_Working__full_time_"                       ,                      
  "Occupation_Working__part_time_",
  "What_is_the_highest_level_of_education_that_you_have_completed",
  "What_is_your_households_annual_income"                   ,                 
  "How_many_cars_do_you_own_in_your_household"                ,                
  "How_frequently_do_you_go_out_for_shopping_activities_per_month" ,           
  "How_frequently_do_you_go_out_for_leisure_and_fun_activities_per_month"  ,   
  "How_frequently_do_you_go_out_for_social_visits_per_month"                ,  
  "How_frequently_do_you_walk__cycle__or_use_a_scooter_per_month"            , 
  "How_frequently_do_you_use_your_private_car_per_month"                      ,
  "How_frequently_do_you_use_public_transport__bus__tram__or_metro__per_month",
  "How_frequently_do_you_use_a_taxi_per_month")


clustering_results <- perform_hierarchical_clustering(normalized_df, clustering_vars)

# View the plots
print(clustering_results$Elbow_Plot)
print(clustering_results$Silhouette_Plot)
print(clustering_results$Gap_Stat_Plot)

# View dataframes and results
Elbow_df <- clustering_results$Elbow_DataFrame
Silhouette_df <- clustering_results$Silhouette_DataFrame
GapStat_df <- clustering_results$Gap_Stat_DataFrame


# Correlation Analysis
calculate_correlation_matrix <- function(data, variables, method = "pearson") {
  # Ensure the method is valid
  if (!method %in% c("pearson", "spearman")) {
    stop("Method must be either 'pearson' or 'spearman'")
  }
  
  # Subset the data for the specified variables
  data_subset <- data[variables]
  
  # Calculate the correlation matrix
  corr_matrix <- cor(data_subset, method = method, use = "complete.obs")
  
  # Initialize a matrix to hold formatted correlation values
  corr_matrix_formatted <- matrix(nrow = nrow(corr_matrix), ncol = ncol(corr_matrix))
  rownames(corr_matrix_formatted) <- colnames(corr_matrix)
  colnames(corr_matrix_formatted) <- colnames(corr_matrix)
  
  # Calculate p-values and apply formatting with stars
  for (i in 1:nrow(corr_matrix)) {
    for (j in 1:ncol(corr_matrix)) {
      if (i == j) {  # Diagonal elements (correlation with itself)
        corr_matrix_formatted[i, j] <- format(round(corr_matrix[i, j], 3), nsmall = 3)
      } else {
        p_value <- cor.test(data_subset[[i]], data_subset[[j]], method = method)$p.value
        stars <- ifelse(p_value < 0.01, "***", 
                        ifelse(p_value < 0.05, "**", 
                               ifelse(p_value < 0.1, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}


correlation_matrix <- calculate_correlation_matrix(normalized_df, clustering_vars, method = "spearman")


## K-MEANS CLUSTERING to make final decision on number of clusters

library(dplyr)
library(cluster)
library(broom)

run_hierarchical_kmeans <- function(df, clustering_vars, k_range) {
  # Create a dissimilarity matrix
  dissimilarity_matrix <- daisy(df[clustering_vars])
  
  # Hierarchical clustering
  hc <- hclust(dissimilarity_matrix, method = "ward.D2")
  
  results <- list()
  
  for (k in k_range) {
    # Cut the dendrogram to get cluster assignments
    cluster_assignments <- cutree(hc, k)
    
    # Calculate centroids
    centroids <- aggregate(df[clustering_vars], by = list(cluster_assignments), FUN = mean)
    
    # Check if the number of centroids matches k
    if (nrow(centroids) != k) {
      stop("The number of centroids does not match the number of clusters.")
    }

    # Remove the cluster assignment column
    centroids <- centroids[, -1]
    
    # Run K-means with the centroids from hierarchical clustering as initial seeds
    set.seed(123)  # For reproducibility
    kmeans_result <- kmeans(df[clustering_vars], centers = centroids, nstart = 1)
    
    # Add cluster assignments to the dataframe
    cluster_col_name <- paste("Cluster_", k, sep = "")
    df[[cluster_col_name]] <- kmeans_result$cluster
    
    results[[cluster_col_name]] <- df
  }
  
  # Return the list of dataframes with added cluster assignments
  return(results)
}

clustering_vars <- c(                                               
  "Gender_Female"   ,                                                           
  "Gender_Male"      ,                                                          
  "Status_UAE_Expatriate_Resident__non_UAE_citizen_" ,                          
  "Status_UAE_Local_Resident__UAE_citizen_"        ,                            
  "Status_UAE_Tourist"                              ,                           
  "Status_UAE_NA"                                    ,                          
  "Occupation_Self_employed__Business_owner"           ,                        
  "Occupation_Unemployed"                               ,                       
  "Occupation_Working__full_time_"                       ,                      
  "Occupation_Working__part_time_",
  "What_is_the_highest_level_of_education_that_you_have_completed",
  "What_is_your_households_annual_income"                   ,                 
  "How_many_cars_do_you_own_in_your_household"                ,                
  "How_frequently_do_you_go_out_for_shopping_activities_per_month" ,           
  "How_frequently_do_you_go_out_for_leisure_and_fun_activities_per_month"  ,   
  "How_frequently_do_you_go_out_for_social_visits_per_month"                ,  
  "How_frequently_do_you_walk__cycle__or_use_a_scooter_per_month"            , 
  "How_frequently_do_you_use_your_private_car_per_month"                      ,
  "How_frequently_do_you_use_public_transport__bus__tram__or_metro__per_month",
  "How_frequently_do_you_use_a_taxi_per_month")


k_range <- 4:6
df_with_clusters_list <- run_hierarchical_kmeans(normalized_df, clustering_vars, k_range)
df_withclusters <- df_with_clusters_list$Cluster_6

# Run ANOVA

library(broom)

run_anova <- function(df, grouping_var, dv_list) {
  anova_results <- lapply(dv_list, function(dv) {
    # Run ANOVA
    anova_model <- aov(reformulate(grouping_var, response = dv), data = df)
    # Extract and tidy the ANOVA results
    tidy_anova <- tidy(anova_model) %>% filter(term == grouping_var)
    # Add the name of the dependent variable
    tidy_anova$Dependent_Variable <- dv
    return(tidy_anova)
  })
  
  # Combine results into a single dataframe
  combined_anova_results <- do.call(rbind, anova_results) %>%
    select(Dependent_Variable, sumsq, meansq, statistic, p.value)
  
  return(combined_anova_results)
}

anova_results_4clusters <- run_anova(df_withclusters, "Cluster_4", clustering_vars)
anova_results_5clusters <- run_anova(df_withclusters, "Cluster_5", clustering_vars)
anova_results_6clusters <- run_anova(df_withclusters, "Cluster_6", clustering_vars)

compiled_results <- list(anova_results_4clusters, anova_results_5clusters, anova_results_6clusters)

summarize_anova_tables <- function(anova_tables) {
  summary_results <- lapply(anova_tables, function(anova_table) {
    # Ensure anova_table is treated as a dataframe
    anova_df <- as.data.frame(anova_table)
    
    # Sum of F-statistics
    f_stat_sum <- sum(anova_df$statistic, na.rm = TRUE)
    
    # Count of significant p-values
    significant_count <- sum(anova_df$p.value < 0.05, na.rm = TRUE)
    
    return(data.frame(F_Stat_Sum = f_stat_sum, Significant_Count = significant_count))
  })
  
  # Combine results into a single dataframe
  combined_summary_results <- do.call(rbind, summary_results)
  return(combined_summary_results)
}

summary_ANOVA_results <- summarize_anova_tables(compiled_results)

# 4-clusters is the optimal solution - Cluster PROFILING

categories_clustervars <- c("Residency"  ,                                                               
                      "Gender"    ,                                                                
                      "Status_UAE"  ,                                                              
                      "Occupation"    ,                                                            
                      "Age"            ,                                                           
                      "Education"          ,                                                       
                      "Household_Income"   ,                                                       
                      "Car_Ownership")

numbers_clustervars <- c("How_frequently_do_you_go_out_for_shopping_activities_per_month"          ,  
                            "How_frequently_do_you_go_out_for_leisure_and_fun_activities_per_month"    , 
                            "How_frequently_do_you_go_out_for_social_visits_per_month"                  ,
                            "How_frequently_do_you_walk__cycle__or_use_a_scooter_per_month"             ,
                            "How_frequently_do_you_use_your_private_car_per_month"                      ,
                            "How_frequently_do_you_use_public_transport__bus__tram__or_metro__per_month",
                            "How_frequently_do_you_use_a_taxi_per_month")

categories_nonclustervars <- c("Country_Code",
                               "Region",
                               "Where_do_you_live")

numbers_nonclustervars <- c("Pro_technology"              ,                                              
                            "Pro_environment"              ,                                             
                            "Public_transport_attitude"     ,                                            
                            "Performance_expectancy"         ,                                           
                            "Effort_expectancy"               ,                                          
                            "Social_Influence"                 ,                                         
                            "Facilitating_Conditions"           ,                                        
                            "Hedonic_Motivation"                 ,                                       
                            "Price_value"                         ,                                      
                            "Perceived_risks"                     ,                                      
                            "Behavioural_Intention")

clusterassignmentcolumn <- "Cluster_4"

# Function to calculate cluster sizes
calculate_cluster_sizes <- function(df, cluster_column) {
  cluster_sizes <- table(df[[cluster_column]])
  return(as.data.frame(cluster_sizes))
}

cluster_sizes_df <- calculate_cluster_sizes(df_withclusters, clusterassignmentcolumn)

# Function to calculate descriptives by cluster

calculate_means_and_sds_by_factors <- function(data, variables, factors) {
  # Create an empty list to store intermediate results
  results_list <- list()
  
  # Iterate over each variable
  for (var in variables) {
    # Create a temporary data frame to store results for this variable
    temp_results <- data.frame(Variable = var)
    
    # Iterate over each factor
    for (factor in factors) {
      # Aggregate data by factor
      agg_data <- aggregate(data[[var]], by = list(data[[factor]]), 
                            FUN = function(x) c(Mean = mean(x, na.rm = TRUE), SD = sd(x, na.rm = TRUE)))
      
      # Create columns for each level of the factor
      for (level in unique(data[[factor]])) {
        level_agg_data <- agg_data[agg_data[, 1] == level, ]
        
        if (nrow(level_agg_data) > 0) {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- level_agg_data$x[1, "Mean"]
          temp_results[[paste0(factor, "_", level, "_SD")]] <- level_agg_data$x[1, "SD"]
        } else {
          temp_results[[paste0(factor, "_", level, "_Mean")]] <- NA
          temp_results[[paste0(factor, "_", level, "_SD")]] <- NA
        }
      }
    }
    
    # Add the results for this variable to the list
    results_list[[var]] <- temp_results
  }
  
  # Combine all the results into a single dataframe
  descriptive_stats_bygroup <- do.call(rbind, results_list)
  return(descriptive_stats_bygroup)
}

factors <- c("Cluster_4")
descriptive_clustervars_bycluster <- calculate_means_and_sds_by_factors(df_withclusters, numbers_clustervars, factors)
descriptive_nonclustervars_bycluster <- calculate_means_and_sds_by_factors(df_withclusters, numbers_nonclustervars, factors)


create_segmented_frequency_tables <- function(data, vars, segmenting_categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each segmenting category
  for (segment in segmenting_categories) {
    # Iterate over each variable for which frequencies are to be calculated
    for (var in vars) {
      # Ensure both the segmenting category and variable are treated as factors
      segment_factor <- factor(data[[segment]])
      var_factor <- factor(data[[var]])
      
      # Calculate counts segmented by the segmenting category
      counts <- table(segment_factor, var_factor)
      
      # Melt the table into a long format for easier handling
      freq_table <- as.data.frame(counts)
      names(freq_table) <- c("Segment", "Level", "Count")
      
      # Add the variable name and calculate percentages
      freq_table$Variable <- var
      freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
      
      # Add the result to the list
      all_freq_tables[[paste(segment, var, sep = "_")]] <- freq_table
    }
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


segmented_freq_df <- create_segmented_frequency_tables(df_withclusters, categories_clustervars, "Cluster_4")


# PLOTS

library(ggplot2)
library(dplyr)
library(tidyr)

create_dot_and_whiskers_plot <- function(df, cluster_var, vars_to_plot) {
  # Reshape the data to long format for easier plotting
  long_df <- df %>%
    select(all_of(c(cluster_var, vars_to_plot))) %>%
    pivot_longer(cols = vars_to_plot, names_to = "Variable", values_to = "Value")
  
  # Calculate means and confidence intervals
  mean_ci_df <- long_df %>%
    group_by(Cluster = .data[[cluster_var]], Variable) %>%
    summarise(
      Mean = mean(Value, na.rm = TRUE),
      Lower_CI = Mean - qnorm(0.975) * sd(Value, na.rm = TRUE) / sqrt(n()),
      Upper_CI = Mean + qnorm(0.975) * sd(Value, na.rm = TRUE) / sqrt(n()),
      .groups = 'drop'
    )
  
  # Create the dot and whiskers plot
  ggplot(mean_ci_df, aes(x = Variable, y = Mean, group = Cluster, color = as.factor(Cluster))) +
    geom_point(position = position_dodge(width = 0.75)) +
    geom_errorbar(aes(ymin = Lower_CI, ymax = Upper_CI), 
                  width = 0.2, position = position_dodge(width = 0.75)) +
    theme_minimal() +
    labs(title = "Dot and Whiskers Plot of Means by Cluster",
         x = "Variable", y = "Mean") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1))
}

# Usage
vars_to_plot1 <- c("Pro_technology", "Pro_environment", "Public_transport_attitude", 
                  "Performance_expectancy", "Effort_expectancy")
vars_to_plot2 <- c("Social_Influence", 
                   "Facilitating_Conditions", "Hedonic_Motivation", "Price_value", 
                   "Perceived_risks", "Behavioural_Intention")

dot_whiskers_plot <- create_dot_and_whiskers_plot(df_withclusters, "Cluster_4", vars_to_plot1)
dot_whiskers_plot2 <- create_dot_and_whiskers_plot(df_withclusters, "Cluster_4", vars_to_plot2)
# Display the plot
print(dot_whiskers_plot)
print(dot_whiskers_plot2)

# ANOVA and POSTHOCS for differences

library(dplyr)
library(broom)
library(purrr)
library(stats)

perform_anova_and_posthoc <- function(df, variables, cluster_var) {
  results <- list()
  
  for (var in variables) {
    # Ensure the cluster variable is correctly referenced
    cluster_factor <- factor(df[[cluster_var]])
    
    # Run ANOVA
    anova_result <- aov(reformulate(cluster_var, response = var), data = df)
    tidy_anova <- tidy(anova_result)
    anova_F <- tidy_anova$statistic[1]
    anova_p <- tidy_anova$p.value[1]
    
    # If ANOVA is significant, run post-hoc pairwise t-tests
    if (anova_p < 0.05) {
      pairwise_results <- pairwise.t.test(df[[var]], cluster_factor, 
                                          p.adjust.method = "bonferroni")
      
      # Tidy up the pairwise results and add to the output
      tidy_pairwise <- as.data.frame(pairwise_results$p.value)
      colnames(tidy_pairwise) <- gsub(" ", "_", colnames(tidy_pairwise))
      tidy_pairwise$Variable <- var
      tidy_pairwise$ANOVA_F <- anova_F
      tidy_pairwise$ANOVA_p <- anova_p
      
      results[[var]] <- tidy_pairwise
    } else {
      # If ANOVA is not significant, add only ANOVA results
      results[[var]] <- data.frame(
        Variable = var,
        ANOVA_F = anova_F,
        ANOVA_p = anova_p,
        stringsAsFactors = FALSE
      )
    }
  }
  
  # Combine all results into a single dataframe
  combined_results <- bind_rows(results, .id = "Variable")
  return(combined_results)
}

results_anova_clusters <- perform_anova_and_posthoc(df_withclusters, vars_to_plot, "Cluster_4")



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
  "Sample Characteristics" = freq_df_2,
  "Descriptives_Attitudinal" = df_descriptives,
  "Correlations" = correlation_matrix,
  "ANOVA Results - 4 clusters" - anova_results_4clusters,
  "ANOVA Results - 5 clusters" - anova_results_5clusters,
  "ANOVA Results - 6 clusters" - anova_results_6clusters,
  "Cluster Sizes" = cluster_sizes_df, 
  "Descriptives - Cluster Vars" = descriptive_clustervars_bycluster,
  "Descriptives - Non-Cluster Vars" = descriptive_nonclustervars_bycluster,
  "Frequencies - Clusters" = segmented_freq_df,
  "ANOVA - Clusters" = results_anova_clusters
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")