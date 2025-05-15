setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/willtregoning")

library(haven)
df <- read_sav("DrugLawReformAudienceSegmentationFINAL_NewVars.sav")

vars_demo <- c("Q2H",
               "Q4" ,
               "Q1H",
               "Q5",
               "Q6",
               "Q7B",
               "Q8B",
               "Q9B",
               "Q11",
               "Q12AB",
               "Q15A",
               "Q16",
               "Q17",
               "Q19",
               "Q20"
               
)

vars_drugsupport <- c("Q21_1"      ,         "Q21_2"        ,       "Q21_3"        ,       "Q21_4"      ,         "Q21_5"   ,           
                      "Q21_6")

vars_supportpolicies <- c("Q22_1"     ,          "Q22_2")

vars_attitudestodrugs <- c("Q23_1"     ,          "Q23_2"      ,         "Q23_3"    ,          
                           "Q23_4"    ,          "Q23_5"      ,         "Q23_6"    ,           "Q23_7"   ,            "Q23_8"   ,            "Q23_9"  ,            
                           "Q23_10"       ,       "Q23_11")

vars_comparison <- c("Q14_1"    ,          
                     "Q14_2"        ,       "Q14_3"   ,            "Q14_4"    ,           "Q14_5"           ,    "Q14_6"      ,         "Q14_7"       ,       
                     "Q14_8"       ,        "Q14_9"       ,        "Q14_10" ,            "Q14_11"    ,          "Q14_12"     ,         "Q14_13",             
                     "Q14_14"        ,      "Q14_15"      ,        "Q14_16"     ,         "Q14_17"       ,       "Q14_18"        ,      "Q14_19")


df[vars_demo] <- lapply(df[vars_demo], as_factor)

# Descriptive Tables

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
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

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


df_descriptive_drugsupport <- calculate_descriptive_stats(df, vars_drugsupport)
df_descriptive_policies <- calculate_descriptive_stats(df, vars_supportpolicies)
df_descriptive_policies <- calculate_descriptive_stats(df, vars_attitudestodrugs)
df_freq_demo <- create_frequency_tables(df, vars_demo)

# Recode categories

library(dplyr)
library(forcats)

df <- df %>%
  mutate(Q11 = fct_recode(Q11,
                          "Other" = "Carer (unpaid)",
                          "Other" = "Full-time student"))

df_freq_demo2 <- create_frequency_tables(df, vars_demo)

# Reliability Tests

## RELIABILITY ANALYSIS

# Scale Construction

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SE_of_the_Mean = numeric(),
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
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = mean(item_data, na.rm = TRUE),
        SE_of_the_Mean = sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data))),
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
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
      ITC = item_itc,
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}

# Reverse Score

reverse_scores_multiple_columns <- function(df, column_names) {
  # Ensure all column names exist in the dataframe
  missing_columns <- setdiff(column_names, names(df))
  if (length(missing_columns) > 0) {
    stop("The following column names do not exist in the dataframe: ", paste(missing_columns, collapse=", "))
  }
  
  # Iterate over each column name, calculate reversed scores, and add as a new column
  for (column_name in column_names) {
    new_column_name <- paste0(column_name, "_rev")
    df[[new_column_name]] <- 10 - df[[column_name]]
  }
  
  return(df)
}

df <- reverse_scores_multiple_columns(df, c("Q23_1"     , "Q23_3"    ,          
                                            "Q23_4"    ,       "Q23_6"    ,           "Q23_7"   ,        "Q23_9"  ,            
                                            "Q23_10"       ,       "Q23_11"))

vars_attitudestodrugs_rev <- c("Q23_1_rev"     ,          "Q23_2"      ,         "Q23_3_rev"    ,          
                               "Q23_4_rev"    ,          "Q23_5"      ,         "Q23_6_rev"    ,           "Q23_7_rev"   ,            "Q23_8"   ,            "Q23_9_rev"  ,            
                               "Q23_10_rev"       ,       "Q23_11_rev")

scales1 <- list(
  "Attitudes_to_Drugs" = vars_attitudestodrugs_rev
)

alpha_results1 <- reliability_analysis(df, scales1)
df_reliability_attitudes <- alpha_results1$statistics
df <- alpha_results1$data_with_scales

scales2 <- list(
  "Drug_Support" = vars_drugsupport
)

alpha_results2 <- reliability_analysis(df, scales2)
df_reliability_drugsupport <- alpha_results2$statistics
df <- alpha_results2$data_with_scales

scales3 <- list(
  "Policy_Support" = vars_supportpolicies
)

alpha_results3 <- reliability_analysis(df, scales3)
df_reliability_policysupport <- alpha_results3$statistics
df <- alpha_results3$data_with_scales

df_descriptives_compscales <- calculate_descriptive_stats(df, c("Attitudes_to_Drugs" ,
                                                                "Drug_Support"      ,  "Policy_Support" ))




# Cluster Analysis

# HIERARCHICAL AND NON-HIERARCHICAL CLUSTERING

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

df_dummy <- dummy_cols(df, select_columns = vars_demo)

# Function to normalize (scale) variables
normalize_variables_min_max <- function(df, variables) {
  df %>%
    mutate(across(all_of(variables), 
                  ~ (.-min(.))/(max(.) - min(.)), 
                  .names = "{.col}_norm"))
}

vars_to_normalize <- c("")
colnames(df_dummy)

df_normalized <- normalize_variables_min_max(df_dummy, c("Attitudes_to_Drugs" ,
                                                         "Drug_Support"   ,     "Policy_Support"))

names(df_normalized) <- gsub(" ", "_", names(df_normalized))
names(df_normalized) <- gsub("\\(", "_", names(df_normalized))
names(df_normalized) <- gsub("\\)", "_", names(df_normalized))
names(df_normalized) <- gsub("\\-", "_", names(df_normalized))
names(df_normalized) <- gsub("/", "_", names(df_normalized))
names(df_normalized) <- gsub("\\\\", "_", names(df_normalized)) 
names(df_normalized) <- gsub("\\?", "", names(df_normalized))
names(df_normalized) <- gsub("\\'", "", names(df_normalized))
names(df_normalized) <- gsub("\\,", "_", names(df_normalized))
names(df_normalized) <- gsub("\\$", "", names(df_normalized))
names(df_normalized) <- gsub("\\+", "", names(df_normalized))


vars_clustering <- c("Attitudes_to_Drugs" ,
                     "Drug_Support"   ,     "Policy_Support")

library(haven)
library(dplyr)

df_nohaven <- df %>%
  mutate(across(all_of(vars_clustering), ~ as.numeric(as.vector(.))))


clustering_results <- perform_hierarchical_clustering(df_nohaven, vars_clustering)

# View the plots
print(dclustering_results$Elbow_Plot)
print(dclustering_results$Silhouette_Plot)
print(dclustering_results$Gap_Stat_Plot)

# View dataframes and results
Elbow_df <- dclustering_results$Elbow_DataFrame
Silhouette_df <- dclustering_results$Silhouette_DataFrame
GapStat_df <- dclustering_results$Gap_Stat_DataFrame


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

k_range <- 3:8
df_with_clusters_list <- run_hierarchical_kmeans(df_nohaven, vars_clustering, k_range)
df_withclusters <- df_with_clusters_list$Cluster_8


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

anova_results_3clusters <- run_anova(df_withclusters, "Cluster_3", vars_clustering)
anova_results_4clusters <- run_anova(df_withclusters, "Cluster_4", vars_clustering)
anova_results_5clusters <- run_anova(df_withclusters, "Cluster_5", vars_clustering)
anova_results_6clusters <- run_anova(df_withclusters, "Cluster_6", vars_clustering)
anova_results_7clusters <- run_anova(df_withclusters, "Cluster_7", vars_clustering)
anova_results_8clusters <- run_anova(df_withclusters, "Cluster_8", vars_clustering)

anova_compiled_results <- list(anova_results_3clusters, anova_results_4clusters, anova_results_5clusters, 
                               anova_results_6clusters, anova_results_7clusters, anova_results_8clusters)

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

ANOVA_results_summarized <- summarize_anova_tables(anova_compiled_results)

# 4-clusters is the optimal solution - Cluster PROFILING

colnames(df_withclusters)

categories_clustervars <- vars_demo

numbers_clustervars <- c("Attitudes_to_Drugs"  ,                 
                         "Drug_Support"          ,                "Policy_Support")

categories_nonclustervars <- vars_comparison

clusterassignmentcolumn <- "Cluster_6"

# Function to calculate cluster sizes
calculate_cluster_sizes <- function(df, cluster_column) {
  cluster_sizes <- table(df[[cluster_column]])
  return(as.data.frame(cluster_sizes))
}

df_clustersizes <- calculate_cluster_sizes(df_withclusters, clusterassignmentcolumn)

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

factors <- c("Cluster_6")
df_descriptive_clustervars_bycluster <- calculate_means_and_sds_by_factors(df_withclusters, numbers_clustervars, factors)

numbers_clustervars2 <- c("Q21_1"             ,                   
                          "Q21_2"              ,                   "Q21_3"   ,                             
                          "Q21_4"               ,                  "Q21_5"    ,                            
                          "Q21_6"                ,                 "Q22_1"     ,                           
                          "Q22_2"                 ,                "Q23_1"      ,                          
                          "Q23_2"                  ,               "Q23_3"       ,                         
                          "Q23_4"                   ,              "Q23_5"        ,                        
                          "Q23_6"                    ,             "Q23_7"         ,                       
                          "Q23_8"                     ,            "Q23_9"          ,                      
                          "Q23_10"                     ,           "Q23_11")

# Modify function to use labels

calculate_means_and_sds_by_factors <- function(data, var_names, cluster_var) {
  # Retrieve variable labels
  var_labels <- sapply(data[, var_names], attr, "label")
  
  # Ensure cluster_var is a factor
  data[[cluster_var]] <- as.factor(data[[cluster_var]])
  
  # Reshape data to long format
  long_data <- pivot_longer(data, cols = var_names, names_to = "Variable", values_to = "Value")
  
  # Replace variable names with labels in the long dataframe
  long_data$Variable <- factor(long_data$Variable, levels = var_names, labels = var_labels)
  
  # Calculate means and standard deviations by cluster and variable
  summary_df <- long_data %>%
    group_by(Cluster = .data[[cluster_var]], Variable) %>%
    summarise(Mean = mean(Value, na.rm = TRUE),
              SD = sd(Value, na.rm = TRUE),
              .groups = 'drop')
  
  return(summary_df)
}



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




df_segmented_freq <- create_segmented_frequency_tables(df_withclusters, categories_clustervars, "Cluster_6")


calculate_means_and_sds_by_factors <- function(data, var_names, cluster_var) {
  # Retrieve variable labels
  var_labels <- sapply(data[, var_names], attr, "label")
  
  # Ensure cluster_var is a factor
  data[[cluster_var]] <- as.factor(data[[cluster_var]])
  
  # Reshape data to long format
  long_data <- pivot_longer(data, cols = var_names, names_to = "Variable", values_to = "Value")
  
  # Replace variable names with labels in the long dataframe
  long_data$Variable <- factor(long_data$Variable, levels = var_names, labels = var_labels)
  
  # Calculate means and standard deviations by cluster and variable
  summary_df <- long_data %>%
    group_by(Cluster = .data[[cluster_var]], Variable) %>%
    summarise(Mean = mean(Value, na.rm = TRUE),
              SD = sd(Value, na.rm = TRUE),
              .groups = 'drop')
  
  return(summary_df)
}

df_descriptive_detclustervars_bycluster <- calculate_means_and_sds_by_factors(df_withclusters, numbers_clustervars2, factors)

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
  plot <- ggplot(mean_ci_df, aes(x = Variable, y = Mean, group = Cluster, color = as.factor(Cluster))) +
    geom_point(position = position_dodge(width = 0.75)) +
    geom_errorbar(aes(ymin = Lower_CI, ymax = Upper_CI), 
                  width = 0.2, position = position_dodge(width = 0.75)) +
    theme_minimal() +
    labs(title = "Dot and Whiskers Plot of Means by Cluster",
         x = "Variable", y = "Mean", color = "Cluster") + # Specify the legend title for 'color'
    theme(axis.text.x = element_text(angle = 45, hjust = 1, vjust = 1))
  
  return(plot)
}

# Usage
vars_to_plot1 <- numbers_clustervars

vars_to_plot2 <- c("Q21_1"             ,                   
                   "Q21_2"              ,                   "Q21_3"   ,                             
                   "Q21_4")

vars_to_plot3 <- c("Q22_1"     ,                           
                   "Q22_2"    )             

vars_to_plot4 <- c("Q23_1"      ,                          
                   "Q23_2"                  ,               "Q23_3"       ,                         
                   "Q23_4"                   ,              "Q23_5"        ,                        
                   "Q23_6"                    ,             "Q23_7"         ,                       
                   "Q23_8"                     ,            "Q23_9"          ,                      
                   "Q23_10"                     ,           "Q23_11")

dot_whiskers_plot <- create_dot_and_whiskers_plot(df_withclusters, "Cluster_6", vars_to_plot1)
dot_whiskers_plot2 <- create_dot_and_whiskers_plot(df_withclusters, "Cluster_6", vars_to_plot2)
dot_whiskers_plot3 <- create_dot_and_whiskers_plot(df_withclusters, "Cluster_6", vars_to_plot3)
dot_whiskers_plot4 <- create_dot_and_whiskers_plot(df_withclusters, "Cluster_6", vars_to_plot4)

# Display the plot
print(dot_whiskers_plot)
print(dot_whiskers_plot2)
print(dot_whiskers_plot3)
print(dot_whiskers_plot4)


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
    if (anova_p < 0.1) {
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

results_anova_clusters <- perform_anova_and_posthoc(df_withclusters, vars_to_plot1, "Cluster_6")



# Chi-square clusters

# Percentages relative to the column

chi_square_analysis_multiple <- function(data, row_vars, col_var) {
  results <- list() # Initialize an empty list to store results
  
  data[[col_var]] <- factor(as.character(data[[col_var]]))
  
  # Iterate over row variables
  for (row_var in row_vars) {
    # Ensure that the variables are factors
    data[[row_var]] <- factor(data[[row_var]])
    
    
    # Create a crosstab with percentages
    crosstab <- prop.table(table(data[[row_var]], data[[col_var]]), margin = 2) * 100
    
    # Check if the crosstab is 2x2
    is_2x2_table <- all(dim(table(data[[row_var]], data[[col_var]])) == 2)
    
    # Perform chi-square test with correction for 2x2 tables
    chi_square_test <- chisq.test(table(data[[row_var]], data[[col_var]]), correct = is_2x2_table)
    
    # Convert crosstab to a dataframe
    crosstab_df <- as.data.frame.matrix(crosstab)
    
    # Create a dataframe for this pair of variables
    for (level in levels(data[[row_var]])) {
      level_df <- data.frame(
        "Row_Variable" = row_var,
        "Row_Level" = level,
        "Column_Variable" = col_var,
        check.names = FALSE
      )
      
      level_df <- cbind(level_df, crosstab_df[level, , drop = FALSE])
      level_df$Chi_Square <- chi_square_test$statistic
      level_df$P_Value <- chi_square_test$p.value
      
      # Add the result to the list
      results[[paste0(row_var, "_", level)]] <- level_df
    }
  }
  
  # Combine all results into a single dataframe
  do.call(rbind, results)
}

row_variables <- vars_demo
column_variable <- "Cluster_6"  # Replace with your column variable name

df_chisq6 <- chi_square_analysis_multiple(df_withclusters, row_variables, column_variable)
column_variable3 <- "Cluster_6"
df_chisq3 <- chi_square_analysis_multiple(df_withclusters, row_variables, column_variable3)

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
  "Frequency" = df_segmented_freq,
  "Descriptive" = df_descriptive_clustervars_bycluster,
  "Detailed - Descriptives" = df_descriptive_detclustervars_bycluster,
  "Cluster Sizes" = df_clustersizes,
  "ANOVA Results" = anova_results_6clusters
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_3clusters.xlsx")
