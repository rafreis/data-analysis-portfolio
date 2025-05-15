setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/edtertained/Third Order")

library(openxlsx)
df <- read.xlsx("Data.xlsx")

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

# Declare vars for clustering

# Change long name

names(df)[names(df) == "8_.Incident.Command.has.been.implemented.in.major.incidents.that.I.or.my.agency.has.responded.to.._NOTE:.If.your.agency.HAS.NOT.utilized.the.Incident.Command.System_.select.“No”.and.proceed.to.Question.16_"] <- "8_.Incident.Command.has.been.implemented.in.major.incidents.that.I.or.my.agency.has.responded.to.."
names(df)[names(df) == "10_.My.agency.has.conducted.trainings.or.held.meetings.with.neighboring.agencies.to.identify.each.otherâ€™s.capabilities_deficiencies."] <- "10_.My.agency.has.conducted.trainings.or.held.meetings.with.neighboring.agencies.to.identify.each.others.capabilities_deficiencies."
names(df)[names(df) == "6_.What.type.of.geographic.area.is.under.your.agencyâ€™s.jurisdiction"] <- "6_.What.type.of.geographic.area.is.under.your.agencys.jurisdiction"

clustering_vars <- c("Gender"  ,       
          "5_.What.is.the.approximate.size.of.your.agency"                  ,                                                                                                                                             
          "6_.What.type.of.geographic.area.is.under.your.agencys.jurisdiction"    ,                                                                                                                                    
          "7_.What.rank.do.you.hold.within.your.agency"                              ,                                                                                                                                    
          "8_.Incident.Command.has.been.implemented.in.major.incidents.that.I.or.my.agency.has.responded.to..",
          "9_.My.agency.has.utilized.predesignated.locations.within.my.jurisdiction.for.staging.areas.during.a.major.incident."       ,                                                                                   
          "10_.My.agency.has.conducted.trainings.or.held.meetings.with.neighboring.agencies.to.identify.each.others.capabilities_deficiencies."  ,                                                                     
          "11_.Self_deployment.by.members.of.my.agency.and_or.nonmembers.of.my.agency.was.observed.during.a.major.incident."                          ,                                                                   
          "12_.My.agency.has.qualified.members.at.the.ready.to.be.deployed.to.a.Rescue.Task.Force._RTF_.twenty_four.hours.a.day.seven.days.a.week."    ,                                                                  
          "13_.I.have.been.certified.in.ICS.300.by.or.through.my.agency."            ,                                                                                                                                   
          "17_.Interoperable.Capabilities"        ,                                                                                                                                                                       
          "18_.LEOs.within.my.agency.utilize.portable.radios.with.interoperable.capabilities."    ,                                                                                                                       
          "19_.I.can.easily.switch.my.issued.portable.radio.to.an.interoperable.channel."          ,                                                                                                                      
          "20_.My.agency.conducts.training.with.interoperable.radios.where.we.practice.how.to.operate.our.issued.portable.radio."   ,                                                                                     
          "21_.My.agency.conducts.interoperable.communication.training.and.exercises.with.other.agencies."                   ,                                                                                            
          "23_.My.agency.does.not.frequently.provide.its.member.with.intelligence.reports."           ,                                                                                                                   
          "24_.My.agency.has.its.own.standing.Intelligence.Unit."                                  ,                                                                                                                      
          "25_.The.intelligence.unit.within.my.agency.works.solely.from.a.remote.location."        ,                                                                                                                      
          "26_.The.intelligence.unit.within.my.agency.has.been.deployed.to.major.incidents."       ,                                                                                                                      
          "27_.My.agency.relies.on.state_federal.resources.for.intelligence.gathering."            ,                                                                                                                      
          "Age.Groups"                                             ,                                                                                                                                                      
          "Years.as.a.Full.Time.Officer"                       )
            

# HIERARCHICAL AND NON-HIERARCHICAL CLUSTERING

## CLUSTERING ALGORITHM - CHOOSE OPTIMAL NUMBER OF CLUSTERS

library(cluster)
library(klaR)
library(factoextra)
library(dplyr)
library(ggplot2)
library(clusterSim)

perform_kmodes_clustering <- function(df, vars) {
  # Select variables
  data_selected <- df %>% dplyr::select(all_of(vars))

  
  # Initialize lists to store results
  sil_results <- list()
  db_results <- list()
  
  for (k in 2:10) {
    # k-modes clustering
    kmodes_model <- kmodes(data_selected, modes = k)
    
    # Silhouette analysis (note: this might need a custom distance function)
    sil_width <- mean(silhouette(as.integer(kmodes_model$cluster), daisy(data_selected))[, 'sil_width'])
    sil_results[[k]] <- data.frame(Number_of_Clusters = k, Average_Silhouette_Width = sil_width)
    
    
  }
  
  # Combine results into dataframes
  sil_df <- do.call(rbind, sil_results)
  
  
  # Silhouette plot
  sil_plot <- ggplot(sil_df, aes(x = Number_of_Clusters, y = Average_Silhouette_Width)) +
    geom_line() +
    geom_point() +
    theme_minimal() +
    labs(title = "Average Silhouette Widths for Different Numbers of Clusters",
         x = "Number of Clusters", y = "Average Silhouette Width")
  
  
  
  # Return results
  return(list(
    KModes_Clustering = kmodes_model,
    Silhouette_Plot = sil_plot,
    
    Silhouette_DataFrame = sil_df
    
  ))
}

#Dummy categorical columns

library(dplyr)
library(tidyr)
library(fastDummies)

df_dummy <- dummy_cols(df, select_columns = clustering_vars)

clustering_vars <- c(                                                                                                                                                                          
                     "Gender_Male"  ,                                                                                                                                                                               
                     "Gender_NA"     ,                                                                                                                                                                              
                                                                                                                                                              
                     "5_.What.is.the.approximate.size.of.your.agency_> 1500"               ,                                                                                                                        
                     "5_.What.is.the.approximate.size.of.your.agency_50 - 500"              ,                                                                                                                       
                     "5_.What.is.the.approximate.size.of.your.agency_501 - 1500"             ,                                                                                                                      
                     "5_.What.is.the.approximate.size.of.your.agency_NA"                      ,                                                                                                                     
                                                                                                                                      
                     "6_.What.type.of.geographic.area.is.under.your.agencys.jurisdiction_Suburban" ,                                                                                                             
                     "6_.What.type.of.geographic.area.is.under.your.agencys.jurisdiction_Urban"     ,                                                                                                            
                     "6_.What.type.of.geographic.area.is.under.your.agencys.jurisdiction_NA"         ,                                                                                                           
                                                                                                                                                      
                     "7_.What.rank.do.you.hold.within.your.agency_Non-Supervisory"                       ,                                                                                                          
                     "7_.What.rank.do.you.hold.within.your.agency_Supervisory"                            ,                                                                                                         
                     "7_.What.rank.do.you.hold.within.your.agency_NA"                                      ,                                                                                                        
                                                                                                            
                     "8_.Incident.Command.has.been.implemented.in.major.incidents.that.I.or.my.agency.has.responded.to.._Yes"                           ,                                                           
                                                                                                         
                                                                                           
                     "9_.My.agency.has.utilized.predesignated.locations.within.my.jurisdiction.for.staging.areas.during.a.major.incident._Yes"           ,                                                          
                                                                                        
                                                                       
                     "10_.My.agency.has.conducted.trainings.or.held.meetings.with.neighboring.agencies.to.identify.each.others.capabilities_deficiencies._Yes",                                                  
                                                                      
                                                                                              
                     "11_.Self_deployment.by.members.of.my.agency.and_or.nonmembers.of.my.agency.was.observed.during.a.major.incident._Yes"                       ,                                                 
                                                                                           
                                                                       
                     "12_.My.agency.has.qualified.members.at.the.ready.to.be.deployed.to.a.Rescue.Task.Force._RTF_.twenty_four.hours.a.day.seven.days.a.week._Yes" ,                                                
                                                                      
                                                                                                                                               
                     "13_.I.have.been.certified.in.ICS.300.by.or.through.my.agency._Yes"                                                                            ,                                               
                                                                                                                                                
                                                                                                                                                                              
                     "17_.Interoperable.Capabilities_Yes"                                                                                                            ,                                              
                                                                                                                                                                                
                                                                                                                            
                     "18_.LEOs.within.my.agency.utilize.portable.radios.with.interoperable.capabilities._Yes"                                                         ,                                             
                                                                                                                                
                     "19_.I.can.easily.switch.my.issued.portable.radio.to.an.interoperable.channel._Yes"                                                               ,                                            
                                                                                         
                     "20_.My.agency.conducts.training.with.interoperable.radios.where.we.practice.how.to.operate.our.issued.portable.radio._Yes"  ,                                                                 
                                                                                                                
                     "21_.My.agency.conducts.interoperable.communication.training.and.exercises.with.other.agencies._Yes"                          ,                                                                
                                                                                                                               
                     "23_.My.agency.does.not.frequently.provide.its.member.with.intelligence.reports._Yes"                                          ,                                                               
                                                                                                                                                        
                     "24_.My.agency.has.its.own.standing.Intelligence.Unit._Yes"                                                                     ,                                                              
                                                                                                                               
                     "25_.The.intelligence.unit.within.my.agency.works.solely.from.a.remote.location._Yes"                                            ,                                                             
                                                                                                                              
                     "26_.The.intelligence.unit.within.my.agency.has.been.deployed.to.major.incidents._Yes"                                            ,                                                            
                                                                                                                                  
                     "27_.My.agency.relies.on.state_federal.resources.for.intelligence.gathering._Yes"                                                  ,                                                           
                                                                                                                                   
                                                                                                                                                                                                 
                     "Age.Groups_> 50"                                                                                                                   ,                                                          
                     "Age.Groups_35-50"                                                                                                                   ,                                                         
                     "Age.Groups_NA"                                                                                                                       ,                                                        
                                                                                                                                                                                
                     "Years.as.a.Full.Time.Officer_> 20"                                                                                                    ,                                                       
                     "Years.as.a.Full.Time.Officer_10-20"                                                                                                    ,                                                      
                     "Years.as.a.Full.Time.Officer_NA" )

# Replace NA by 0
df_dummy <- df_dummy %>% 
  mutate(across(all_of(clustering_vars), ~replace(., is.na(.), 0)))

# Convert to factors
for(col in clustering_vars) {
  df_dummy[[col]] <- factor(df_dummy[[col]])
}

clustering_results <- perform_kmodes_clustering(df_dummy, clustering_vars)

# View the plots
print(clustering_results$Silhouette_Plot)

# View dataframes and results
Silhouette_df <- clustering_results$Silhouette_DataFrame

library(klaR)
library(dplyr)
library(cluster)

perform_kmodes_for_selected_clusters <- function(df, vars, num_clusters) {
  # Select variables
  data_selected <- df %>% dplyr::select(all_of(vars))
  
  # Run k-modes clustering
  kmodes_model <- kmodes(data_selected, modes = num_clusters)
  
  # Return the model and a summary
  return(list(
    KModes_Model = kmodes_model,
    Summary = summary(kmodes_model)
  ))
}

optimal_clusters <- 2

# Running the function again to store results
clustering_results_2 <- perform_kmodes_for_selected_clusters(df_dummy, clustering_vars, optimal_clusters)

# Access the results
kmodes_model_2 <- clustering_results_2$KModes_Model
model_summary_2 <- clustering_results_2$Summary

# Running the function again to store results
clustering_results_3 <- perform_kmodes_for_selected_clusters(df_dummy, clustering_vars, optimal_clusters)

# Access the results
kmodes_model_3 <- clustering_results_3$KModes_Model
model_summary_3 <- clustering_results_3$Summary

# COMPARE AND PROFILE CLUSTERS
library(dplyr)
library(tidyr)
library(ggplot2)
library(chisq.posthoc.test)

perform_kmodes_comparison <- function(df, kmodes_model, vars) {
  # Add cluster assignments to the dataframe
  df_with_clusters <- df %>% mutate(Cluster = kmodes_model$cluster)
  
  # Initialize a list to store chi-square test results
  chi_sq_results <- list()
  
  # Perform chi-square tests for each variable
  for (var in vars) {
    # Construct a contingency table for each categorical variable against the clusters
    contingency_table <- table(df_with_clusters[[var]], df_with_clusters$Cluster)
    
    # Perform chi-square test
    chi_sq_test <- chisq.test(contingency_table)
    
    # Store results
    chi_sq_results[[var]] <- list(
      Variable = var,
      Chi_Square = chi_sq_test$statistic,
      p_value = chi_sq_test$p.value
      # Post-hoc analysis can be added if needed
    )
  }
  
  # Combine chi-square results into a single dataframe
  combined_chi_sq_results <- bind_rows(chi_sq_results, .id = "Variable")
  
  return(list(
    Chi_Square_Results = combined_chi_sq_results,
    DataFrame_with_Clusters = df_with_clusters
  ))
}

comparison_vars <- c("Gender"  ,       
                     "5_.What.is.the.approximate.size.of.your.agency"                  ,                                                                                                                                             
                     "6_.What.type.of.geographic.area.is.under.your.agencys.jurisdiction"    ,                                                                                                                                    
                     "7_.What.rank.do.you.hold.within.your.agency"                              ,                                                                                                                                    
                     "8_.Incident.Command.has.been.implemented.in.major.incidents.that.I.or.my.agency.has.responded.to..",
                     "9_.My.agency.has.utilized.predesignated.locations.within.my.jurisdiction.for.staging.areas.during.a.major.incident."       ,                                                                                   
                     "10_.My.agency.has.conducted.trainings.or.held.meetings.with.neighboring.agencies.to.identify.each.others.capabilities_deficiencies."  ,                                                                     
                     "11_.Self_deployment.by.members.of.my.agency.and_or.nonmembers.of.my.agency.was.observed.during.a.major.incident."                          ,                                                                   
                     "12_.My.agency.has.qualified.members.at.the.ready.to.be.deployed.to.a.Rescue.Task.Force._RTF_.twenty_four.hours.a.day.seven.days.a.week."    ,                                                                  
                     "13_.I.have.been.certified.in.ICS.300.by.or.through.my.agency."            ,                                                                                                                                   
                     "17_.Interoperable.Capabilities"        ,                                                                                                                                                                       
                     "18_.LEOs.within.my.agency.utilize.portable.radios.with.interoperable.capabilities."    ,                                                                                                                       
                     "19_.I.can.easily.switch.my.issued.portable.radio.to.an.interoperable.channel."          ,                                                                                                                      
                     "20_.My.agency.conducts.training.with.interoperable.radios.where.we.practice.how.to.operate.our.issued.portable.radio."   ,                                                                                     
                     "21_.My.agency.conducts.interoperable.communication.training.and.exercises.with.other.agencies."                   ,                                                                                            
                     "23_.My.agency.does.not.frequently.provide.its.member.with.intelligence.reports."           ,                                                                                                                   
                     "24_.My.agency.has.its.own.standing.Intelligence.Unit."                                  ,                                                                                                                      
                     "25_.The.intelligence.unit.within.my.agency.works.solely.from.a.remote.location."        ,                                                                                                                      
                     "26_.The.intelligence.unit.within.my.agency.has.been.deployed.to.major.incidents."       ,                                                                                                                      
                     "27_.My.agency.relies.on.state_federal.resources.for.intelligence.gathering."            ,                                                                                                                      
                     "Age.Groups"                                             ,                                                                                                                                                      
                     "Years.as.a.Full.Time.Officer"                       )
  
kmodes_chisq_comparison <- perform_kmodes_comparison(df_dummy, kmodes_model_2, comparison_vars)

df_wClusters <- kmodes_chisq_comparison$DataFrame_with_Clusters
df_chisq_results2 <- kmodes_chisq_comparison$Chi_Square_Results

kmodes_chisq_comparison3 <- perform_kmodes_comparison(df_dummy, kmodes_model_3, comparison_vars)
df_chisq_results3 <- kmodes_chisq_comparison3$Chi_Square_Results

# Calculate the sum of the Chi-Square values
chi_square_sum <- sum(df_chisq_results2$Chi_Square)

# Count the number of p-values less than 0.05
p_value_count <- sum(df_chisq_results2$p_value < 0.05)

# Output the results
chi_square_sum
p_value_count

# Calculate the sum of the Chi-Square values
chi_square_sum3 <- sum(df_chisq_results3$Chi_Square)

# Count the number of p-values less than 0.05
p_value_count3 <- sum(df_chisq_results3$p_value < 0.05)

# Output the results
chi_square_sum3
p_value_count3

library(stringr)

perform_kmodes_profiling_with_plots <- function(df, cluster_var, vars) {
  df_with_clusters <- df %>% mutate(!!sym(cluster_var) := as.factor(.[[cluster_var]]))
  
  create_segmented_frequency_tables <- function(data, vars, cluster_var) {
    freq_tables <- list()
    for (var in vars) {
      # Create frequency table
      freq_table <- as.data.frame(table(data[[var]], data[[cluster_var]]))
      names(freq_table) <- c("Category", "Cluster", "Count")
      
      # Calculate the total counts for each cluster
      total_counts <- aggregate(Count ~ Cluster, data = freq_table, FUN = sum)
      names(total_counts) <- c("Cluster", "TotalCount")
      
      # Merge total counts with the frequency table
      freq_table <- merge(freq_table, total_counts, by = "Cluster")
      
      # Calculate the percentage
      freq_table$Percentage <- with(freq_table, (Count / TotalCount) * 100)
      
      # Add the variable name
      freq_table$Variable <- var
      freq_tables[[var]] <- freq_table
    }
    return(do.call(rbind, freq_tables))
  }
  
  freq_tables <- create_segmented_frequency_tables(df_with_clusters, vars, cluster_var)
  
  # Calculate the number of pages needed based on the number of variables
  num_plots_per_page <- 16
  num_pages <- ceiling(length(unique(freq_tables$Variable)) / num_plots_per_page)
  
  plot_list <- list()
  percentage_plot_list <- list()  # List to store percentage plots
  
  for (i in 1:num_pages) {
    vars_subset <- unique(freq_tables$Variable)[((i - 1) * num_plots_per_page + 1):(i * num_plots_per_page)]
    names_with_spaces <- gsub("\\.", " ", vars_subset)
    wrapped_names <- sapply(names_with_spaces, stringr::str_wrap, width = 35)
    
    # Subset for count plots
    freq_tables_subset <- freq_tables %>% filter(Variable %in% vars_subset)
    
    # Plot for counts
    p <- ggplot(freq_tables_subset, aes(x = Cluster, y = Count, fill = Category)) +
      geom_bar(stat = "identity", position = position_dodge()) +
      facet_wrap(~ Variable, scales = "free", ncol = 4, labeller = as_labeller(setNames(wrapped_names, vars_subset))) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 6),
            strip.text.x = element_text(size = 7, hjust = 0.5),
            legend.position = "bottom") +
      labs(y = "Count", x = "Cluster")
    
    # Save count plot
    ggsave(paste0("clusterprofiles_counts_page_", i, ".png"), p, width = 20, height = 20, units = "cm")
    plot_list[[i]] <- p
    
    # Plot for percentages
    p_percentage <- ggplot(freq_tables_subset, aes(x = Cluster, y = Percentage, fill = Category)) +
      geom_bar(stat = "identity", position = position_dodge()) +
      facet_wrap(~ Variable, scales = "free", ncol = 4, labeller = as_labeller(setNames(wrapped_names, vars_subset))) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 6),
            strip.text.x = element_text(size = 7, hjust = 0.5),
            legend.position = "bottom") +
      labs(y = "Percentage", x = "Cluster")  # Update label for y-axis
    
    # Save percentage plot
    ggsave(paste0("clusterprofiles_percentages_page_", i, ".png"), p_percentage, width = 20, height = 20, units = "cm")
    percentage_plot_list[[i]] <- p_percentage
  }
  
  # Include the percentage_plot_list in your return statement
  return(list(
    DataFrame_with_Clusters = df_with_clusters,
    Frequency_Tables = freq_tables,
    Count_Plot_List = plot_list,
    Percentage_Plot_List = percentage_plot_list  # Add this line to return the percentage plots
  ))
}


profiling_2clusters <- perform_kmodes_profiling_with_plots(df_wClusters, "Cluster", comparison_vars)
df_freq_table2 <- profiling_2clusters$Frequency_Tables

plot_list <- profiling_2clusters$Plot_List

# Loop through the list and print each plot
for (plot in plot_list) {
  print(plot)
}


# Merge tables

df_merged_results <- left_join(df_freq_table2, df_chisq_results2, by = "Variable")


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

data_list <- list(
  "Silhouette Analysis" = Silhouette_df,
  "Data with Clusters" = df_wClusters, 
  "Chi-Square Results" = df_chisq_results2, 
  "Cluster Profiling" = df_merged_results
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")




