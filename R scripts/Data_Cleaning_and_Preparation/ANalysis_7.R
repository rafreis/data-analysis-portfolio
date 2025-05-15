setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/jazzolynn")

library(openxlsx)
df <- read.xlsx("Post-Survey_CPM_TIC_Fall_2023_February_18__2024_15.18_Numeric.xlsx", sheet = "CleanData")


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

colnames(df)

# Definins variables

knowledge_vars <- c("Knowledge..Definition.of.a.crisis"                          ,                                                                                                                                                                    
                    "Knowledge..Concept.of.crisis.prevention.management._CPM_"    ,                                                                                                                                                                   
                    "Knowledge..Definition.of.trauma"                              ,                                                                                                                                                                  
                    "Knowledge..Definition.of.triggers"                             ,                                                                                                                                                                 
                    "Knowledge..Concept.of.trauma.informed.care._TIC_"               ,                                                                                                                                                                
                    "Knowledge..Concept.of.creating.a.culture.of.safety"              ,                                                                                                                                                               
                     "Knowledge..Concept.of.creating.a.culture.of.civility"            ,                                                                                                                                                               
                    "Knowledge..Concept.of.various.forms.of.aggression.such.as.physical_.psychological_.and.verbal" ,                                                                                                                                 
                    "Knowledge..Concept.of.adverse.childhood.experiences._ACEs_"                                     ,                                                                                                                                
                     "Knowledge..Definition.of.public.health"                                                         ,                                                                                                                                
                    "Knowledge..Concept.of.recognizing.violence.as.a.public.health.crisis"                             ,                                                                                                                              
                    "Knowledge..Concept.of.public.health.model.as.a.framework.for.crisis.prevention.management.interventions",                                                                                                                        
                    "Knowledge..Concept.of.situational.awareness.in.the.context.of.crisis.prevention.management"              ,                                                                                                                       
                    "Knowledge..Concept.of.the.four.Fs._fight_.flight_.freeze_.and.fawn_.and.their.relevance.to.crisis.prevention.management"   )

confidence_vars <- c("Confidence.Level.Handling.aggressive.events"     ,                                                                                                                                                                               
                      "Confidence.Level.Understanding.the.communication.steps.involved.in.crisis.de_escalation.such.as.listening_.restating.concerns.and.offering.choices"      ,                                                                       
                     "Confidence.Level.Familiarity.with.the.phases.of.the.behavior.in.the.crisis.model._1..calm_.2..triggering.event_.4..emotional.responses_.etc_.and.their.application.in.crisis.prevention"   ,                                     
                     "Confidence.Level.Awareness.of.verbal_.nonverbal_.and.behavioral.cues.that.may.indicate.potential.violence.or.disruptive.behaviors"                                               ,                                               
                     "Confidence.Level.Ability.to.handle.situations.where.verbal.escalation.continues.including.strategies.like.providing.choices.and.redirection"                                    ,                                                
                      "Confidence.Level.Familiarity.with.allowing.one.person.to.communicate.at.a.time.when.managing.disruptive.behavior"                                                               ,                                                
                     "Confidence.Level.How.comfortable.are.you.in.asking.individuals.to.stop.disruptive.behavior.and.explaining.why.it.makes.others.feel.unsafe"                                        ,                                              
                      "Confidence.Level.When.managing.the.disruptive.behavior_.how.confident.do.you.follow.the.practice.of.slowing.down.cognitive.processing.and.providing.clear.directions.with.two.options.if.possible"  ,                            
                     "Confidence.Level.When.managing.disruptive.behavior_.how.confident.are.you.in.your.effectiveness.in.reinforcing.a.collaborative.approach"         ,                                                                               
                     "Confidence.Level.Communication.about.what.you.can.do.about.the.situation.while.asking.about.individual.needs"                                     ,                                                                              
                     "Confidence.Level.Assessing.the.need.for.involvement.of.fellow.colleagues.when.managing.disruptive.behaviors"                                       ,                                                                             
                     "Confidence.Level.When.managing.disruptive.behaviors.how.confident.do.you.feel.about.effectively.reinforcing.a.collaborative.approach.and.expressing.a.desire.to.work.together.with.the.individual" )

demographics <- c("What.is.your.age"         ,                                                                                                                                                                                                      
                  "What.is.your.sex"          ,                                                                                                                                                                                                     
                  "Choose.one.or.more.races.that.you.consider.yourself.to.be" ,                                                                                                                                                                     
                  "Are.you.of.Spanish_.Hispanic_.or.Latino.origin"             ,                                                                                                                                                                    
                  "What.is.the.highest.level.of.education.you.have.completed",
                  "Do.you.aim.to.understand.the.individuals.needs.and.validate.their.feelings.in.disruptive.situations"            ,                                                                                                                
                  "Do.you.focus.on.addressing.unsafe.behavior.rather.than.the.individual.when.intervening.in.disruptive.situations" ,
                  "Have.you.ever.received.training.on.identifying.potential.external.triggering.causes")

satisfaction_vars <- c( "How.satisfied.are.you.with.the.clarity.and.comprehensiveness.of.the.educational.material.provided.on.crisis.prevention.management_.trauma_informed.care_.and.related.concepts"                              ,                    
                        "How.satisfied.are.you.with.the.overall.educational.programs.effectiveness.in.enhancing.your.knowledge.and.confidence.in.handling.aggressive.incidents.and.challenging.situations.at.the.shelter"             ,                   
                       "On.a.scale.of.1_5_.to.what.extent.do.you.feel.that.the.concepts.covered.in.the.program_.such.as.creating.a.culture.of.safety_.trauma_informed.care_.and.crisis.prevention.management_.were.relevant.to.your.role.at.the.shelter")


other_vars <- c("On.a.scale.of.1_5_.when.managing.disruptive.behavior_.how.well.do.you.follow.the.practice.of.slowing.down.cognitive.processing.and.providing.clear.directions.with.two.options_.if.possible"  ,
                                                                                                                                                            
                "On.a.scale.of.1_5_.how.often.do.you.introduce.yourself.and.use.the.persons.name.when.addressing.individuals.in.disruptive.behavior.management"                          ,                                                        
                                                                                                                              
                "On.a.scale.of.1_5_.how.well.do.you.keep.the.focus.on.the.present.moment.when.addressing.disruptive.behavior")

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


df_freq <- create_frequency_tables(df, demographics)

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


df_segmented_freq <- create_segmented_frequency_tables(df, demographics, "Pre_Post")

# Reliability Analysis
# Scale Construction

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SE_of_the_Mean = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val <- alpha_results$total$raw_alpha
    
    # Initialize item_itc to NA to avoid scoping issues
    item_itc <- NA
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      if (!is.null(alpha_results$item.stats[item, "raw.r"])) {
        item_itc <- alpha_results$item.stats[item, "raw.r"] # Get ITC for the item
      }
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = mean(item_data, na.rm = TRUE),
        SE_of_the_Mean = sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data))),
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
        Alpha = NA  # Alpha is NA for individual items
      ))
    }
    
    # Calculate the mean score for the scale and add it as a new column to `data`
    scale_mean <- rowMeans(subset_data, na.rm = TRUE, dims = 1)
    data[[scale]] <- scale_mean
    
    # Append scale statistics to results
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = mean(scale_mean, na.rm = TRUE),
      SE_of_the_Mean = sd(scale_mean, na.rm = TRUE) / sqrt(sum(!is.na(scale_mean))),
      StDev = sd(scale_mean, na.rm = TRUE),
      ITC = NA,  # ITC does not apply to the scale as a whole
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}

scales <- list(
  "Knowledge" = knowledge_vars,
  "Confidence.Level" = confidence_vars,
  "Satisfaction" = satisfaction_vars)

df_reliability <- reliability_analysis(df,scales)

df_reliability_stat <- df_reliability$statistics

df <- df_reliability$data_with_scales

# Descriptives

library(dplyr)

calculate_descriptive_stats_bygroups <- function(data, desc_vars, cat_column) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Category = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics for each category
  for (var in desc_vars) {
    # Group data by the categorical column
    grouped_data <- data %>%
      group_by(!!sym(cat_column)) %>%
      summarise(
        Mean = mean(!!sym(var), na.rm = TRUE),
        SEM = sd(!!sym(var), na.rm = TRUE) / sqrt(n()),
        SD = sd(!!sym(var), na.rm = TRUE)
        
      ) %>%
      mutate(Variable = var)
    
    # Append the results for each variable and category to the results dataframe
    results <- rbind(results, grouped_data)
  }
  
  return(results)
}

knowledge_vars <- c(knowledge_vars, "Knowledge")
confidence_vars <- c(confidence_vars, "Confidence.Level")
satisfaction_vars <- c(satisfaction_vars, "Satisfaction")

all_vars <- c(knowledge_vars, confidence_vars, satisfaction_vars, other_vars)

df_descriptives <- calculate_descriptive_stats_bygroups(df, all_vars, "Pre_Post")

# Mann-Whitney Test

library(dplyr)

# Function to calculate effect size for Mann-Whitney U tests
# Note: There's no direct equivalent to Cohen's d for Mann-Whitney, but we can use rank biserial correlation as an effect size measure.

run_mann_whitney_tests <- function(data, group_col, group1, group2, measurement_cols) {
  results <- data.frame(Variable = character(),
                        U_Value = numeric(),
                        P_Value = numeric(),
                        Effect_Size = numeric(),
                        Group1_Size = integer(),
                        Group2_Size = integer(),
                        stringsAsFactors = FALSE)
  
  # Iterate over each measurement column
  for (var in measurement_cols) {
    # Extract group data with proper filtering and subsetting
    group1_data <- data %>%
      filter(.[[group_col]] == group1) %>%
      select(var) %>%
      pull() %>%
      na.omit()
    
    group2_data <- data %>%
      filter(.[[group_col]] == group2) %>%
      select(var) %>%
      pull() %>%
      na.omit()
    
    group1_size <- length(group1_data)
    group2_size <- length(group2_data)
    
    if (group1_size > 0 && group2_size > 0) {
      # Perform Mann-Whitney U test
      mw_test_result <- wilcox.test(group1_data, group2_data, paired = FALSE, exact = FALSE)
      
      # Calculate effect size (rank biserial correlation) for Mann-Whitney U test
      effect_size <- (mw_test_result$statistic - (group1_size * group2_size / 2)) / sqrt(group1_size * group2_size * (group1_size + group2_size + 1) / 12)
      
      results <- rbind(results, data.frame(
        Variable = var,
        U_Value = mw_test_result$statistic,
        P_Value = mw_test_result$p.value,
        Effect_Size = effect_size,
        Group1_Size = group1_size,
        Group2_Size = group2_size
      ))
    }
  }
  
  return(results)
}



df_mannwhitney <- run_mann_whitney_tests(df, "Pre_Post", "Pre", "Post", all_vars)

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
  "Cleaned Data" = df,
  "Frequency Table" = df_freq, 
  "Frequency Table - Segmented" = df_segmented_freq,
  "Descriptives" = df_descriptives,
  "Reliability Test" = df_reliability_stat,
  "MannWhitney" = df_mannwhitney
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")



# Plot

library(dplyr)
library(ggplot2)

# Assuming 'df' is your dataset and 'Pre_Post' is the column with levels 'Pre' and 'Post'

# First, calculate the mean for each variable by 'Pre_Post' status
mean_by_pre_post <- df %>%
  pivot_longer(cols = c(knowledge_vars, confidence_vars), names_to = "Variable", values_to = "Score") %>%
  group_by(Variable, Pre_Post) %>%
  summarise(Mean_Score = mean(Score, na.rm = TRUE), .groups = 'drop')

# Calculate percentage change from Pre to Post for each variable
percentage_changes <- mean_by_pre_post %>%
  group_by(Variable) %>%
  summarise(Percentage_Change = (Mean_Score[Pre_Post == "Post"] - Mean_Score[Pre_Post == "Pre"]) / Mean_Score[Pre_Post == "Pre"] * 100)

# Separate the data into knowledge and satisfaction for plotting
knowledge_changes <- percentage_changes %>%
  filter(Variable %in% knowledge_vars)

library(stringr)  # for str_wrap

# Custom function to wrap labels and replace characters
wrap_labels <- function(variable_names) {
  # Replace dots and underscores with spaces
  variable_names <- gsub("\\.", " ", variable_names)
  variable_names <- gsub("_", " ", variable_names)
  # Wrap labels to a specified width
  variable_names <- str_wrap(variable_names, width = 70)
  return(variable_names)
}

# Update the plotting function
plot_percentage_change <- function(data, title) {
  ggplot(data, aes(x = wrap_labels(Variable), y = Percentage_Change)) +
    geom_bar(stat = "identity", aes(fill = Variable), show.legend = FALSE) +
    geom_text(aes(label = sprintf("%.1f%%", Percentage_Change)), position = position_dodge(width = 0.9), hjust = -0.1, size = 3.5, color = "black") + # Add data labels outside the bars
    theme_minimal() +
    coord_flip() + # Flip coordinates for horizontal bars
    labs(title = title, x = "Variable", y = "Mean Percentage Change (%)") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          axis.title.x = element_blank(),
          plot.title = element_text(size = 14, face = "bold"))
}

# Assuming the `knowledge_changes` data frame is already created with the percentage changes calculated
# Now, we need to calculate the percentage changes for the 'confidence_vars'

confidence_changes <- percentage_changes %>%
  filter(Variable %in% confidence_vars)

# Plot for Knowledge Variables
plot_knowledge <- plot_percentage_change(knowledge_changes, "Knowledge Variables Percentage Change")

# Plot for Confidence Level Variables
plot_confidence <- plot_percentage_change(confidence_changes, "Confidence Level Variables Percentage Change")

# Display the plots
print(plot_knowledge)
print(plot_confidence)




# Update the plotting function for grayscale
plot_percentage_change <- function(data, title) {
  ggplot(data, aes(x = wrap_labels(Variable), y = Percentage_Change)) +
    geom_bar(stat = "identity", fill = "grey80", show.legend = FALSE) +  # Set bar fill to a shade of gray
    geom_text(aes(label = sprintf("%.1f%%", Percentage_Change)), position = position_dodge(width = 0.9), hjust = -0.1, size = 3.5, color = "black") + # Add data labels outside the bars
    theme_minimal() +
    coord_flip() + # Flip coordinates for horizontal bars
    labs(title = title, x = "", y = "Mean Percentage Change (%)") + # Remove x-axis title
    theme(axis.text.x = element_text(angle = 45, hjust = 1, color = "black"),
          axis.title.x = element_blank(),
          plot.title = element_text(size = 14, face = "bold"),
          axis.text.y = element_text(color = "black")) +
    scale_fill_grey(start = 0, end = 1)  # Use grayscale for the fill
}


# Plot for Knowledge Variables
plot_knowledge <- plot_percentage_change(knowledge_changes, "Knowledge Variables Percentage Change")

# Plot for Confidence Level Variables
plot_confidence <- plot_percentage_change(confidence_changes, "Confidence Level Variables Percentage Change")

# Display the plots
print(plot_knowledge)
print(plot_confidence)