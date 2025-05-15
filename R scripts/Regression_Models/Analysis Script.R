setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/tiffanysb91")

library(openxlsx)
df <- read.csv("data_clean.csv")

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

scales_SE <- c("On.a.whole..I.am.satisfied.with.myself."         ,                                                                                                                           
               "At.times.I.think.I.am.no.good.at.all."            ,                                                                                                                          
               "I.feel.that.I.have.a.number.of.good.qualities."    ,                                                                                                                         
               "I.am.able.to.do.things.as.well.as.most.other.people." ,                                                                                                                      
               "I.feel.I.do.not.have.much.to.be.proud.of.",                                                                                                                                  
               "I.certainly.feel.useless.at.times."        ,                                                                                                                                 
               "I.feel.that.I.am.a.person.of.worth."        ,                                                                                                                                
               "I.wish.I.could.have.more.respect.for.myself." ,                                                                                                                              
               "All.in.all..I.am.inclined.to.think.that.I.am.a.failure.",                                                                                                                    
               "I.take.a.positive.attitude.toward.myself.")

scales_SE_rev <- c(                                                                                                                          
               "At.times.I.think.I.am.no.good.at.all."            ,                                                                                                                          
                                                                                                                                        
                                                                                                                                    
               "I.feel.I.do.not.have.much.to.be.proud.of.",                                                                                                                                  
               "I.certainly.feel.useless.at.times."        ,                                                                                                                                 
                                                                                                                                              
               "I.wish.I.could.have.more.respect.for.myself." ,                                                                                                                              
               "All.in.all..I.am.inclined.to.think.that.I.am.a.failure."                                                                                                                  
               )

# Count the number of rows with 'Neither Agree nor Disagree' in at least one of the specified columns
rows_with_response <- apply(df[scales_SE], 1, function(row) {
  any(row == "Neither agree nor disagree")
})

count <- sum(rows_with_response)

count

scales_reversecode <- c("I.feel.that.I.am.a.person.of.worth." ,"I.feel.that.I.have.a.number.of.good.qualities.")

scales_exposuretodisccontent <- c("Frequency.of.Discriminatory.Content...Facebook"    ,                                                                                                                         
                                "Frequency.of.Discriminatory.Content...Instagram"    ,                                                                                                                        
                                "Frequency.of.Discriminatory.Content...Twitter.or.X"   ,                                                                                                                      
                                "Frequency.of.Discriminatory.Content...TikTok"        ,                                                                                                                       
                                "Frequency.of.Discriminatory.Content....YouTube"     ,                                                                                                                        
                                "Frequency.of.Discriminatory.Content....Snapchat"   ,                                                                                                                         
                                "Frequency.of.Discriminatory.Content....YouTube.1",
                                "Frequency.of.Discriminatory.Content...Other..please.specify.")

scales_typetodisccontent <- c("Type.of.Discriminatory.Content...Negative.media.depictions.and.representations.of.African.Americans"       ,                                                                 
                              "Type.of.Discriminatory.Content...Stereotypical.depictions.and.representations.of.African.Americans"         ,                                                                
                              "Type.of.Discriminatory.Content...Rhetoric.from.public.figures.and.politicians.targeting.African.Americans"   ,                                                               
                              "Type.of.Discriminatory.Content...Presence.of.online.hate.groups.or.figures.targeting.African.Americans"       ,                                                              
                              "Type.of.Discriminatory.Content...Images.and.videos.depicting.brutal.attacks.targeting.African.Americans"       ,                                                             
                              "Type.of.Discriminatory.Content...Other..please.specify.")

scales_otheragreem <- c("Viewing.discriminatory.content.on.social.media....negatively.affects.my.mood"                      ,                                                                         
                        "Viewing.discriminatory.content.on.social.media....creates.a.lot.of.anxiety.and.distress.in.me"      ,                                                                        
                        "Viewing.discriminatory.content.on.social.media....can.affect.how.I.feel.for.the.rest.of.the.day"     ,                                                                       
                        "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.feel.about.myself"         ,                                                                      
                        "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.think.others.in.the..real..world.think.about.me"    ,                                             
                        "Viewing.discriminatory.content.on.social.media....negatively.affects.how.I.think.about.my.self.worth")


scales_freq_cat <- c("Gender"                         ,                                                                                                                                            
                     "Education.Level"                 ,                                                                                                                                           
                     "Occupation...Selected.Choice",
                     "Facebook.Usage"               ,                                                                                                                                              
                     "Instagram.Usage"               ,                                                                                                                                             
                     "X.Twitter.or.X.Usage"           ,                                                                                                                                            
                     "TikTok.Usage"                    ,                                                                                                                                           
                     "Snapchat.Usage"                   ,                                                                                                                                          
                     "YouTube.Usage"                     ,                                                                                                                                         
                     "Tumblr.Usage"                       ,                                                                                                                                        
                     "Reddit.Usage",
                     "Other..please.specify.",
                     "Have.you.ever.reported.discriminatory.content.on.social.media.platforms.to.platform.administrators."                            ,                                            
                     "Have.you.ever.unfollowed.or.blocked.individuals.or.accounts.on.social.media.due.to.the.presence.of.discriminatory.content.in.their.posts.",                                  
                     "Do.you.believe.that.social.media.platforms.effectively.address.and.moderate.discriminatory.content."         ,                                                               
                     "Have.you.ever.participated.in.online.discussions.or.activism.related.to.combating.discriminatory.content."    ,                                                              
                     "To.what.degree.do.you.believe.these.discussions.or.acts.of.activism.have.been.helpful.or.empowering."          ,                                                             
                     "How.aware.are.you.of.the.ability.to.modify.the.algorithms.that.determine.the.content.shown.to.you.on.social.media.platforms."        ,                                       
                     "How.comfortable.do.you.feel.discussing.the.impact.of.discriminatory.content.on.social.media.with.your.peers.or.support.networks."     ,                                      
                     "How.comfortable.would.you.be.discussing.the.impact.of.social.media.and.discriminatory.content.with.a.therapist.or.mental.health.professional.",                              
                     "To.what.extent.are.you.inclined.to.seek.support.or.engage.with.therapists.specializing.in.mental.health.and.self.esteem.for.individuals.impacted.by.discriminatory.content.",
                     "How.effective.do.you.find.your.current.coping.mechanisms.in.managing.the.emotional.toll.of.discriminatory.content.on.social.media.")


scales_freq_cont <- "Age"

scales_mediausage <- c("Facebook.Usage"               ,                                                                                                                                              
                     "Instagram.Usage"               ,                                                                                                                                             
                     "X.Twitter.or.X.Usage"           ,                                                                                                                                            
                     "TikTok.Usage"                    ,                                                                                                                                           
                     "Snapchat.Usage"                   ,                                                                                                                                          
                     "YouTube.Usage"                     ,                                                                                                                                         
                     "Tumblr.Usage"                       ,                                                                                                                                        
                     "Reddit.Usage")

# Recoding

library(dplyr)

recode_all_columns <- function(df, recoding_scheme, columns = NULL) {
  if (is.null(columns)) {
    columns <- names(df)
  }
  
  df %>%
    mutate(across(all_of(columns), function(x) {
      sapply(x, function(value) {
        if (as.character(value) %in% names(recoding_scheme)) {
          return(recoding_scheme[[as.character(value)]])
        } else {
          return(value)
        }
      })
    }))
}

# Recoding scheme for agreement scale
numeric_scheme <- list(
  'Strongly Disagree' = 1,
  'Strongly disagree' = 1,
  'Somewhat disagree' = 2,
  'Neither agree nor disagree' = 3,
  'Somewhat agree' = 4,
  'Strongly agree' = 5,
  'Strongly Agree' = 5
)

df_recoded <- recode_all_columns(df, numeric_scheme, scales_SE)

numeric_scheme_freq <- list(
  'Never' = 1,
  'Seldom' = 2,
  'Sometimes' = 3,
  'Often' = 4,
  'Almost Always' = 5
)

df_recoded <- recode_all_columns(df_recoded, numeric_scheme_freq, scales_exposuretodisccontent)
df_recoded <- recode_all_columns(df_recoded, numeric_scheme_freq, scales_typetodisccontent)
df_recoded <- recode_all_columns(df_recoded, numeric_scheme, scales_otheragreem)

# Reverse Score Items

# Convert columns to numeric
df_recoded[scales_SE_rev] <- lapply(df_recoded[scales_SE_rev], function(x) as.numeric(as.character(x)))

# Apply the reverse scoring
df_recoded[scales_SE_rev] <- lapply(df_recoded[scales_SE_rev], function(x) 6 - x)


# Scale Construction

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SE_of_the_Mean = numeric(),
    StDev = numeric(),
    ITC = numeric(),
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    
    # Filter out rows with any missing values in the scale items
    subset_data_complete_cases <- subset_data[rowSums(is.na(subset_data)) == 0, ]
    
    alpha_results <- alpha(subset_data_complete_cases)
    alpha_val <- alpha_results$total$raw_alpha
    
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"]
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = mean(item_data, na.rm = TRUE),
        SE_of_the_Mean = sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data))),
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,
        Alpha = NA
      ))
    }
    
    # Calculate the scale sum only for rows without missing values
    scale_sum <- rowSums(subset_data_complete_cases, na.rm = FALSE)
    data[[paste0(scale, "_Sum")]] <- NA  # Initialize the column with NA
    data[rowSums(is.na(subset_data)) == 0, paste0(scale, "_Sum")] <- scale_sum  # Assign scale sum for complete cases
    
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = mean(scale_sum, na.rm = TRUE),
      SE_of_the_Mean = sd(scale_sum, na.rm = TRUE) / sqrt(sum(!is.na(scale_sum))),
      StDev = sd(scale_sum, na.rm = TRUE),
      ITC = NA,  # ITC is not applicable for the scale as a whole
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}

scales <- list(
  "Self-Esteem" = scales_SE
)

df_recoded <- df_recoded %>%
  mutate(across(all_of(c(scales_SE, scales_exposuretodisccontent, scales_typetodisccontent, scales_otheragreem)), ~as.numeric(as.character(.))))


alpha_results <- reliability_analysis(df_recoded, scales)

df_recoded <- alpha_results$data_with_scales
df_reliability <- alpha_results$statistics


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


df_freq <- create_frequency_tables(df, scales_freq_cat)

# Descriptive Statistics

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

df_descriptive_stats <- calculate_descriptive_stats(df_recoded, c(scales_freq_cont, scales_exposuretodisccontent, scales_typetodisccontent, scales_otheragreem, "Self-Esteem_Sum"))

# Categorize Self-Esteem

categorize <- function(data, variable_name, cutoffs, labels) {
  # Ensure that the number of labels is one more than the number of cutoffs
  if (length(labels) != length(cutoffs) + 1) {
    stop("The number of labels must be equal to the number of cutoffs plus one.")
  }
  
  # Create a temporary copy of the data to handle NAs
  temp_data <- data
  
  # Create a new variable for categorized data based on the provided variable_name
  categorized_var_name <- paste0("categorized_", variable_name)
  temp_data[[categorized_var_name]] <- cut(
    temp_data[[variable_name]],
    breaks = c(-Inf, cutoffs, Inf), 
    labels = labels,
    include.lowest = TRUE
  )
  
  # Merge the newly categorized column back into the original data
  data[[categorized_var_name]] <- temp_data[[categorized_var_name]]
  
  return(data)
}


cutoffs <- c(17, 30)
labels <- c("Low Self-Esteem", "Moderate Self-Esteem", "High Self-Esteem")
df_recoded <- categorize(df_recoded, "Self-Esteem_Sum", cutoffs, labels)


df_recoded <- df_recoded %>%
  mutate(across(all_of("Age"), ~as.numeric(as.character(.))))


age_cutoffs <- c(25)
labels_age <- c("19-25", "25+")
df_recoded <- categorize(df_recoded, "Age", age_cutoffs, labels_age)

df_SE_Age_freq <- create_frequency_tables(df_recoded, c("categorized_Age", "categorized_Self-Esteem_Sum"))


scales_demog <- c("Gender"                         , 
                  "categorized_Age",
                  "Education.Level"                 ,                                                                                                                                           
                  "Occupation...Selected.Choice")

# One-way ANOVA

# Rename "Self-Esteem_Sum" to "SelfEsteemSum" using base R
names(df_recoded)[names(df_recoded) == "Self-Esteem_Sum"] <- "SelfEsteemSum"


# Levenes Test
library(car)
library(multcomp)
library(nparcomp)
library(dplyr)
library(psych)

perform_levenes_test <- function(data, variables, factors) {
  
  
  
  # Initialize an empty dataframe to store results
  levene_results <- data.frame(Variable = character(), 
                               Factor = character(),
                               F_Value = numeric(), 
                               DF1 = numeric(),
                               DF2 = numeric(),
                               P_Value = numeric(),
                               stringsAsFactors = FALSE)
  
  # Perform Levene's test for each variable and factor
  for (var in variables) {
    for (factor in factors) {
      # Perform Levene's Test
      test_result <- leveneTest(reformulate(factor, response = var), data = data)
      
      # Extract the F value, DF, and p-value
      F_value <- as.numeric(test_result[1, "F value"])
      DF1 <- as.numeric(test_result[1, "Df"])
      DF2 <- as.numeric(test_result[2, "Df"])
      P_value <- as.numeric(test_result[1, "Pr(>F)"])
      
      # Append the results to the dataframe
      levene_results <- rbind(levene_results, 
                              data.frame(Variable = var, 
                                         Factor = factor,
                                         F_Value = F_value, 
                                         DF1 = DF1,
                                         DF2 = DF2,
                                         P_Value = P_value))
    }
  }
  
  return(levene_results)
}

response_vars <- "SelfEsteemSum"
factors <- scales_demog

levene_test_results <- perform_levenes_test(df_recoded, response_vars, factors)

perform_one_way_anova <- function(data, response_vars, factors) {
  anova_results <- data.frame()  # Initialize an empty dataframe to store results
  
  # Iterate over response variables
  for (var in response_vars) {
    # Iterate over factors
    for (factor in factors) {
      # One-way ANOVA
      anova_model <- aov(reformulate(factor, response = var), data = data)
      model_summary <- summary(anova_model)
      
      # Extract relevant statistics
      model_results <- data.frame(
        Variable = var,
        Effect = factor,
        Sum_Sq = model_summary[[1]]$"Sum Sq"[1],  # Extract Sum Sq for the factor
        Mean_Sq = model_summary[[1]]$"Mean Sq"[1],  # Extract Mean Sq for the factor
        Df = model_summary[[1]]$Df[1],  # Extract Df for the factor
        FValue = model_summary[[1]]$"F value"[1],  # Extract F value for the factor
        pValue = model_summary[[1]]$"Pr(>F)"[1]  # Extract p-value for the factor
      )
      
      anova_results <- rbind(anova_results, model_results)
    }
  }
  
  return(anova_results)
}


# Define your response variables (dependent variables) and the factor (independent variable)
response_vars <-  "SelfEsteemSum" # Add more variables as needed
factor <- scales_demog  # Specify the factor (independent variable)

# Perform one-way ANOVA
one_way_anova_results <- perform_one_way_anova(df_recoded, response_vars, factor)

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


descriptive_stats_bygroup <- calculate_means_and_sds_by_factors(df_recoded, "SelfEsteemSum", scales_demog)

## CORRELATION ANALYSIS

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


correlation_matrix <- calculate_correlation_matrix(df_recoded, c("SelfEsteemSum", scales_exposuretodisccontent), method = "spearman")
#correlation_matrix2 <- calculate_correlation_matrix(df_recoded, c("SelfEsteemSum", "ExposureToDiscContent_Sum"), method = "spearman")

colnames(df_recoded)

# Sum the specified items to create a new scale
scales_exposuretodisccontent_updated <- c("Frequency.of.Discriminatory.Content...Facebook",                                                                                                                        
                                          "Frequency.of.Discriminatory.Content...Instagram",                                                                                                                       
                                          "Frequency.of.Discriminatory.Content...Twitter.or.X",                                                                                                                    
                                          "Frequency.of.Discriminatory.Content...TikTok",                                                                                                                         
                                          "Frequency.of.Discriminatory.Content....YouTube",                                                                                                                      
                                          "Frequency.of.Discriminatory.Content....Snapchat",
                                          "Frequency.of.Discriminatory.Content....YouTube.1")

df_recoded$ExposureToDiscContent_Sum <- apply(df_recoded[scales_exposuretodisccontent_updated], 1, function(row) {
  if (any(is.na(row))) {
    return(NA) # Return NA if there's any missing value in the items for this row
  } else {
    return(sum(row)) # Sum the items if there are no missing values
  }
})

correlation_matrix2 <- calculate_correlation_matrix(df_recoded, c("SelfEsteemSum", "ExposureToDiscContent_Sum"), method = "pearson")


# OLS Regression

library(broom)

fit_ols_and_format <- function(data, predictors, response_vars, save_plots = FALSE) {
  ols_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Fit the OLS multiple regression model
    lm_model <- lm(formula, data = data)
    
    # Get the summary of the model
    model_summary <- summary(lm_model)
    
    # Extract the R-squared, F-statistic, and p-value
    r_squared <- model_summary$r.squared
    f_statistic <- model_summary$fstatistic[1]
    p_value <- pf(f_statistic, model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE)
    
    # Print the summary of the model for fit statistics
    print(summary(lm_model))
    
    # Extract the tidy output and assign it to ols_results
    ols_results <- broom::tidy(lm_model) %>%
      mutate(ResponseVariable = response_var, R_Squared = r_squared, F_Statistic = f_statistic, P_Value = p_value)
    
    # Optionally print the tidy output
    print(ols_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lm_model) ~ fitted(lm_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lm_model), main = "Q-Q Plot")
      qqline(residuals(lm_model))
      dev.off()
    }
    
    # Store the results in a list
    ols_results_list[[response_var]] <- ols_results
  }
  
  return(ols_results_list)
}

df_modelresults <- fit_ols_and_format(
  data = df_recoded,
  predictors = c("ExposureToDiscContent_Sum", scales_demog),  # Replace with actual predictor names
  response_vars = "SelfEsteemSum",       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults <- bind_rows(df_modelresults)


# Alternative 1
df_modelresults1 <- fit_ols_and_format(
  data = df_recoded,
  predictors = c("ExposureToDiscContent_Sum"),  # Replace with actual predictor names
  response_vars = "SelfEsteemSum",       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults1 <- bind_rows(df_modelresults1)

# Alternative2

df_modelresults2 <- fit_ols_and_format(
  data = df_recoded,
  predictors = c(scales_exposuretodisccontent_updated),  # Replace with actual predictor names
  response_vars = "SelfEsteemSum",       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe

df_modelresults2 <- bind_rows(df_modelresults2)


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
  "Descriptives" = df_descriptive_stats,
  "Frequency Table" = df_freq,
  "Frequency Table 2" = df_SE_Age_freq,
  "Reliability Analysis" = df_reliability,
  "Levene's Test" = levene_test_results,
  "Anova" = one_way_anova_results,
  "Correlations" = correlation_matrix,
  "Model results" = df_modelresults
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")


