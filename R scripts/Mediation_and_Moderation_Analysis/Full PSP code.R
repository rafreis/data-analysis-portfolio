library(readxl)
library(sem)
library(lavaan)
library(semTools)
library(dplyr)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/davidhoward/davidhoward123-attachments")
psp_data <- read_excel("PSPquantdata.xls")

library(psych)

## DEFINING CODING SCHEMES ##
coding_list <- list(Q17_1 = c("1" = "Male", "2" = "Female", "3" = "Non-Binary", "4" = "Non Confotming", "5" = "Queer Identified"))

## SAMPLE CHARACTERIZATION ##

sample_characterization <- function(data, columns, coding_scheme = NULL) {
  combined_df <- data.frame(
    Variable = character(),
    Value = character(),
    Frequency = numeric(),
    Percentage = numeric(),
    stringsAsFactors = FALSE
  )
  
  for (col in columns) {
    table_data <- table(data[[col]], useNA = "ifany")
    
    # Convert the table into a data frame
    df <- as.data.frame(table_data)
    colnames(df) <- c("Value", "Frequency")
    
    # If there's a coding scheme provided, replace the codes with their corresponding values
    if (!is.null(coding_scheme) && !is.null(coding_scheme[[col]])) {
      df$Value <- as.character(df$Value)
      levels_to_change <- intersect(df$Value, names(coding_scheme[[col]]))
      df$Value[which(df$Value %in% levels_to_change)] <- coding_scheme[[col]][levels_to_change]
    }
    
    df$Percentage <- round((df$Frequency / sum(df$Frequency)) * 100, 1)
    df$Variable <- col
    df <- df[, c("Variable", "Value", "Frequency", "Percentage")]
    
    combined_df <- rbind(combined_df, df)
  }
  
  return(combined_df)
}

sample_stats <- sample_characterization(psp_data, columns = c("Q17_1"), coding_scheme = coding_list)

## Age ##

age_mean <- round(mean(psp_data$Age, na.rm = TRUE), 3)
age_sd <- round(sd(psp_data$Age, na.rm = TRUE), 3)

## RELIABILITY ANALYSIS ##

# Define the scales and corresponding items
scales <- list(
  AMI = c("Q24_1", "Q24_2", "Q24_3", "Q24_4", "Q24_5", "Q24_6", "Q24_7", "Q24_8", "Q24_9", "Q24_10", "Q24_11", "Q24_12", "Q24_13", "Q24_14", "Q24_15", "Q24_16", "Q24_17", "Q24_18", "Q24_19", "Q24_20"),
  ASI = c("Q23_1", "Q23_2", "Q23_3", "Q23_4", "Q23_5", "Q23_6", "Q23_7", "Q23_8", "Q23_9", "Q23_10", "Q23_11", "Q23_12", "Q23_13", "Q23_14", "Q23_15", "Q23_16", "Q23_17", "Q23_18", "Q23_19", "Q23_20", "Q23_21", "Q23_22"),
  MSS = c("Q25_1", "Q25_2", "Q25_3", "Q25_4", "Q25_5", "Q25_6", "Q25_7", "Q25_8"),
  BFI = c("Q26_1", "Q26_2", "Q26_3", "Q26_4", "Q26_5", "Q26_6", "Q26_7", "Q26_8", "Q26_9", "Q26_10", "Q26_11", "Q26_12", "Q26_13", "Q26_14", "Q26_15", "Q26_16", "Q26_17", "Q26_18"),
  SDO = c("Q27_1", "Q27_2", "Q27_3", "Q27_4", "Q27_5", "Q27_6", "Q27_7", "Q27_8", "Q27_9", "Q27_10", "Q27_11", "Q27_12", "Q27_13", "Q27_14", "Q27_15", "Q27_16"),
  RWA = c("Q28_1", "Q28_2", "Q28_3", "Q28_4", "Q28_5", "Q28_6", "Q28_7", "Q28_8", "Q28_9", "Q28_10", "Q28_11", "Q28_12", "Q28_13", "Q28_14", "Q28_15"),
  Agreeableness = c("Q26_1", "Q26_3", "Q26_5", "Q26_7", "Q26_9", "Q26_10", "Q26_12", "Q26_14", "Q26_17"),
  Openness = c("Q26_2", "Q26_4", "Q26_6", "Q26_8", "Q26_11", "Q26_13", "Q26_15", "Q26_16", "Q26_18")
)

# Reverse-score items

# Items to reverse score
reverse_items_5 <- c(
  "Q25_6", "Q25_7", "Q25_8",
  "Q26_1", "Q26_5","Q26_10","Q26_14",
  "Q26_13","Q26_16",
  "Q27_9", "Q27_10", "Q27_11", "Q27_12", "Q27_14", "Q27_15", "Q27_16",
  "Q28_2","Q28_4","Q28_6","Q28_8","Q28_10","Q28_12","Q28_14")

reverse_items_6 <- c("Q23_3", "Q23_6", "Q23_7","Q23_13","Q23_18", "Q23_21")


## ASI AMI - 6
## BFI SDO RWA Agreeableness Openness MSS - 5


# Reverse scoring
psp_data[ , reverse_items_6] <- 7 - psp_data[ , reverse_items_6]
psp_data[ , reverse_items_5] <- 6 - psp_data[ , reverse_items_5]


library(psych)

# Scale refinement - Test if Alpha > 0.700, if not, iteratively drop items with correlations < 0.300

refine_scale_items <- function(data, items) {
  dropped_items <- list()
  
  while (TRUE) {
    result <- alpha(data[, items])
    current_alpha <- result$total$raw_alpha
    
    if (current_alpha >= 0.700) {
      break
    }
    
    
    item_total_correlations <- result$item.stats$r.drop
    min_item <- which.min(item_total_correlations)
    
    if (item_total_correlations[min_item] >= 0.300) {
      break
    }
    
    dropped_items[[items[min_item]]] <- item_total_correlations[min_item]
    items <- items[-min_item]
  }
  
  list(refined_items = items, final_alpha = current_alpha, dropped = dropped_items)
}

refined_results <- lapply(scales, refine_scale_items, data = psp_data)

# Extracting refined items to a separate list
refined_scales_list <- lapply(refined_results, function(x) x$refined_items)


# Print refined results and save the composite scales
for (scale_name in names(refined_results)) {
  result <- refined_results[[scale_name]]
  
  cat(sprintf("Refined Cronbach's Alpha for %s: %f\n", scale_name, result$final_alpha))
  cat(sprintf("Final items for %s: %s\n", scale_name, paste(result$refined_items, collapse=", ")))
  
  if (length(result$dropped) > 0) {
    cat(sprintf("Dropped items and their correlations for %s:\n", scale_name))
    for (item_name in names(result$dropped)) {
      cat(sprintf("%s: %f\n", item_name, result$dropped[[item_name]]))
    }
  }
  
  # Add the composite scale to psp_data
  psp_data[, scale_name] <- rowMeans(psp_data[, result$refined_items], na.rm = TRUE)
  cat("\n")
}

## DESCRIPTIVE STATISTICS - RELIABILITY ##

descriptives_scales <- data.frame(Scale = character(), 
                                  Item = character(), 
                                  Mean = numeric(), 
                                  StandardDeviation = numeric(), 
                                  Alpha = numeric(),
                                  stringsAsFactors = FALSE)

for (scale_name in names(refined_results)) {
  result <- refined_results[[scale_name]]
  
  for (item in result$refined_items) {
    descriptives_scales <- rbind(descriptives_scales, c(scale_name, 
                                                        item, 
                                                        round(mean(psp_data[[item]], na.rm = TRUE), 3),
                                                        round(sd(psp_data[[item]], na.rm = TRUE), 3),
                                                        round(result$final_alpha, 3)))
  }
  
  descriptives_scales <- rbind(descriptives_scales, c(scale_name,
                                                      paste(scale_name, "(Composite Scale)", sep = " "),
                                                      round(mean(psp_data[[scale_name]], na.rm = TRUE), 3),
                                                      round(sd(psp_data[[scale_name]], na.rm = TRUE), 3),
                                                      round(result$final_alpha, 3)))
}

colnames(descriptives_scales) <- c("Scale", "Item", "Mean", "SD", "Cronbach's Alpha")

## MEANS AND STANDARD DEVIATIONS BY CATEGORIES ##

# Initialize an empty dataframe to store the final results
descriptives_by_category <- data.frame(Scale = character(),
                                       stringsAsFactors = FALSE)

# Loop through each scale
for (scale_name in names(refined_results)) {
  
  # Initialize a list to store mean and SD for each category for the current scale
  scale_results <- list(Scale = scale_name)
  
  # Loop through each category column in the coding list
  for (col in names(coding_list)) {
    
    # Loop through each category in the column
    for (category_code in names(coding_list[[col]])) {
      
      category_name <- coding_list[[col]][category_code]
      
      # Filter data for the current category
      filtered_data <- psp_data[[scale_name]][which(psp_data[[col]] == category_code)]
      
      # Check if there are at least 3 cases for this category
      if (length(filtered_data) >= 3) {
        
        # Calculate mean and standard deviation
        mean_val <- round(mean(filtered_data, na.rm = TRUE), 3)
        sd_val <- round(sd(filtered_data, na.rm = TRUE), 3)
        
        # Store results in the list
        scale_results[paste(category_name, "Mean")] <- mean_val
        scale_results[paste(category_name, "SD")] <- sd_val
      }
    }
  }
  
  # Convert list to dataframe and rbind to the final dataframe
  descriptives_by_category <- rbind(descriptives_by_category, as.data.frame(as.list(scale_results)))
}


## CHECK NORMALITY ##

library(moments)

normality_stats_dataframe <- function(data, column_names) {
  
  # Initialize an empty dataframe for results
  results_df <- data.frame(
    Variable = character(), 
    Skewness = numeric(), 
    Kurtosis = numeric(), 
    ShapiroW_p_value = numeric(), 
    stringsAsFactors = FALSE
  )
  
  # Loop through each variable/column
  for (col_name in column_names) {
    
    # Extract the column
    column_data <- data[[col_name]]
    
    # Calculate skewness, kurtosis and Shapiro-Wilk test
    skew_val <- skewness(column_data, na.rm = TRUE)
    kurt_val <- kurtosis(column_data, na.rm = TRUE)
    shapiro_test <- shapiro.test(column_data)
    
    # Append to the results dataframe
    results_df <- rbind(results_df, data.frame(
      Variable = col_name,
      Skewness = round(skew_val, 3),
      Kurtosis = round(kurt_val, 3),
      ShapiroW_p_value = round(shapiro_test$p.value, 3)
    ))
  }
  
  return(results_df)
}

scale_names <- names(refined_results)  # Get the scale names from your refined_results list
normality_stats_result <- normality_stats_dataframe(psp_data, scale_names)
print(normality_stats_result)


## CORRELATION MATRIX ##

generate_scale_correlation_dataframe <- function(data, scale_names, correlation_method = "pearson") {
  # Extract scales from data
  scale_data <- data[, scale_names, drop = FALSE]
  
  # Remove any NA values
  scale_data_complete <- na.omit(scale_data)
  
  # Calculate the correlation matrix based on the chosen method
  if (correlation_method == "pearson") {
    correlation_matrix <- cor(scale_data_complete, method = "pearson")
  } else if (correlation_method == "spearman") {
    correlation_matrix <- cor(scale_data_complete, method = "spearman")
  } else {
    stop("Invalid correlation method. Choose 'pearson' or 'spearman'.")
  }
  
  # Round the values
  correlation_matrix_rounded <- round(correlation_matrix, 3)
  
  # Convert the matrix into a dataframe
  correlation_dataframe <- as.data.frame(correlation_matrix_rounded)
  
  return(correlation_dataframe)
}

scale_list <- names(refined_results)  # Get list of scale names from your refined_results list
correlation_dataframe_result <- generate_scale_correlation_dataframe(psp_data, scale_list, correlation_method = "spearman")
head(correlation_dataframe_result)

## STRUCTURAL EQUATION MODELLING ##

#Recode the Gender Variable and drop cases with small sample size

psp_data <- psp_data %>%
  mutate(Q17_1 = case_when(
    Q17_1 == "1" ~ "1",  # Male -> 1
    Q17_1 == "2" ~ "0",  # Female -> 0
    TRUE ~ NA_character_  # Set other genders to NA
  ))

# Convert to numeric
psp_data$Q17_1 <- as.numeric(psp_data$Q17_1)

#Rename
colnames(psp_data)[colnames(psp_data) == "Q17_1"] <- "Gender"

#Rename Observed scales to avoid conflict with latent factors naming

psp_data <- psp_data %>%
  rename(
    observed_AMI = AMI,
    observed_ASI = ASI,
    observed_MSS = MSS,
    observed_BFI = BFI,
    observed_SDO = SDO,
    observed_RWA = RWA,
    observed_Agreeableness = Agreeableness,
    observed_Openness = Openness
  )

# Generate the measurement model syntax
generate_measurement_syntax <- function(scales_list) {
  syntax <- "# Measurement Model"
  for (scale_name in names(scales_list)) {
    items <- scales_list[[scale_name]]
    item_syntax <- paste(items, collapse = " + ")
    syntax <- paste(syntax, sprintf("%s =~ %s", scale_name, item_syntax), sep = "\n")
  }
  return(syntax)
}

measurement_model <- generate_measurement_syntax(refined_scales_list)

# Define each structural model
model_AMI_personality <- "
# Structural Model 1
AMI ~ SDO
AMI ~ RWA
RWA ~ Openness
SDO ~ Agreeableness
Agreeableness ~ Gender
SDO ~ Gender
"

model_AMI_SocialPsychology <- "
# Structural Model 2
SDO ~ RWA
RWA ~ Openness
SDO ~ Agreeableness
Agreeableness ~ Gender
SDO ~ Gender
AMI ~ Gender
"

model_AMI_Combined <- "
# Structural Model 3
AMI ~ SDO
AMI ~ RWA
SDO ~ RWA
RWA ~ Openness
SDO ~ Agreeableness
Agreeableness ~ Gender
SDO ~ Gender
AMI ~ Gender
"

model_ASI_personality <- "
# Structural Model 4
ASI ~ SDO
SDO ~ RWA
ASI ~ RWA
RWA ~ Openness
SDO ~ Agreeableness
Agreeableness ~ Gender
SDO ~ Gender
"

model_ASI_SocialPsychology <- "
# Structural Model 5
SDO ~ RWA
RWA ~ Openness
SDO ~ Agreeableness
Agreeableness ~ Gender
SDO ~ Gender
ASI ~ Gender
"

model_ASI_Combined <- "
# Structural Model 6
ASI ~ SDO
ASI + SDO ~ RWA
RWA ~ Openness
SDO ~ Agreeableness
Agreeableness ~ Gender
SDO ~ Gender
ASI ~ Gender
"

# Combine Sintaxes

model_AMI_personality <- paste(measurement_model, model_AMI_personality)
model_AMI_SocialPsychology <- paste(measurement_model, model_AMI_SocialPsychology)
model_AMI_Combined <- paste(measurement_model, model_AMI_Combined)
model_ASI_personality <- paste(measurement_model, model_ASI_personality)
model_ASI_SocialPsychology <- paste(measurement_model, model_ASI_SocialPsychology)
model_ASI_Combined <- paste(measurement_model, model_ASI_Combined)


# Fit Models

# Get summary of fit measures
summary_fit <- summary(fit_AMI_personality, fit.measures=TRUE, standardized=TRUE)

print(summary_fit)

factor_loadings <- standardizedSolution(fit)
print(factor_loadings$loadings)

reg_coefficients <- standardizedSolution(fit)
print(reg_coefficients$regressions)


rsquare <- inspect(fit, what="rsquare")
print(rsquare)

