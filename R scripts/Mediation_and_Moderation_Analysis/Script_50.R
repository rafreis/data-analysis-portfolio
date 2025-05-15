# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/reem")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Cleandata.xlsx")

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
    
    # Calculate the mean score for the scale and add as a new column
    scale_mean <- rowMeans(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_mean
    scale_mean_overall <- mean(scale_mean, na.rm = TRUE)
    scale_sem <- sd(scale_mean, na.rm = TRUE) / sqrt(sum(!is.na(scale_mean)))
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = scale_mean_overall,
      SEM = scale_sem,
      StDev = sd(scale_mean, na.rm = TRUE),
      ITC = NA,  # ITC is not applicable for the total scale
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}


library(fastDummies)

vars <- c("Gender", "Level.of.proficiency.in.English", 
          "Experience.with.Technology", "Use.of.AI.for.Academics")

df_recoded <- dummy_cols(df, select_columns = vars)

names(df_recoded) <- gsub(" ", "_", names(df_recoded))
names(df_recoded) <- gsub("\\(", "_", names(df_recoded))
names(df_recoded) <- gsub("\\)", "_", names(df_recoded))
names(df_recoded) <- gsub("\\-", "_", names(df_recoded))
names(df_recoded) <- gsub("/", "_", names(df_recoded))
names(df_recoded) <- gsub("\\\\", "_", names(df_recoded)) 
names(df_recoded) <- gsub("\\?", "", names(df_recoded))
names(df_recoded) <- gsub("\\'", "", names(df_recoded))
names(df_recoded) <- gsub("\\,", "_", names(df_recoded))



colnames(df_recoded)


library(dplyr)


recode_all_columns <- function(df, recoding_scheme, columns = NULL) {
  if (is.null(columns)) {
    columns <- names(df)
  }
  
  df %>%
    mutate(across(all_of(columns), function(x) {
      sapply(x, function(value) {
        value_char <- as.character(value)  # Convert value to character for matching
        if (value_char %in% names(recoding_scheme)) {
          return(recoding_scheme[[value_char]])  # Apply recoding if match found
        } else {
          return(value)  # Return original value if no match
        }
      })
    }))
}

# Recoding scheme for agreement scale
numeric_scheme <- list(
  'strongly disagree' = 1,
  'disagree' = 2,
  'neutral' = 3,
  'agree' = 4,
  'strongly agree' = 5
)


# Recoding the specified columns in the sample dataframe
df_recoded <- recode_all_columns(df_recoded, numeric_scheme)


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

scales_freq <- c("Gender" ,                                         
 "Major"                   ,                        
 "Year.in.college"          ,                       
 "Level.of.proficiency.in.English"  ,               
 "Experience.with.Technology"        ,              
 "Use.of.AI.for.Academics"      ,
 "How.often.do.you.use.AI.for.educational.purposes",
 "comfortable.using.AI.in.studies"  )


df_freq <- create_frequency_tables(df, scales_freq)




scales <- list(
 
  "Institutional Factors" = c("instructors.support.AI", "Uni_coll..Supports.AI", 
                              "AI.incorporated.in.curriculum.", "Im.taught.how.to.use.AI", 
                              "My.uni..clarified.its.position.on.the.use.of.AI",
                              "AI.is.readily.available.at.Uni", "I.receive.AI.support"),
  "Socio-Cultural Factors" = c("Friends.and.peers.use.AI", "AI.is.normal.and.acceptable", 
                               "AI.is.acceptable.in.my.culture."),
  
  "Perceived Usefulness" = c("AI.enhance.learning.experience", "AI.beneficial.for.academic.success.",
                             "AI.beneficial.for.my.field.of.study", "AI.prepares.for.future.career"),
  "Perceived Ease of Use" = c("comfortable.using.AI.in.studies", "AI.easy.to.use", 
                              "confident.in.using.AI", "AI.doesnt.require.a.lot.of.effort"),
  "Perceived Risk" = c("AI.is.risky_dangerous", "AI.lead.to.plagiarism", 
                       "AI.might.get.me.in.trouble", "AI.tools.might.give.me.incorrect.information.")
)

vars <- c("instructors.support.AI", "Uni_coll..Supports.AI", 
                                        "AI.incorporated.in.curriculum.", "Im.taught.how.to.use.AI", 
                                        "My.uni..clarified.its.position.on.the.use.of.AI",
                                        "AI.is.readily.available.at.Uni", "I.receive.AI.support",
            "Friends.and.peers.use.AI", "AI.is.normal.and.acceptable", 
                                         "AI.is.acceptable.in.my.culture.",
            
            "AI.enhance.learning.experience", "AI.beneficial.for.academic.success.",
                                       "AI.beneficial.for.my.field.of.study", "AI.prepares.for.future.career",
           "comfortable.using.AI.in.studies", "AI.easy.to.use", 
                                        "confident.in.using.AI", "AI.doesnt.require.a.lot.of.effort",
            "AI.is.risky_dangerous", "AI.lead.to.plagiarism", 
                                 "AI.might.get.me.in.trouble", "AI.tools.might.give.me.incorrect.information.")


colnames(df_recoded)

df_recoded$Age <- as.numeric(df_recoded$Age)

# Convert the specified variables to numeric
# Convert variables to numeric
df_recoded <- df_recoded %>%
  mutate(across(all_of(vars), ~as.numeric(as.character(.))))

# Function to remove attributes and convert to numeric
df_recoded <- df_recoded %>%
  mutate(across(all_of(vars), ~ {
    x <- .  # Get the current column
    attributes(x) <- NULL  # Remove all attributes
    as.numeric(x)  # Convert to numeric
  }))

str(df_recoded)

alpha_results <- reliability_analysis(df_recoded, scales)

df_descriptives <- alpha_results$statistics

ivs_direct <- c("Age",    "Gender_male"    ,                               
                 "Level.of.proficiency.in.English_advanced"     ,    
                 "Level.of.proficiency.in.English_intermediate"  ,   "Level.of.proficiency.in.English_proficient"      ,
                 "Experience.with.Technology_advanced"            ,  
                 "Experience.with.Technology_expert"          ,      "Experience.with.Technology_intermediate"         ,
                 "Use.of.AI.for.Academics_yes"   )


dv <- "How.often.do.you.use.AI.for.educational.purposes"

# Define the recoding scheme using a named vector
recode_scheme <- c("never" = 1, "rarely" = 2, "sometimes" = 3, "always" = 4)

# Recode the variable in the dataframe
df_recoded <- df_recoded %>%
  mutate(`How.often.do.you.use.AI.for.educational.purposes` = 
           ifelse(is.na(`How.often.do.you.use.AI.for.educational.purposes`), 
                  "never", 
                  `How.often.do.you.use.AI.for.educational.purposes`))

# Apply the recoding scheme
df_recoded <- df_recoded %>%
  mutate(`How.often.do.you.use.AI.for.educational.purposes` =
           recode(`How.often.do.you.use.AI.for.educational.purposes`,
                  !!!recode_scheme))

df_recoded <- df_recoded %>%
  mutate(Level.of.proficiency.in.English_adv = case_when(
    Level.of.proficiency.in.English == "advanced" | Level.of.proficiency.in.English == "proficient" ~ 1,
    TRUE ~ 0
  ))

df_recoded <- df_recoded %>%
  mutate(Experience.with.Technology_adv = case_when(
    Experience.with.Technology == "expert" | Experience.with.Technology == "advanced" ~ 1,
    TRUE ~ 0
  ))


colnames(df_recoded)

scales <- list(
  
  "Institutional Factors" = c("instructors.support.AI", "Uni_coll..Supports.AI", 
                              "AI.incorporated.in.curriculum.", "Im.taught.how.to.use.AI", 
                              "My.uni..clarified.its.position.on.the.use.of.AI",
                              "AI.is.readily.available.at.Uni", "I.receive.AI.support"),
  "Socio-Cultural Factors" = c("Friends.and.peers.use.AI", "AI.is.normal.and.acceptable", 
                               "AI.is.acceptable.in.my.culture."),
  
  "Perceived Usefulness" = c("AI.enhance.learning.experience", "AI.beneficial.for.academic.success.",
                             "AI.beneficial.for.my.field.of.study", "AI.prepares.for.future.career"),
  "Perceived Ease of Use" = c("comfortable.using.AI.in.studies", "AI.easy.to.use", 
                              "confident.in.using.AI", "AI.doesnt.require.a.lot.of.effort"),
  "Perceived Risk" = c("AI.is.risky_dangerous", "AI.lead.to.plagiarism", 
                       "AI.might.get.me.in.trouble", "AI.tools.might.give.me.incorrect.information.")
)


library(lavaan)
library(dplyr)
library(lavaan)
library(dplyr)

# Define the SEM model
model <- '
  # Latent Variables (Measurement Model)
  Institutional_Factors =~ instructors.support.AI + Uni_coll..Supports.AI + AI.incorporated.in.curriculum. + Im.taught.how.to.use.AI + AI.is.readily.available.at.Uni + I.receive.AI.support
  Socio_Cultural_Factors =~ Friends.and.peers.use.AI + AI.is.normal.and.acceptable + AI.is.acceptable.in.my.culture.
  Perceived_Usefulness =~ AI.enhance.learning.experience + AI.beneficial.for.academic.success. + AI.beneficial.for.my.field.of.study + AI.prepares.for.future.career
  Perceived_Ease_of_Use =~ comfortable.using.AI.in.studies + AI.easy.to.use + confident.in.using.AI + AI.doesnt.require.a.lot.of.effort
  Perceived_Risk =~ AI.is.risky_dangerous + AI.lead.to.plagiarism + AI.might.get.me.in.trouble + AI.tools.might.give.me.incorrect.information.

  # Direct Predictors of the DV
  How.often.do.you.use.AI.for.educational.purposes ~ Age + Gender_male + Level.of.proficiency.in.English_adv + Experience.with.Technology_adv

  # Latent Constructs also Predict the DV
  How.often.do.you.use.AI.for.educational.purposes ~ b1*Institutional_Factors + b2*Socio_Cultural_Factors + b3*Perceived_Usefulness + b4*Perceived_Ease_of_Use + b5*Perceived_Risk
'

# Fit the model using df_recoded
fit <- sem(model, data = df_recoded, missing = "listwise")

# Summarize results with standardized estimates and fit measures
results <- summary(fit, standardized = TRUE, fit.measures = TRUE)

# Extract coefficients
coefficients_df <- parameterEstimates(fit, standardized = TRUE, se = TRUE, ci = TRUE, header = TRUE, rsquare = TRUE)

# Extract fit measures
fit_measures <- fitMeasures(fit, fit.measures = "all")
fit_measures_df <- as.data.frame(fit_measures)

# Print results
print(results)
print(fit_measures_df)

library(lavaan)
library(dplyr)
library(writexl)


# Separate different result sections
latent_variables_df <- coefficients_df %>% filter(op == "=~")  # Factor loadings
regressions_df <- coefficients_df %>% filter(op == "~")  # Regression paths
covariances_df <- coefficients_df %>% filter(op == "~~")  # Covariances
variances_df <- coefficients_df %>% filter(op == "~~" & lhs == rhs)  # Variances

# Create a list of all dataframes for export
results_list <- list(
  "Sample" = df_freq,
  "Reliability" = df_descriptives,
  "Fit Measures" = fit_measures_df,
  "Latent Variables" = latent_variables_df,
  "Regressions" = regressions_df,
  "Covariances" = covariances_df,
  "Variances" = variances_df
)

# Export all results to an Excel file
write_xlsx(results_list, "Results.xlsx")

# Print message for user
print("SEM results exported to 'SEM_Results.xlsx' successfully.")


library(DiagrammeR)

grViz("
digraph SEM {
  
  # Graph layout: Left to right (horizontal)
  graph [rankdir = LR, splines=polyline]

  # Node styles (White background, black borders, uniform font size)
  node [shape = ellipse, style=filled, fillcolor=white, fontcolor=black, color=black, fontsize=12]
  Institutional_Factors [label=<<i>Institutional Factors</i><br/>R² = 0.79>]
  Socio_Cultural_Factors [label=<<i>Socio-Cultural Factors</i><br/>R² = 0.35>]
  Perceived_Usefulness [label=<<i>Perceived Usefulness</i><br/>R² = 0.33>]
  Perceived_Ease_of_Use [label=<<i>Perceived Ease of Use</i><br/>R² = 0.55>]
  Perceived_Risk [label=<<i>Perceived Risk</i><br/>R² = 0.57>]

  node [shape = box, style=filled, fillcolor=white, fontcolor=black, color=black, fontsize=12]
  AI_Adoption [label='AI Usage\nR² = 0.81']
  Age
  Male
  Tech_Experience
  English_Proficiency

  # Arrows with standardized coefficients (rounded to three decimals), using explicit p-values
  edge [fontsize=12, penwidth=1.5]

  Institutional_Factors -> AI_Adoption [label='b = -0.053, p = .395']
  Socio_Cultural_Factors -> AI_Adoption [label='b = 0.033, p = .768']
  Perceived_Usefulness -> AI_Adoption [label='b = 0.054, p = .904']
  Perceived_Ease_of_Use -> AI_Adoption [label='b = 0.312, p = .558']
  Perceived_Risk -> AI_Adoption [label='b = -0.051, p = .415']

  Age -> AI_Adoption [label='b = 0.080, p = .051']
  Male -> AI_Adoption [label='b = -0.171, p < .001']
  Tech_Experience -> AI_Adoption [label='b = 0.125, p = .003']
  English_Proficiency -> AI_Adoption [label='b = -0.016, p = .689']
}
")



library(tikzDevice)

tikz("sem_model.tex", standAlone = TRUE)

cat("\\begin{tikzpicture}[->,>=stealth',auto,node distance=3cm,main node/.style={circle,draw}]")

cat("\\node[main node] (1) {Institutional Factors};")
cat("\\node[main node] (2) [right of=1] {Socio-Cultural Factors};")
cat("\\node[main node] (3) [right of=2] {Perceived Usefulness};")
cat("\\node[main node] (4) [right of=3] {Perceived Ease of Use};")
cat("\\node[main node] (5) [right of=4] {Perceived Risk};")
cat("\\node[main node] (6) [below of=3] {AI Adoption};")

cat("\\path[every node/.style={font=\\sffamily\\small}]")
cat("(1) edge node {b1 = -0.05} (6)")
cat("(2) edge node {b2 = 0.03} (6)")
cat("(3) edge node {b3 = 0.05} (6)")
cat("(4) edge node {b4 = 0.31} (6)")
cat("(5) edge node {b5 = -0.05} (6);")

cat("\\end{tikzpicture}")

dev.off()
