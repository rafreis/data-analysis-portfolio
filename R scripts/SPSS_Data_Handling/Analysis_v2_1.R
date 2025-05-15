setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/gaelleferes")

library(haven)
df <- read_sav("Student level good merged school level teacher 201123.sav")

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

library(dplyr)

# School-level statistics
df_school_stats <- df %>%
  group_by(NAME_CNT, ID_SCHOOL) %>%
  summarise(
    School_Country = first(NAME_CNT),
    School_Avg_PisaMath = mean(PisaMath, na.rm = TRUE),
    School_Avg_PisaReading = mean(PisaReading, na.rm = TRUE),
    School_Avg_PisaScience = mean(PisaScience, na.rm = TRUE),
    School_Student_Count = n_distinct(ID_STUDENT),
    School_Avg_director_leadership = mean(director_leadership, na.rm = TRUE),
    School_Avg_AVGESCS = mean(AVGESCS, na.rm = TRUE),
    School_Job_Satisfaction = first(JobSatisfaction_mean_mean),
    .groups = 'drop'
  )

# Country-level statistics
df_country_stats <- df %>%
  group_by(NAME_CNT) %>%
  summarise(
    Country = first(NAME_CNT),
    Avg_PisaMath = mean(PisaMath, na.rm = TRUE),
    SD_PisaMath = sd(PisaMath, na.rm = TRUE),
    Avg_PisaReading = mean(PisaReading, na.rm = TRUE),
    SD_PisaReading = sd(PisaReading, na.rm = TRUE),
    Avg_PisaScience = mean(PisaScience, na.rm = TRUE),
    SD_PisaScience = sd(PisaScience, na.rm = TRUE),
    Desk_Percentage = mean(InyourhomeAdesktostudyat == 1, na.rm = TRUE),
    Room_Percentage = mean(InyourhomeAroomofyourown == 1, na.rm = TRUE),
    QuietPlace_Percentage = mean(InyourhomeAquietplacetostudy == 1, na.rm = TRUE),
    Computer_Percentage = mean(InyourhomeAcomputeryoucanuseforschoolwork == 1, na.rm = TRUE),
    Internet_Percentage = mean(InyourhomeAlinktotheInternet == 1, na.rm = TRUE),
    CNT_PISA_read = first(CNT_PISA_read),
    CNT_PISA_Math = first(CNT_PISA_Math),
    CNT_PISA_science = first(CNT_PISA_science),
    CNT_ESCS = first(CNT_ESCS),
    CNT_SHW_SELFDIRECT = first(CNT_SHW_SELFDIRECT),
    CNT_SHW_POWER = first(CNT_SHW_POWER),
    CNT_SHW_SECUR = first(CNT_SHW_SECUR),
    CNT_SHW_HEDONISM = first(CNT_SHW_HEDONISM),
    CNT_SHW_BENEVOLENCE = first(CNT_SHW_BENEVOLENCE),
    CNT_SHW_ACHIEV = first(CNT_SHW_ACHIEV),
    CNT_SHW_STIMUL = first(CNT_SHW_STIMUL),
    CNT_SHW_CONFORM = first(CNT_SHW_CONFORM),
    CNT_SHW_SPIRIT = first(CNT_SHW_SPIRIT),
    CNT_SHW_TRAD = first(CNT_SHW_TRAD),
    CNT_Gov_expend = first(CNT_Gov_expend),
    Govexp_GDPDol = first(Govexp_GDPDol),
    CNT_GDP_DOLLAR = first(CNT_GDP_DOLLAR),                           
    CNT_GDP_PERCAPITA = first(CNT_GDP_PERCAPITA), 
    .groups = 'drop'
  )

# Merging school_stats and country_stats
final_stats <- left_join(df_country_stats, df_school_stats, by = "NAME_CNT")

# Calculate the final required statistics
df_final_descriptives <- final_stats %>%
  group_by(NAME_CNT) %>%
  summarise(
    Country = first(Country),
    School_Count = n_distinct(ID_SCHOOL),
    Student_Count = sum(School_Student_Count),
    Country_Avg_director_leadership = mean(School_Avg_director_leadership, na.rm = TRUE),
    SD_director_leadership = sd(School_Avg_director_leadership, na.rm = TRUE),
    Country_Avg_AVGESCS = mean(School_Avg_AVGESCS, na.rm = TRUE),
    SD_AVGESCS = sd(School_Avg_AVGESCS, na.rm = TRUE),
    Country_Job_Satisfaction = mean(School_Job_Satisfaction, na.rm = TRUE),
    SD_Job_Satisfaction = sd(School_Job_Satisfaction, na.rm = TRUE),
    Avg_PisaMath = first(Avg_PisaMath),
    SD_PisaMath = first(SD_PisaMath),
    Avg_PisaReading = first(Avg_PisaReading),
    SD_PisaReading = first(SD_PisaReading),
    Avg_PisaScience = first(Avg_PisaScience),
    SD_PisaScience = first(SD_PisaScience),
    Desk_Percentage = first(Desk_Percentage),
    Room_Percentage = first(Room_Percentage),
    QuietPlace_Percentage = first(QuietPlace_Percentage),
    Computer_Percentage = first(Computer_Percentage),
    Internet_Percentage = first(Internet_Percentage),
    CNT_PISA_read = first(CNT_PISA_read),
    CNT_PISA_Math = first(CNT_PISA_Math),
    CNT_PISA_science = first(CNT_PISA_science),
    CNT_ESCS = first(CNT_ESCS),
    CNT_SHW_SELFDIRECT = first(CNT_SHW_SELFDIRECT),
    CNT_SHW_POWER = first(CNT_SHW_POWER),
    CNT_SHW_SECUR = first(CNT_SHW_SECUR),
    CNT_SHW_HEDONISM = first(CNT_SHW_HEDONISM),
    CNT_SHW_BENEVOLENCE = first(CNT_SHW_BENEVOLENCE),
    CNT_SHW_ACHIEV = first(CNT_SHW_ACHIEV),
    CNT_SHW_STIMUL = first(CNT_SHW_STIMUL),
    CNT_SHW_CONFORM = first(CNT_SHW_CONFORM),
    CNT_SHW_SPIRIT = first(CNT_SHW_SPIRIT),
    CNT_SHW_TRAD = first(CNT_SHW_TRAD),
    CNT_Gov_expend = first(CNT_Gov_expend),
    Govexp_GDPDol = first(Govexp_GDPDol),
    CNT_GDP_DOLLAR = first(CNT_GDP_DOLLAR),                           
    CNT_GDP_PERCAPITA = first(CNT_GDP_PERCAPITA), 
    .groups = 'drop'
  )

#Log-transformation

dvs <- c("CNT_GDP_DOLLAR"        ,          "CNT_GDP_PERCAPITA")

log_transform_dvs <- function(data, dvs) {
  for (dv in dvs) {
    # Creating a new column name for the log-transformed variable
    new_col_name <- paste("log", dv, sep = "_")
    
    # Applying the log transformation
    # Adding 1 to avoid log(0) which is undefined
    data[[new_col_name]] <- log(data[[dv]] + 1)
  }
  return(data)
}

# Applying the function to your data
df_final_descriptives <- log_transform_dvs(df_final_descriptives, dvs)

# Scatterplots

library(ggplot2)
library(tidyr)
library(gridExtra)
library(ggrepel)

# Reshape the data for plotting
long_df <- gather(df_final_descriptives, key = "measure", value = "value", 
                  Avg_PisaMath, Avg_PisaReading, Avg_PisaScience)

# Define the y-axes
y_axes <- c("Country_Avg_AVGESCS", "Country_Job_Satisfaction", "Country_Avg_director_leadership", "Desk_Percentage", 
            "Room_Percentage", "QuietPlace_Percentage", "Computer_Percentage", 
            "Internet_Percentage", "CNT_SHW_SELFDIRECT", "CNT_SHW_POWER", 
            "CNT_SHW_SECUR", "CNT_SHW_HEDONISM", "CNT_SHW_BENEVOLENCE", 
            "CNT_SHW_ACHIEV", "CNT_SHW_STIMUL", "CNT_SHW_CONFORM", 
            "CNT_SHW_SPIRIT", "CNT_SHW_TRAD", "CNT_Gov_expend", "Govexp_GDPDol", "log_CNT_GDP_DOLLAR" ,                          
            "log_CNT_GDP_PERCAPITA")

# Define a color palette for the Pisa scores
color_palette <- c("blue", "green", "red")  # Adjust these colors as needed

# List to store plots
plots <- list()

# Create and save scatter plots with different colors for each Pisa score
for(y_axis in y_axes) {
  plot_name <- paste0("scatterplot_", y_axis, ".jpeg")
  p <- ggplot(long_df, aes_string(x = "value", y = y_axis, label = "NAME_CNT")) +
    geom_point(aes(color = measure)) +  # Color by measure (Pisa scores)
    geom_text_repel(size = 2) +
    facet_wrap(~ measure, scales = "free") +
    labs(x = "Average PISA Score", y = y_axis) +
    theme_minimal() +
    theme(strip.text = element_blank()) +
    scale_color_manual(values = color_palette)  # Apply color palette
  
  plots[[y_axis]] <- p
  
  # Save the plot
  ggsave(plot_name, plot = p, width = 10, height = 6)
}

# Optionally print all plots
for(plot in plots) {
  print(plot)
}

# Calculate ICC
library(lme4)
library(performance)

# Fit mixed-effects models for ICC
model_country_math <- lmer(PisaMath ~ 1 + (1 | NAME_CNT), data = df, REML = FALSE)
model_school_math <- lmer(PisaMath ~ 1 + (1 | ID_SCHOOL), data = df, REML = FALSE)

model_country_reading <- lmer(PisaReading ~ 1 + (1 | NAME_CNT), data = df, REML = FALSE)
model_school_reading <- lmer(PisaReading ~ 1 + (1 | ID_SCHOOL), data = df, REML = FALSE)

model_country_science <- lmer(PisaScience ~ 1 + (1 | NAME_CNT), data = df, REML = FALSE)
model_school_science <- lmer(PisaScience ~ 1 + (1 | ID_SCHOOL), data = df, REML = FALSE)

# Calculate ICC
icc_country_math <- icc(model_country_math)
icc_school_math <- icc(model_school_math)

icc_country_reading <- icc(model_country_reading)
icc_school_reading <- icc(model_school_reading)

icc_country_science <- icc(model_country_science)
icc_school_science <- icc(model_school_science)

# Print ICC results
print(icc_country_math)
print(icc_school_math)
print(icc_country_reading)
print(icc_school_reading)
print(icc_country_science)
print(icc_school_science)

# Outlier Evaluation

vars_outliers <- c("Country_Avg_director_leadership",         
                   "Country_Avg_AVGESCS"    ,          "Country_Job_Satisfaction"   ,    
                   "Avg_PisaMath"        ,               
                   "Avg_PisaReading"          ,       "Avg_PisaScience"              ,  
                    "Desk_Percentage"       ,          "Room_Percentage"               , 
                   "QuietPlace_Percentage"      ,     "Computer_Percentage"    ,         "Internet_Percentage"            ,
                   "CNT_PISA_read"               ,    "CNT_PISA_Math"           ,        "CNT_PISA_science"               ,
                   "CNT_ESCS"                     ,   "CNT_SHW_SELFDIRECT"       ,       "CNT_SHW_POWER"                  ,
                   "CNT_SHW_SECUR"          ,         "CNT_SHW_HEDONISM"          ,      "CNT_SHW_BENEVOLENCE"            ,
                   "CNT_SHW_ACHIEV"          ,        "CNT_SHW_STIMUL"             ,     "CNT_SHW_CONFORM"                ,
                   "CNT_SHW_SPIRIT"           ,       "CNT_SHW_TRAD"                ,    "CNT_Gov_expend"                 ,
                   "Govexp_GDPDol", "CNT_GDP_DOLLAR"        ,          "CNT_GDP_PERCAPITA")
library(dplyr)

calculate_z_scores <- function(data, vars, id_var) {
  # Calculate z-scores for each variable
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~scale(.) %>% as.vector, .names = "z_{.col}")) %>%
    select(!!sym(id_var), starts_with("z_"))
  
  return(z_score_data)
}


id_var <- "NAME_CNT"
df_z_scores <- calculate_z_scores(df_final_descriptives, vars_outliers, id_var)
dvs_z <- c("log_CNT_GDP_DOLLAR"        ,          "log_CNT_GDP_PERCAPITA")
df_z_scores_log <- calculate_z_scores(df_final_descriptives, dvs_z, id_var)

# Change values in the original dataframe to log values

library(dplyr)

# Joining df_final_descriptives with df
df_updated <- df %>%
  left_join(select(df_final_descriptives, NAME_CNT, log_CNT_GDP_DOLLAR, log_CNT_GDP_PERCAPITA), by = "NAME_CNT") %>%
  # Replace the values
  mutate(
    CNT_GDP_DOLLAR = log_CNT_GDP_DOLLAR,
    CNT_GDP_PERCAPITA = log_CNT_GDP_PERCAPITA
  )

## CORRELATION ANALYSIS

vars_model2 <- c("InyourhomeAdesktostudyat"       ,           "InyourhomeAroomofyourown"     ,            
                 "InyourhomeAquietplacetostudy"    ,          "InyourhomeAcomputeryoucanuseforschoolwork", "InyourhomeAlinktotheInternet"    ,         
                 
                 
                 "AVGESCS_mean_mean"                  ,       "JobSatisfaction_mean_mean",
                 "director_leadership"            ,         "CNT_SHW_SELFDIRECT"     ,                  
                 "CNT_SHW_POWER"                   ,          "CNT_SHW_SECUR"         ,                    "CNT_SHW_HEDONISM"   ,                      
                 "CNT_SHW_BENEVOLENCE"              ,         "CNT_SHW_ACHIEV"         ,                   "CNT_SHW_STIMUL"      ,                     
                 "CNT_SHW_CONFORM"                   ,        "CNT_SHW_SPIRIT"          ,                  "CNT_SHW_TRAD"         ,                    
                 "CNT_Gov_expend"                     ,       "Govexp_GDPDol"            ,                 "log_CNT_GDP_DOLLAR"        ,                   
                 "log_CNT_GDP_PERCAPITA")


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


correlation_matrix <- calculate_correlation_matrix(df_updated, vars_model2, method = "pearson")

# Mixed Effects Models
vars_model1 <- c("JobSatisfaction_mean_mean",
                "director_leadership")


vars_model2 <- c("InyourhomeAdesktostudyat"       ,           "InyourhomeAroomofyourown"     ,            
                 "InyourhomeAquietplacetostudy"    ,          "InyourhomeAcomputeryoucanuseforschoolwork", "InyourhomeAlinktotheInternet"    ,         
                                               
                
                 "AVGESCS_mean_mean"                  ,       "JobSatisfaction_mean_mean",
                 "director_leadership"            ,         "CNT_SHW_SELFDIRECT"     ,                  
                 "CNT_SHW_POWER"                   ,          "CNT_SHW_SECUR"         ,                    "CNT_SHW_HEDONISM"   ,                      
                 "CNT_SHW_BENEVOLENCE"              ,         "CNT_SHW_ACHIEV"         ,                   "CNT_SHW_STIMUL"      ,                     
                 "CNT_SHW_CONFORM"                   ,        "CNT_SHW_SPIRIT"          ,                  "CNT_SHW_TRAD"         ,                    
                 "CNT_Gov_expend"                     ,       "Govexp_GDPDol"            ,                 "log_CNT_GDP_DOLLAR"        ,                   
                 "log_CNT_GDP_PERCAPITA")


vars_model2_nocollinear <- c("InyourhomeAdesktostudyat"       ,           "InyourhomeAroomofyourown"     ,            
                 "InyourhomeAquietplacetostudy"    ,          "InyourhomeAcomputeryoucanuseforschoolwork", "InyourhomeAlinktotheInternet"    ,         
                 
                 
                 "AVGESCS_mean_mean"                  ,       "JobSatisfaction_mean_mean",
                 "director_leadership"            ,         "CNT_SHW_SELFDIRECT"     ,                  
                 "CNT_SHW_SECUR"         ,                    "CNT_SHW_HEDONISM"   ,                      
                 "CNT_SHW_BENEVOLENCE"              ,                            
                 "CNT_SHW_CONFORM"                   ,        "CNT_SHW_SPIRIT"          ,                  "CNT_SHW_TRAD"         ,                    
                 "CNT_Gov_expend"                     ,       "Govexp_GDPDol"            ,                 "log_CNT_GDP_DOLLAR"        ,                   
                 "log_CNT_GDP_PERCAPITA")

library(lme4)
library(broom.mixed)
library(afex)
library(car)

df$NAME_CNT <- as.factor(df$NAME_CNT)
df$ID_SCHOOL <- as.factor(df$ID_SCHOOL)

df_school_stats$NAME_CNT <- as.factor(df_school_stats$NAME_CNT)

# Function to fit LMM with multiple random effects and return formatted results
fit_lmm_and_format <- function(formula, data, save_plots = FALSE, plot_param = "") {
  # Fit the linear mixed-effects model
  lmer_model <- lmer(formula, data = data)
  
  # Print the summary of the model for fit statistics
  model_summary <- summary(lmer_model)
  print(model_summary)
  
  # Extract the tidy output for fixed effects and assign it to lmm_results
  lmm_results <- tidy(lmer_model, "fixed")
  
  # Optionally print the tidy output for fixed effects
  print(lmm_results)
  
  # Extract random effects variances
  random_effects_variances <- as.data.frame(VarCorr(lmer_model))
  print(random_effects_variances)
  
  # Initialize vif_values as NULL
  vif_values <- NULL
  
  # Calculate and print VIF for fixed effects if there are multiple fixed effects
  if (length(fixef(lmer_model)) > 1) {
    vif_values <- vif(lmer_model)
    print(vif_values)
  }
  
  # Generate and save residual plots
  if (save_plots) {
    residual_plot_filename <- paste0("Residuals_vs_Fitted_", plot_param, ".jpeg")
    qq_plot_filename <- paste0("QQ_Plot_", plot_param, ".jpeg")
    
    jpeg(residual_plot_filename)
    plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
    dev.off()
    
    jpeg(qq_plot_filename)
    qqnorm(residuals(lmer_model), main = "Q-Q Plot")
    qqline(residuals(lmer_model))
    dev.off()
  }
  
  # Return results
  return(list(fixed = lmm_results, random_variances = random_effects_variances, vif = vif_values))
}

vars_model_school <- c("School_Job_Satisfaction",
                 "School_Avg_director_leadership")

# MATH MODEL - SCHOOL LEVEL
full_model_formula_math_school <- paste("School_Avg_PisaMath ~", paste(vars_model_school, collapse = " + "), "+ (1|NAME_CNT)")
df_lmm_results_math_school <- fit_lmm_and_format(full_model_formula_math_school, df_school_stats, save_plots = FALSE)
df_lmm_results_math_school_fixed <- df_lmm_results_math_school$fixed


# SCIENCE MODEL - SCHOOL LEVEL
full_model_formula_science_school <- paste("School_Avg_PisaScience ~", paste(vars_model_school, collapse = " + "), "+ (1|NAME_CNT)")
df_lmm_results_science_school <- fit_lmm_and_format(full_model_formula_science_school, df_school_stats, save_plots = FALSE)
df_lmm_results_science_school_fixed <- df_lmm_results_science_school$fixed


# READING MODEL - SCHOOL LEVEL
full_model_formula_reading_school <- paste("School_Avg_PisaReading ~", paste(vars_model_school, collapse = " + "), "+ (1|NAME_CNT)")
df_lmm_results_reading_school <- fit_lmm_and_format(full_model_formula_reading_school, df_school_stats, save_plots = FALSE)
df_lmm_results_reading_school_fixed <- df_lmm_results_reading_school$fixed


# MATH MODELS
# Join all variables to create the formula for the full model
full_model_formula_math2 <- paste("PisaMath ~", paste(vars_model1, collapse = " + "), "+ (1|NAME_CNT) + (1|ID_SCHOOL)")
full_model_formula_math3 <- paste("PisaMath ~", paste(vars_model2_nocollinear, collapse = " + "), "+ (1|NAME_CNT) + (1|ID_SCHOOL)")

# Fit the null model
df_lmm_results_math_1 <- fit_lmm_and_format("PisaMath ~ (1|NAME_CNT) + (1|ID_SCHOOL)", df, save_plots = FALSE)
df_lmm_results_math_1_fixed <- df_lmm_results_math_1$fixed
df_lmm_results_math_1_random <- df_lmm_results_math_1$random_variances

# Fit the full model
df_lmm_results_math_2 <- fit_lmm_and_format(full_model_formula_math2, df, save_plots = FALSE)
df_lmm_results_math_2_fixed <- df_lmm_results_math_2$fixed
df_lmm_results_math_2_random <- df_lmm_results_math_2$random_variances

# Fit the full model
df_lmm_results_math_3 <- fit_lmm_and_format(full_model_formula_math3, df_updated, save_plots = FALSE)
df_lmm_results_math_3_fixed <- df_lmm_results_math_3$fixed
df_lmm_results_math_3_random <- df_lmm_results_math_3$random_variances

# Eliminate collinearity
full_model_formula_math3 <- paste("PisaMath ~", paste(vars_model2_nocollinear, collapse = " + "), "+ (1|NAME_CNT) + (1|ID_SCHOOL)")
df_lmm_results_math_3 <- fit_lmm_and_format(full_model_formula_math3, df_updated, save_plots = TRUE, plot_param = "Math")
df_lmm_results_math_3_fixed <- df_lmm_results_math_3$fixed
df_lmm_results_math_3_random <- df_lmm_results_math_3$random_variances

# SCIENCE MODELS
# Join all variables to create the formula for the full model
full_model_formula_science2 <- paste("PisaScience ~", paste(vars_model1, collapse = " + "), "+ (1|NAME_CNT) + (1|ID_SCHOOL)")
full_model_formula_science3 <- paste("PisaScience ~", paste(vars_model2_nocollinear, collapse = " + "), "+ (1|NAME_CNT) + (1|ID_SCHOOL)")

# Fit the null model
df_lmm_results_science_1 <- fit_lmm_and_format("PisaMath ~ (1|NAME_CNT) + (1|ID_SCHOOL)", df, save_plots = FALSE)
df_lmm_results_science_1_fixed <- df_lmm_results_science_1$fixed
df_lmm_results_science_1_random <- df_lmm_results_science_1$random_variances

# Fit the full model
df_lmm_results_science_2 <- fit_lmm_and_format(full_model_formula_science2, df, save_plots = FALSE)
df_lmm_results_science_2_fixed <- df_lmm_results_science_2$fixed
df_lmm_results_science_2_random <- df_lmm_results_science_2$random_variances

# Fit the full model
df_lmm_results_science_3 <- fit_lmm_and_format(full_model_formula_science3, df_updated, save_plots = TRUE, plot_param = "Science")
df_lmm_results_science_3_fixed <- df_lmm_results_science_3$fixed
df_lmm_results_science_3_random <- df_lmm_results_science_3$random_variances

# READING MODELS
# Join all variables to create the formula for the full model
full_model_formula_reading2 <- paste("PisaReading ~", paste(vars_model1, collapse = " + "), "+ (1|NAME_CNT) + (1|ID_SCHOOL)")
full_model_formula_reading3 <- paste("PisaReading ~", paste(vars_model2_nocollinear, collapse = " + "), "+ (1|NAME_CNT) + (1|ID_SCHOOL)")

# Fit the null model
df_lmm_results_reading_1 <- fit_lmm_and_format("PisaReading ~ (1|NAME_CNT) + (1|ID_SCHOOL)", df, save_plots = FALSE)
df_lmm_results_reading_1_fixed <- df_lmm_results_reading_1$fixed
df_lmm_results_reading_1_random <- df_lmm_results_reading_1$random_variances

# Fit the full model
df_lmm_results_reading_2 <- fit_lmm_and_format(full_model_formula_reading2, df, save_plots = FALSE)
df_lmm_results_reading_2_fixed <- df_lmm_results_reading_2$fixed
df_lmm_results_reading_2_random <- df_lmm_results_reading_2$random_variances

# Fit the full model
df_lmm_results_reading_3 <- fit_lmm_and_format(full_model_formula_reading3, df_updated, save_plots = TRUE, plot_param = "Reading")
df_lmm_results_reading_3_fixed <- df_lmm_results_reading_3$fixed
df_lmm_results_reading_3_random <- df_lmm_results_reading_3$random_variances


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
  "Country-Level Statistics" = df_final_descriptives,
  "Correlation Matrix" = correlation_matrix, 
  "Model Results - Math 2" = df_lmm_results_math_2_fixed,
  "Model Results - Math 3" = df_lmm_results_math_3_fixed,
  "Model Results - Science 2" = df_lmm_results_science_2_fixed,
  "Model Results - Science 3" = df_lmm_results_science_3_fixed,
  "Model Results - Reading 2" = df_lmm_results_reading_2_fixed,
  "Model Results - Reading 3" = df_lmm_results_reading_3_fixed
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables_v2.xlsx")





# Columns where NAs are allowed
na_allowed_cols <- c("N_BREAK", "Q22A", "Q22B", "Q22C", "Q22D", "Q22E", "Q22F", "Q22G", "Q22H", "Q22I", "Q22J", "Q22K")

# All column names in the dataset
all_cols <- names(df)

# Columns where NAs should lead to row removal (all columns except for na_allowed_cols)
na_remove_cols <- setdiff(all_cols, na_allowed_cols)

# Remove rows with NAs in columns not listed in na_allowed_cols
df_clean <- df[!rowSums(is.na(df[na_remove_cols])), ]

# Alternative Models

library(mgcv)

# Function to fit GAMs and return formatted results without plotting
fit_gam_and_format <- function(formula, data, family = gaussian(), method = "ML") {
  # Attempt to fit the GAM model
  tryCatch({
    gam_model <- gam(formula, data = data, family = family, method = method)
    
    # If model fits successfully, proceed with summary and checks
    summary_result <- summary(gam_model)
    print(summary_result)  # Print to ensure this works
    
    gam_check_result <- tryCatch({
      gam.check(gam_model)
    }, error = function(e) {
      cat("Error in gam.check: ", e$message, "\n")
    })
    
    # Return model summary and diagnostic check results without plotting
    return(list(summary = summary_result, check = gam_check_result))
  }, error = function(e) {
    cat("Error in fitting GAM model: ", e$message, "\n")
    return(NULL)
  })
}


# Example application of the function for Math, corrected for GAM syntax
gam_results_math <- fit_gam_and_format(
  formula = PisaMath ~ s(JobSatisfaction_mean_mean) + 
    s(director_leadership) + 
    s(AVGESCS_mean_mean) + 
    s(NAME_CNT, bs="re") + 
    s(ID_SCHOOL, bs="re"), 
  data = df_clean, method = "ML"
)

# Example for Science
gam_results_science <- fit_gam_and_format(formula = PisaScience ~ s(JobSatisfaction_mean_mean) + s(director_leadership) + s(AVGESCS_mean_mean) + (1|NAME_CNT) + (1|ID_SCHOOL), data = df, discrete = TRUE, method = "ML")

# Example for Reading
gam_results_reading <- fit_gam_and_format(formula = PisaReading ~ s(JobSatisfaction_mean_mean) + s(director_leadership) + s(AVGESCS_mean_mean) + (1|NAME_CNT) + (1|ID_SCHOOL), data = df_updated)
