setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/niloopej")

library(openxlsx)
df <- read.xlsx("CleanData.xlsx")
df_attributes <- read.xlsx("CleanData.xlsx", sheet = 'Attributes')

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

scales_dem_cat <- c("Gender"    ,      "Canada_Region" ,  "Age_Cat", "Education",
                    "Employment"     ,
                    "Location" ,       "Income"        ,  "HH_size"   ,      "Num_Child")

scales_freq <- c("Have.you.heard.of.the.slick_coat.trait.in.the.gene_edited.cattle.before"         ,                                                                     
                 "What.factors.would.influence.your.willingness.to.pay.for.the.slick_coat.trait.in.the.gene_edited.meat.products._.Reducing.methane.emissions"   ,             
                 "What.factors.would.influence.your.willingness.to.pay.for.the.slick_coat.trait.in.the.gene_edited.meat.products._.Human.health.benefits"         ,            
                 "What.factors.would.influence.your.willingness.to.pay.for.the.slick_coat.trait.in.the.gene_edited.meat.products._.Improving.animal.welfare"       ,           
                 "What.factors.would.influence.your.willingness.to.pay.for.the.slick_coat.trait.in.the.gene_edited.meat.products._.I.prefer.not.to.answer"          ,          
                 "How.concerned.are.you.about.the.heat.stress.on.cattle.due.to.climate.change"  ,                                                                             
                 "What.do.you.think.Animal.Welfare.means._.Suffering"                             ,                                                                            
                 "What.do.you.think.Animal.Welfare.means._.Emotions"                               ,                                                                           
                 "What.do.you.think.Animal.Welfare.means._.Happiness"                               ,                                                                          
                 "What.do.you.think.Animal.Welfare.means._.Stress"                                   ,                                                                         
                 "What.do.you.think.Animal.Welfare.means._.Natural_outdoor.conditions"                ,                                                                        
                 "What.do.you.think.Animal.Welfare.means._.Housing_clean.environment_healthy"          ,                                                                       
                 "What.do.you.think.Animal.Welfare.means._.Behaviour"                                   ,                                                                      
                 "What.do.you.think.Animal.Welfare.means._.Health_medical.treatments"                    ,                                                                     
                 "What.do.you.think.Animal.Welfare.means._.Feeding_concentrate"                           ,                                                                    
                 "What.do.you.think.Animal.Welfare.means._.I.prefer.not.to.answer",
                 "From.which.sources.do.you.normally.receive.information.related.to.animal.welfare._.Classic.sources._radio_.TV_.magazines_"     ,                             
                 "From.which.sources.do.you.normally.receive.information.related.to.animal.welfare._.Family_friends"                              ,                            
                 "From.which.sources.do.you.normally.receive.information.related.to.animal.welfare._.Scientific.sources._conferences_professional_publication_" ,              
                 "From.which.sources.do.you.normally.receive.information.related.to.animal.welfare._.Company._food.company.website_food.label_" ,                              
                 "From.which.sources.do.you.normally.receive.information.related.to.animal.welfare._.Government.institution.websites"          ,                               
                 "From.which.sources.do.you.normally.receive.information.related.to.animal.welfare._.I.don’t.usually.receive.information.related.to.animal.welfare"   ,        
                 "From.which.sources.do.you.normally.receive.information.related.to.animal.welfare._.I.prefer.not.to.answer",
                 "Please.indicate.three.aspects.related.to.animal.welfare.in.which.you.think.the.legislation.of.Canada.would.need.to.be.more.strict._if.any_._.None._.nothing")

scales_choice <- c("Block_selection",
                   "CE_Steak1_1_1"   ,                                                                                                                                           
                   "CE_Steak1_1_2"    ,                                                                                                                                          
                   "CE_Steak1_1_3"     ,                                                                                                                                         
                   "CE_Steak1_1_4"      ,                                                                                                                                        
                   "CE_Steak1_1_5"       ,                                                                                                                                       
                   "CE_Steak1_1_6"        ,                                                                                                                                      
                   "CE_Steak1_1_7"          ,                                                                                                                                    
                   "CE_Steak1_1_8"         ,                                                                                                                                  
                   "CE_Steak1_2_1"           ,                                                                                                                                   
                   "CE_Steak1_2_2"           ,                                                                                                                                   
                   "CE_Steak1_2_3"            ,                                                                                                                                  
                   "CE_Steak1_2_4"             ,                                                                                                                                 
                   "CE_Steak1_2_5"            ,                                                                                                                                  
                   "CE_Steak1_2_6"             ,                                                                                                                                 
                   "CE_Steak1_2_7"              ,                                                                                                                                
                   "CE_Steak1_2_8"               ,                                                                                                                               
                   "CE_Steak1_3_1"                ,                                                                                                                              
                   "CE_Steak1_3_2"                 ,                                                                                                                             
                   "CE_Steak1_3_3"                  ,                                                                                                                            
                   "CE_Steak1_3_4"                   ,                                                                                                                           
                   "CE_Steak1_3_5"                    ,                                                                                                                          
                   "CE_Steak1_3_6"                     ,                                                                                                                         
                   "CE_Steak1_3_7"                      ,                                                                                                                        
                   "CE_Steak1_3_8")

scales_desc <- c("In.your.opinion_.how.well.informed.are.you.about.animal.welfare. "              ,                                                                            
                 "In.general.terms_.you.think.that.the.level.of.animal.welfare.in.Canada.is"       ,                                                                           
                 "You.think.that.the.information.you.receive.about.animal.welfare.is",
                 "Is.product.labeling.important to.you.to.buy.gene_edited.products"    ,                                                                                       
                 "In.general.terms_.how.much.informed.do.you.think.the.citizens.of.your.country.are.about.animal.welfare.issues")



colnames(df)

# Frequency Tables

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


df_freq_sociod <- create_frequency_tables(df, scales_dem_cat)
df_freq <- create_frequency_tables(df, scales_freq)
df_freq_choice <- create_frequency_tables(df, scales_choice)

# Transform data

library(tidyr)
library(dplyr)

# Convert df to long format and ensure Block_Selection is properly aligned with df_attributes
df_long <- df %>%
  pivot_longer(cols = starts_with("CE_Steak"), names_to = "Data.column", values_to = "Choice") %>%
  mutate(Block_Selection = as.character(Block_selection)) %>%
  filter(!is.na(Choice))

# Generate a key for joining that incorporates both the choice and block
df_long <- df_long %>%
  mutate(Choice_Block_Key = paste(Choice, Data.column, Block_selection, sep = "_"))

df_attributes <- df_attributes %>%
  mutate(Choice_Block_Key = paste(Choice, Data.column, Block_selection, sep = "_"))

# Now perform the join
df_joined <- left_join(df_long, df_attributes, by = "Choice_Block_Key")

# Generate a key in df_joined that includes RMSID for a direct match with the choice made
df_joined_chosen <- df_joined %>%
  mutate(Chosen = 1) %>%
  select(RMSID, Choice_Block_Key, Chosen)

# First, ensure you have a unique list of all possible choices per block from df_attributes
choices_per_block <- df_attributes %>%
  select(Choice, Data.column, Block_selection) %>%
  distinct()

# Then, create a template dataframe for all respondents, ensuring it includes all possible choice scenarios per block
df_all_possible_choices <- df_long %>%
  select(RMSID, Block_selection) %>%
  distinct() %>%
  group_by(RMSID, Block_selection) %>%
  left_join(choices_per_block, by = "Block_selection") %>%
  ungroup() %>%
  mutate(Choice_Block_Key = paste(Choice, Data.column, Block_selection, sep = "_"))

df_final_with_attributes <- left_join(df_all_possible_choices, df_attributes, by = "Choice_Block_Key")

df_final <- left_join(df_final_with_attributes, df_joined_chosen, by = c("RMSID", "Choice_Block_Key"))

# Fill NAs in 'Chosen' column with 0 to indicate these were not the selected options
df_final <- df_final %>%
  mutate(Chosen = ifelse(is.na(Chosen), 0, Chosen))


df_final <- df_final %>%
  # Remove duplicate rows, if any
  distinct() %>%
  # Select and rename columns to remove '.x' and '.y', keeping the more relevant versions
  select(
    RMSID,
    Block_selection = Block_selection.x,
    Choice = Choice.x,
    Data_column = Data.column.x,
    Animal_Welfare,
    Hair,
    Breeding,
    Price,
    Chosen
  )

df_sociodemographics <- df %>%
  select(RMSID, Gender, Canada_Region, Age_Cat, Education, Employment, Location, Income, HH_size, Num_Child)

# Left join df_final_cleaned with df_sociodemographics to attach sociodemographic information
df_final_with_sociodemographics <- left_join(df_final, df_sociodemographics, by = "RMSID")

# Modelling

library(mlogit)

# Create a unique identifier for each choice situation per respondent
df_final_with_sociodemographics$choice_id <- interaction(df_final_with_sociodemographics$RMSID, 
                                                         df_final_with_sociodemographics$Data_column)

# Convert data types 
df_final_with_sociodemographics$Chosen <- as.logical(df_final_with_sociodemographics$Chosen)


df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(across(where(is.character) & !all_of(c("Data_column")), as.factor))

# Handling attributes when 'None' is chosen
df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(across(c(Animal_Welfare, Hair, Breeding), 
                ~if_else(df_final_with_sociodemographics$Choice == "None of these products", "None", .),
                .names = "modified_{.col}"))

# Converting modified attributes to factors
df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(across(starts_with("modified_"), as.factor))

# Adding a dummy for 'None'
df_final_with_sociodemographics$None_Dummy <- if_else(df_final_with_sociodemographics$Choice == "None of these products", 1, 0)


# MODEL w/out Random Parameters

library(mlogit)

df_final_with_sociodemographics <- droplevels(df_final_with_sociodemographics)

df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(across(c(Animal_Welfare, Hair, Breeding), ~if_else(Choice == "None of these products", "None", .), .names = "modified_{.col}")) %>%
  mutate(across(starts_with("modified_"), as.factor))

df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(across(starts_with("modified_"), as.factor)) %>%  # Ensure factor levels for categorical variables
  mutate(Price = as.numeric(as.character(Price)))  # Convert Price to numeric ensuring there are no NAs due to improper types

df_final_with_sociodemographics$None_Dummy <- as.factor(df_final_with_sociodemographics$None_Dummy)

df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(
    Choice = relevel(Choice, ref = "None of these products"),
    modified_Animal_Welfare = relevel(modified_Animal_Welfare, ref = "None"),
    modified_Hair = relevel(modified_Hair, ref = "None"),
    modified_Breeding = relevel(modified_Breeding, ref = "None")
  )

str(df_final_with_sociodemographics)

# Check for NAs in the original data
sum(is.na(df_final_with_sociodemographics))
# Check for NAs in the mlogit_data
sum(is.na(mlogit_data))

# Preparing the dataset for mlogit
mlogit_data <- mlogit.data(df_final_with_sociodemographics, 
                           choice = "Chosen", 
                           shape = "long", 
                           chid.var = "choice_id", 
                           alt.var = "Choice",
                           drop.index = TRUE,
                           varying = NULL)
str(mlogit_data)

# Define the function for the mlogit analysis
generate_mlogit_analysis <- function(formula, data) {
  # Fit the model
  cl_model <- mlogit(formula, data = data)
  
  # Generate summary
  summary_cl_model <- summary(cl_model)
  
  # Coefficients table
  coef_table <- summary_cl_model$CoefTable
  tidy_output <- as.data.frame(coef_table)
  tidy_output$Variable <- rownames(coef_table)
  df_clmodel1 <- tidy_output[, c("Variable", "Estimate", "Std. Error", "z-value", "Pr(>|z|)")]
  
  # WTP calculations
  beta_price <- df_clmodel1$Estimate[df_clmodel1$Variable == "Price"]
  df_clmodel1$WTP <- NA
  relevant_variables <- names(df_clmodel1)[grepl("modified", names(df_clmodel1))]
  
  for(var in relevant_variables) {
    if(var %in% df_clmodel1$Variable) {
      df_clmodel1$WTP[df_clmodel1$Variable == var] <- -df_clmodel1$Estimate[df_clmodel1$Variable == var] / beta_price
    }
  }
  
  # Model fit statistics
  null_model <- mlogit(Chosen ~ 1, data = data)
  logLik_full <- logLik(cl_model)
  logLik_null <- logLik(null_model)
  lrt_statistic <- 2 * (logLik_full - logLik_null)
  df_lrt <- attr(logLik_full, "df") - attr(logLik_null, "df")
  lrt_p_value <- pchisq(lrt_statistic, df_lrt, lower.tail = FALSE)
  mcfadden_r_squared <- 1 - (logLik_full / logLik_null)
  
  df_model_fit <- data.frame(
    Statistic = c("Chi-Square", "P-Value", "McFadden's R-squared", "Log-Lik"),
    Value = c(lrt_statistic, lrt_p_value, mcfadden_r_squared, logLik_full)
  )
  
  # Return both dataframes
  list(Coefficients_WTP = df_clmodel1, Model_Fit_Stats = df_model_fit)
}

# Running the model
results <- generate_mlogit_analysis(Chosen ~ 1 + modified_Animal_Welfare + modified_Hair + modified_Breeding + Price, mlogit_data)


# Model with Interactions

model_formula <- as.formula("Chosen ~ 1 + Animal_Welfare + Hair + Breeding + Price + 
                             Gender:Price + 
                             Gender:Hair +  
                             Gender:Breeding +  
                             Gender:Animal_Welfare +
                             Canada_Region:Price + 
                             Canada_Region:Hair +  
                             Canada_Region:Breeding +  
                             Canada_Region:Animal_Welfare +
                             Age_Cat:Price + 
                             Age_Cat:Hair +  
                             Age_Cat:Breeding +  
                             Age_Cat:Animal_Welfare +
                             Education:Price + 
                             Education:Hair +  
                             Education:Breeding +  
                             Education:Animal_Welfare +
                             Employment:Price + 
                             Employment:Hair +  
                             Employment:Breeding +  
                             Employment:Animal_Welfare +
                             Location:Price + 
                             Location:Hair +  
                             Location:Breeding +  
                             Location:Animal_Welfare +
                             Income:Price + 
                             Income:Hair +  
                             Income:Breeding +  
                             Income:Animal_Welfare +
                             HH_size:Price + 
                             HH_size:Hair +  
                             HH_size:Breeding +  
                             HH_size:Animal_Welfare +
                             Num_Child:Price + 
                             Num_Child:Hair +  
                             Num_Child:Breeding +  
                             Num_Child:Animal_Welfare| 0")

results_socio <- generate_mlogit_analysis(model_formula, mlogit_data)
df_clmodel2_coefs <- results_socio$Coefficients_WTP
df_clmodel2_fit <- results_socio$Model_Fit_Stats








# Random Model

library(gmnl)

# Create dummy variables
df <- df_final_with_sociodemographics

# Generate dummy variables for specified columns
df_dummies <- model.matrix(~ modified_Animal_Welfare + modified_Hair + modified_Breeding - 1, data = df)

# Convert to a data frame and rename columns to remove spaces
df_dummies <- as.data.frame(df_dummies)
colnames(df_dummies) <- make.names(colnames(df_dummies))

# Bind the dummy variables back to the original dataframe
df_final_with_sociodemographics <- bind_cols(df, df_dummies)


df_long <- mlogit.data(df_final_with_sociodemographics, 
                       choice = "Chosen", 
                       shape = "long", 
                       chid.var = "choice_id", 
                       alt.var = "Choice",
                       drop.index = TRUE,
                       id.var = "RMSID")
colnames(df_long)
str(df_long)

mxl_model <- gmnl(
  formula = Chosen ~ 1 + Price + modified_Animal_WelfareEar.Tag + modified_Animal_WelfareHot.Branding + modified_HairSlick + modified_HairNormal + modified_BreedingGE + modified_BreedingGS + modified_BreedingCB| 0,
  data = df_long,
  model = "mixl",
  ranp = c(modified_Animal_WelfareEar.Tag = "n", modified_Animal_WelfareHot.Branding = "n", modified_HairSlick = "n", modified_HairNormal = "n", modified_BreedingGE = "n", modified_BreedingGS = "n", modified_BreedingCB = "n"),
  R = 100,
  haltons = NA,
  panel = TRUE,
  id.var = "RMSID"  # Ensure this is the correct individual identifier in your dataset
)

summary_mxl_model <- summary(mxl_model)
print(summary_mxl_model)
str(summary_mxl_model)

# Extract coefficients and standard errors
coef_table <- summary_mxl_model$CoefTable

df_mxmodel1_coefs <- as.data.frame(coef_table)

# Ensure variable names are added as a column
df_mxmodel1_coefs$Variable <- rownames(coef_table)

# Reorder columns for readability
df_mxmodel1_coefs <- df_mxmodel1_coefs[, c("Variable", "Estimate", "Std. Error", "z-value", "Pr(>|z|)")]

# Assuming price coefficient is fixed and others are random:
beta_price <- coef_table["Price", "Estimate"]

# Initialize WTP column with NA values
df_mxmodel1_coefs$WTP <- NA

# Calculate WTP for attributes relative to price (if applicable)
# Note: Adjust variable names as needed based on your specific model setup
attributes <- c("modified_Animal_WelfareEar.Tag", "modified_Animal_WelfareHot.Branding", "modified_HairSlick", "modified_HairNormal","modified_BreedingGE", "modified_BreedingGS", "modified_BreedingCB")
for(attr in attributes) {
  if(attr %in% rownames(coef_table)) {
    df_mxmodel1_coefs$WTP[df_mxmodel1_coefs$Variable == attr] <- -df_mxmodel1_coefs$Estimate[df_mxmodel1_coefs$Variable == attr] / beta_price
  }
}

# Model fit statistics
logLik_mxmodel1 <- logLik(mxl_model)
n_obs_mxmodel1 <- length(mlogit_data$Chosen) / length(unique(mlogit_data$idx))
df_lrt_mxmodel1 <- df_mxmodel1_coefs %>% filter(!is.na(WTP)) %>% nrow() # degrees of freedom for LRT

# Generate frequency tables for each factor in the model to check for perfect linear combinations
factor_vars <- c("Gender", "Canada_Region", "Age_Cat", "Education", "Employment", "Location", "Income", "HH_size", "Num_Child",
                 "modified_Animal_WelfareEar.Tag", "modified_Animal_WelfareHot.Branding", "modified_HairSlick", 
                 "modified_HairNormal", "modified_BreedingGE", "modified_BreedingGS", "modified_BreedingCB")

# Applying table function to each factor variable
lapply(df_long[factor_vars], table)


# Specify the model with interactions and random parameters
mxl_model2 <- gmnl(
  formula = Chosen ~ 1 + Price +
    modified_Animal_WelfareEar.Tag + modified_Animal_WelfareHot.Branding +
    modified_HairSlick + modified_HairNormal +
    modified_BreedingGE + modified_BreedingGS + modified_BreedingCB +
    Gender:Price + 
    Gender:modified_HairSlick + 
    Gender:modified_HairNormal +
    Gender:modified_BreedingGE + 
    Gender:modified_BreedingGS + 
    Gender:modified_BreedingCB +
    Gender:modified_Animal_WelfareEar.Tag +
    Gender:modified_Animal_WelfareHot.Branding +
    Canada_Region:Price + 
    Canada_Region:modified_HairSlick +  
    Canada_Region:modified_HairNormal + 
    Canada_Region:modified_BreedingGE +  
    Canada_Region:modified_BreedingGS +
    Canada_Region:modified_BreedingCB +
    Canada_Region:modified_Animal_WelfareEar.Tag +
    Canada_Region:modified_Animal_WelfareHot.Branding +
    Age_Cat:Price + 
    Age_Cat:modified_HairSlick +  
    Age_Cat:modified_HairNormal +  
    Age_Cat:modified_BreedingGE +  
    
    Age_Cat:modified_BreedingGS + 
    Age_Cat:modified_BreedingCB +
    Age_Cat:modified_Animal_WelfareEar.Tag +
    Age_Cat:modified_Animal_WelfareHot.Branding +
    Education:Price + 
    Education:modified_HairSlick + 
    Education:modified_HairNormal +
    Education:modified_BreedingGE + 
    Education:modified_BreedingGS +
    Education:modified_BreedingCB +
    Education:modified_Animal_WelfareEar.Tag +
    Education:modified_Animal_WelfareHot.Branding +
    Employment:Price + 
    Employment:modified_HairSlick +  
    Employment:modified_HairNormal +
    Employment:modified_BreedingGE + 
    Employment:modified_BreedingGS +
    Employment:modified_BreedingCB +
    Employment:modified_Animal_WelfareEar.Tag +
    Employment:modified_Animal_WelfareHot.Branding +
    Location:Price + 
    Location:modified_HairSlick + 
    Location:modified_HairNormal +
    Location:modified_BreedingGE +  
    Location:modified_BreedingGS +
    Location:modified_BreedingCB +
    Location:modified_Animal_WelfareEar.Tag +
    Location:modified_Animal_WelfareHot.Branding +
    Income:Price + 
    Income:modified_HairSlick +  
    Income:modified_HairNormal +
    Income:modified_BreedingGE +  
    Income:modified_BreedingGS +
    Income:modified_BreedingCB +
    Income:modified_Animal_WelfareEar.Tag +
    Income:modified_Animal_WelfareHot.Branding +
    HH_size:Price + 
    HH_size:modified_HairSlick +  
    HH_size:modified_BreedingGE +
    HH_size:modified_BreedingGS +
    HH_size:modified_BreedingCB +
    HH_size:modified_Animal_WelfareEar.Tag +
    HH_size:modified_Animal_WelfareHot.Branding +
    Num_Child:Price + 
    Num_Child:modified_HairSlick +  
    Num_Child:modified_HairNormal + 
    Num_Child:modified_BreedingGE + 
    Num_Child:modified_BreedingGS +
    Num_Child:modified_BreedingCB +
    Num_Child:modified_Animal_WelfareEar.Tag +
    Num_Child:modified_Animal_WelfareHot.Branding| 0,
  data = df_long,
  model = "mixl",
  ranp = c(
    modified_Animal_WelfareEar.Tag = "n", 
    modified_Animal_WelfareHot.Branding = "n", 
    modified_HairSlick = "n", 
    modified_HairNormal = "n", 
    modified_BreedingGE = "n", 
    modified_BreedingGS = "n",
    modified_BreedingCB = "n"
  ),
  R = 100,
  panel = TRUE,
  id.var = "RMSID"
)

# Check the summary of the model
summary_mxl_model2 <- summary(mxl_model2)
print(summary_mxl_model2)

# Prepare the output table for inspection
coef_table2 <- summary_mxl_model2$CoefTable
df_mxmodel2_coefs <- as.data.frame(coef_table2)
df_mxmodel2_coefs$Variable <- rownames(coef_table2)
df_mxmodel2_coefs <- df_mxmodel2_coefs[, c("Variable", "Estimate", "Std. Error", "z-value", "Pr(>|z|)")]

# Display the coefficients table
print(df_mxmodel2_coefs)

for(attr in attributes) {
  if(attr %in% rownames(coef_table)) {
    df_mxmodel2_coefs$WTP[df_mxmodel2_coefs$Variable == attr] <- -df_mxmodel2_coefs$Estimate[df_mxmodel2_coefs$Variable == attr] / beta_price
  }
}

print(df_mxmodel2_coefs)