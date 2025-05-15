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

scales_desc <- c("In.your.opinion_.how.well.informed.are.you.about.animal.welfare."              ,                                                                            
                 "In.general.terms_.you.think.that.the.level.of.animal.welfare.in.Canada.is"       ,                                                                           
                 "You.think.that.the.information.you.receive.about.animal.welfare.is",
                 "Is_product_labeling_importanttoyou_to_buy_geneedited_products"         ,                                                                                       
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

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(), # Standard Error of the Mean
    SD = numeric(),  # Standard Deviation
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    
    # Check if the variable data is numeric
    if (is.numeric(variable_data)) {
      mean_val <- mean(variable_data, na.rm = TRUE)
      sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
      sd_val <- sd(variable_data, na.rm = TRUE)
      
      # Append the results for each variable to the results dataframe
      results <- rbind(results, data.frame(
        Variable = var,
        Mean = mean_val,
        SEM = sem_val,
        SD = sd_val
      ))
    } else {
      warning(paste("Variable", var, "is not numeric or logical. Descriptive statistics cannot be calculated."))
    }
  }
  
  return(results)
}




df_descriptive_stats <- calculate_descriptive_stats(df, scales_desc)
str(df)

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
# Drop the None answer
df_final_with_sociodemographics <- df_final_with_sociodemographics %>% 
  filter(Choice != "None of these products")

# Convert data types 
df_final_with_sociodemographics$Chosen <- as.logical(as.numeric(df_final_with_sociodemographics$Chosen) - 1)

df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(across(where(is.character) & !all_of(c("Data_column")), as.factor))

df_final_with_sociodemographics <- na.omit(df_final_with_sociodemographics)

# Exclude cases that chose 'None'

df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  group_by(choice_id) %>%
  filter(any(Chosen == TRUE)) %>%
  ungroup()


# MODEL w/out Random Parameters

library(mlogit)

df_final_with_sociodemographics <- droplevels(df_final_with_sociodemographics)

# Preparing the dataset for mlogit
mlogit_data <- mlogit.data(df_final_with_sociodemographics, 
                           choice = "Chosen", 
                           shape = "long", 
                           chid.var = "choice_id", 
                           alt.var = "Choice",
                           drop.index = TRUE,
                           varying = NULL)

# Define the function
run_conjoint_analysis <- function(data, choice_col, formula, id_var, chid_var, alt_var) {
  
  # Convert data to mlogit format
  mlogit_data <- mlogit.data(data, 
                             choice = choice_col, 
                             shape = "long", 
                             id.var = id_var,
                             chid.var = chid_var,
                             alt.var = alt_var)
  
  # Fit the model with the specified formula
  model <- mlogit(formula, data = mlogit_data)
  
  # Generate summary of the model
  summary_model <- summary(model)
  
  # Extracting coefficients and formatting
  coef_table <- summary_model$CoefTable
  coefficients_df <- as.data.frame(coef_table)
  coefficients_df$Variable <- rownames(coef_table)
  coefficients_df <- coefficients_df[, c("Variable", "Estimate", "Std. Error", "z-value", "Pr(>|z|)")]
  
  # Model fit statistics in a dataframe
  model_fit_df <- data.frame(
    LogLik = as.numeric(logLik(model)), 
    AIC = AIC(model), 
    BIC = BIC(model)
  )
  
  # Return the results
  return(list(Coefficients = coefficients_df, Model_Fit_Statistics = model_fit_df))
}

results <- generate_mlogit_analysis(Chosen ~ 1 + Animal_Welfare + Hair + Breeding + Price | 0, mlogit_data)

df_clmodel1_coefs <- results$Coefficients_WTP
df_clmodel1_fit <- results$Model_Fit_Stats

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
df_dummies <- model.matrix(~ Animal_Welfare + Hair + Breeding - 1, data = df)

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
  formula = Chosen ~ 1 + Price + Animal_WelfareEar.Tag + HairSlick + BreedingGE + BreedingGS| 0,
  data = df_long,
  model = "mixl",
  ranp = c(Animal_WelfareEar.Tag = "n", HairSlick = "n", BreedingGE = "n", BreedingGS = "n"),
  R = 2,
  panel = TRUE 
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
attributes <- c("Animal_WelfareEar.Tag", "HairSlick", "BreedingGE", "BreedingGS")
for(attr in attributes) {
  if(attr %in% rownames(coef_table)) {
    df_mxmodel1_coefs$WTP[df_mxmodel1_coefs$Variable == attr] <- -df_mxmodel1_coefs$Estimate[df_mxmodel1_coefs$Variable == attr] / beta_price
  }
}

# Model fit statistics
logLik_mxmodel1 <- logLik(mxl_model)
n_obs_mxmodel1 <- length(mlogit_data$Chosen) / length(unique(mlogit_data$idx))
df_lrt_mxmodel1 <- df_mxmodel1_coefs %>% filter(!is.na(WTP)) %>% nrow() # degrees of freedom for LRT


mxl_model2 <- gmnl(
  formula = Chosen ~ 1 + Price + Animal_WelfareEar.Tag + HairSlick + BreedingGE + BreedingGS +
    Gender:Price + 
    Gender:HairSlick +  
    Gender:BreedingGE + 
    Gender:BreedingGS +  
    Gender:Animal_WelfareEar.Tag +
    Canada_Region:Price + 
    Canada_Region:HairSlick +  
    Canada_Region:BreedingGE +  
    Canada_Region:BreedingGS +
    Canada_Region:Animal_WelfareEar.Tag +
    Age_Cat:Price + 
    Age_Cat:HairSlick +  
    Age_Cat:BreedingGE +  
    Age_Cat:BreedingGS + 
    Age_Cat:Animal_WelfareEar.Tag +
    Education:Price + 
    Education:HairSlick +  
    Education:BreedingGE + 
    Education:BreedingGS +
    Education:Animal_WelfareEar.Tag +
    Employment:Price + 
    Employment:HairSlick +  
    Employment:BreedingGE + 
    Employment:BreedingGS +
    Employment:Animal_WelfareEar.Tag +
    Location:Price + 
    Location:HairSlick +  
    Location:BreedingGE +  
    Location:BreedingGS +
    Location:Animal_WelfareEar.Tag +
    Income:Price + 
    Income:HairSlick +  
    Income:BreedingGE +  
    Income:BreedingGS + 
    Income:Animal_WelfareEar.Tag +
    HH_size:Price + 
    HH_size:HairSlick +  
    HH_size:BreedingGE +
    HH_size:BreedingGS +
    
    HH_size:Animal_WelfareEar.Tag +
    Num_Child:Price + 
    Num_Child:HairSlick +  
    Num_Child:BreedingGE + 
    Num_Child:BreedingGS +
    Num_Child:Animal_WelfareEar.Tag | 0,
  data = df_long,
  model = "mixl",
  ranp = c(Animal_WelfareEar.Tag = "n", HairSlick = "n", BreedingGE = "n", BreedingGS = "n"),
  R = 2,
  panel = TRUE
)

summary_mxl_model2 <- summary(mxl_model2)
print(summary_mxl_model2)

coef_table2 <- summary_mxl_model2$CoefTable

df_mxmodel2_coefs <- as.data.frame(coef_table2)

# Ensure variable names are added as a column
df_mxmodel2_coefs$Variable <- rownames(coef_table2)

# Reorder columns for readability
df_mxmodel2_coefs <- df_mxmodel2_coefs[, c("Variable", "Estimate", "Std. Error", "z-value", "Pr(>|z|)")]

# Assuming price coefficient is fixed and others are random:
beta_price2 <- coef_table2["Price", "Estimate"]

# Initialize WTP column with NA values
df_mxmodel2_coefs$WTP <- NA

# Calculate WTP for attributes relative to price (if applicable)
# Note: Adjust variable names as needed based on your specific model setup
attributes <- c("Animal_WelfareEar.Tag", "HairSlick", "BreedingGE", "BreedingGS")
for(attr in attributes) {
  if(attr %in% rownames(coef_table2)) {
    df_mxmodel2_coefs$WTP[df_mxmodel2_coefs$Variable == attr] <- -df_mxmodel2_coefs$Estimate[df_mxmodel2_coefs$Variable == attr] / beta_price2
  }
}

# Model fit statistics
logLik_mxmodel2 <- logLik(mxl_model2)
n_obs_mxmodel2 <- length(mlogit_data$Chosen) / length(unique(mlogit_data$idx))
df_lrt_mxmodel2 <- df_mxmodel2_coefs %>% filter(!is.na(WTP)) %>% nrow() # degrees of freedom for LRT


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

df_longdf <- as.data.frame(df_long)

# Example usage
data_list <- list(
  "Long-Shaped Data" = df_final_with_attributes,
  "Frequency Table - SocioD" = df_freq_sociod,
  "Frequency Table - Choice" = df_freq_choice, 
  "Frequency Table - Others" = df_freq, 
  "Descriptives" = df_descriptive_stats,
  "Coefs - CL Model" = df_clmodel1_coefs,
  "Model fit - CL Model" = df_clmodel1_fit,
  "Coefs - CL Model Int." = df_clmodel2_coefs,
  "Model fit - CL Model Int." = df_clmodel2_fit,
  "Coefs - RPL Model" = df_mxmodel1_coefs,
  "Coefs - RPL Model Int." = df_mxmodel2_coefs
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

#Save models

# Save model to a file
saveRDS(mxl_model, file = "model_file.rds")
saveRDS(mxl_model2, file = "model_fileint.rds")

# Load model back into R
loaded_model2 <- readRDS("model_fileint.rds")

loaded_model2$logLik
model_summary <- summary(loaded_model)
print(model_summary)
