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

df_attributes_filtered <- dplyr::select(df_attributes, Choice_Block_Key, Animal_Welfare, Hair, Breeding)

# Perform the join with only the specified columns from df_attributes
df_joined <- dplyr::left_join(df_long, df_attributes_filtered, by = "Choice_Block_Key")


# Generate a key in df_joined that includes RMSID for a direct match with the choice made
df_joined_chosen <- df_joined %>%
  mutate(Chosen = 1) %>%
  dplyr::select(RMSID, Choice_Block_Key, Chosen)



# First, ensure you have a unique list of all possible choices per block from df_attributes
choices_per_block <- df_attributes %>%
  dplyr::select(Choice, Data.column, Block_selection) %>%
  distinct()

# Then, create a template dataframe for all respondents, ensuring it includes all possible choice scenarios per block
df_all_possible_choices <- df_long %>%
  dplyr::select(RMSID, Block_selection) %>%
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
  dplyr::select(
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
  dplyr::select(RMSID, Gender, Canada_Region, Age_Cat, Education, Employment, Location, Income, HH_size, Num_Child)

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


# Random Model

library(gmnl)

# Create dummy variables
# Prepare the data frame
df <- df_final_with_sociodemographics

# Convert factor variables to character to prevent automatic dropping of levels
df$modified_Animal_Welfare <- as.character(df$modified_Animal_Welfare)
df$modified_Hair <- as.character(df$modified_Hair)
df$modified_Breeding <- as.character(df$modified_Breeding)

# Reconvert characters to factors specifying that all levels should be included
df$modified_Animal_Welfare <- factor(df$modified_Animal_Welfare, levels = unique(df$modified_Animal_Welfare))
df$modified_Hair <- factor(df$modified_Hair, levels = unique(df$modified_Hair))
df$modified_Breeding <- factor(df$modified_Breeding, levels = unique(df$modified_Breeding))

# Define the contrasts to keep all levels for each factor
contrasts <- list(
  modified_Animal_Welfare = contrasts(df$modified_Animal_Welfare, contrasts = FALSE),
  modified_Hair = contrasts(df$modified_Hair, contrasts = FALSE),
  modified_Breeding = contrasts(df$modified_Breeding, contrasts = FALSE)
)

# Create the dummy variables, including all categories
df_dummies <- model.matrix(~ modified_Animal_Welfare + modified_Hair + modified_Breeding + 0, 
                           data = df, 
                           contrasts.arg = contrasts)

# Add a dummy variable for NA values manually
df_dummies <- as.data.frame(df_dummies)


colnames(df_dummies) <- make.names(colnames(df_dummies))

# Bind the dummy variables back to the original dataframe
df_final_with_sociodemographics <- bind_cols(df, df_dummies)

# Print the column names to verify all dummies are included
colnames(df_final_with_sociodemographics)


df_long <- mlogit.data(df_final_with_sociodemographics, 
                       choice = "Chosen", 
                       shape = "long", 
                       chid.var = "choice_id", 
                       alt.var = "Choice",
                       drop.index = TRUE,
                       id.var = "RMSID")
colnames(df_long)

mlogit.data <- mlogit.data(df_final_with_sociodemographics, 
                       choice = "Chosen", 
                       shape = "long", 
                       chid.var = "choice_id", 
                       alt.var = "Choice",
                       drop.index = TRUE,
                       id.var = "RMSID")

str(df_long)


mxl_model <- gmnl(
  formula = Chosen ~ 1 + log_price + modified_Animal_WelfareHot.Branding + modified_Animal_WelfareEar.Tag + modified_HairNormal + modified_HairSlick + modified_BreedingGE + modified_BreedingCB + modified_BreedingGS | 0,
  data = df_long,
  model = "mixl",
  ranp = c(modified_Animal_WelfareHot.Branding = "n",  modified_Animal_WelfareEar.Tag = "n",modified_HairNormal = "n",  modified_HairSlick = "n",  modified_BreedingGE = "n", modified_BreedingCB = "n", modified_BreedingGS = "n"),
  R = 500,
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
beta_price <- coef_table["log_price", "Estimate"]

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
n_obs_mxmodel1 <- length(df_long$Chosen) / length(unique(df_long$idx))
df_lrt_mxmodel1 <- df_mxmodel1_coefs %>% filter(!is.na(WTP)) %>% nrow() # degrees of freedom for LRT


mxl_model_simpl <- gmnl(
  formula = Chosen ~ 1 + log_price + modified_Animal_WelfareHot.Branding +  modified_HairNormal + modified_BreedingGE + modified_BreedingCB | 0,
  data = df_long,
  model = "mixl",
  ranp = c(modified_Animal_WelfareHot.Branding = "n",  modified_HairNormal = "n",   modified_BreedingGE = "n", modified_BreedingCB = "n"),
  R = 500,
  haltons = NA,
  panel = TRUE,
  id.var = "RMSID"  # Ensure this is the correct individual identifier in your dataset
)


summary_mxl_model_simpl <- summary(mxl_model_simpl)
print(summary_mxl_model_simpl)
str(summary_mxl_model_simpl)

# Extract coefficients and standard errors
coef_table_simpl <- summary_mxl_model_simpl$CoefTable

df_mxmodel1_coefs_simpl <- as.data.frame(coef_table_simpl)

# Ensure variable names are added as a column
df_mxmodel1_coefs_simpl$Variable <- rownames(coef_table_simpl)

# Reorder columns for readability
df_mxmodel1_coefs_simpl <- df_mxmodel1_coefs_simpl[, c("Variable", "Estimate", "Std. Error", "z-value", "Pr(>|z|)")]

# Assuming price coefficient is fixed and others are random:
beta_price <- coef_table_simpl["log_price", "Estimate"]

# Initialize WTP column with NA values
df_mxmodel1_coefs_simpl$WTP <- NA

# Calculate WTP for attributes relative to price (if applicable)
# Note: Adjust variable names as needed based on your specific model setup
attributes <- c("modified_Animal_WelfareEar.Tag", "modified_HairSlick", "modified_BreedingGE", "modified_BreedingGS")
for(attr in attributes) {
  if(attr %in% rownames(coef_table_simpl)) {
    df_mxmodel1_coefs_simpl$WTP[df_mxmodel1_coefs_simpl$Variable == attr] <- -df_mxmodel1_coefs_simpl$Estimate[df_mxmodel1_coefs$Variable == attr] / beta_price
  }
}

# Model fit statistics
logLik_mxmodel1_simpl <- logLik(mxl_model_simpl)
n_obs_mxmodel1_simpl <- length(df_long$Chosen) / length(unique(df_long$idx))
df_lrt_mxmodel1_simpl <- df_mxmodel1_coefs_simpl %>% filter(!is.na(WTP)) %>% nrow() # degrees of freedom for LRT







# Load necessary library
library(dplyr)

# Load necessary libraries
library(dplyr)
library(forcats)  # For handling factors more effectively

# Update the dataset by explicitly managing factor levels
df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(
    Employment = fct_collapse(Employment,
                              "Other Employment or Unemployed" = c("Home-maker", "Student", "I prefer not to answer", "Unemployed", "Part-time employment")),
    Education = fct_collapse(Education,
                             "Primary or Less" = c("Did not attend school", "Primary education")),
    Location = fct_collapse(Location,
                            "Rural or not answered" = c("I prefer not to answer", "Rural")),
    Income = fct_collapse(Income,
                          "Less than $25,000 or not answered" = c("I prefer not to answer", "Less than $25,000")),
    Num_Child = fct_collapse(Num_Child,
                             "3 or more" = c("More than 3", "3")),
    HH_size = fct_collapse(HH_size,
                           "More than 4 or not answered" = c("I prefer not to answer", "More than 4"))
  )

df_final_with_sociodemographics <- df_final_with_sociodemographics %>%
  mutate(Canada_Region = fct_collapse(Canada_Region,
                                      `Atlantic & Northern Canada` = c("Atlantic Canada", "Northern Canada")))

# Calculate the count of each distinct value for each column
value_counts <- lapply(df_final_with_sociodemographics, table)

# Print the counts for each column
value_counts


df_long <- mlogit.data(df_final_with_sociodemographics, 
                       choice = "Chosen", 
                       shape = "long", 
                       chid.var = "choice_id", 
                       alt.var = "Choice",
                       drop.index = TRUE,
                       id.var = "RMSID")


# Specify the model with interactions and random parameters
mxl_model2 <- gmnl(
  formula = Chosen ~ 1 + Price +
    modified_Animal_WelfareHot.Branding + modified_Animal_WelfareNone +
    modified_HairNormal + modified_HairNone +
    modified_BreedingGE + modified_BreedingCB + modified_BreedingNone +
    Gender:Price + 
    Gender:modified_HairNormal + 
    Gender:modified_HairNone +
    Gender:modified_BreedingGE + 
    Gender:modified_BreedingCB + 
    Gender:modified_BreedingNone +
    Gender:modified_Animal_WelfareHot.Branding +
    Gender:modified_Animal_WelfareNone,
    
  data = df_long,
  model = "mixl",
  ranp = c(
    modified_Animal_WelfareHot.Branding = "n", 
    modified_Animal_WelfareNone = "n", 
    modified_HairNone = "n", 
    modified_HairNormal = "n", 
    modified_BreedingGE = "n", 
    modified_BreedingCB = "n",
    modified_BreedingNone = "n"
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



library(dplyr)

# Assuming df_long is your dataset loaded with all the necessary variables
distinct_counts <- df_long %>%
  summarise(
    
    Count_modified_Animal_WelfareEar_Tag = n_distinct(modified_Animal_WelfareEar.Tag),
    Count_modified_Animal_WelfareHot_Branding = n_distinct(modified_Animal_WelfareHot.Branding),
    Count_modified_HairSlick = n_distinct(modified_HairSlick),
    Count_modified_HairNormal = n_distinct(modified_HairNormal),
    Count_modified_BreedingGE = n_distinct(modified_BreedingGE),
    Count_modified_BreedingGS = n_distinct(modified_BreedingGS),
    Count_modified_BreedingCB = n_distinct(modified_BreedingCB),
    Count_Gender = n_distinct(Gender),
    Count_Age_Cat = n_distinct(Age_Cat),
    Count_Education = n_distinct(Education)
  )

# Print the counts of distinct factors for each variable
print(distinct_counts)



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
  RML = df_mxmodel1_coefs,
  RML2 = df_mxmodel2_coefs
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "New_Models.xlsx")
