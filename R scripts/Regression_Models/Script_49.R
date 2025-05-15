# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/pooyapourak")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("The Nonprofit Alliance 2024 Survey Raw Data Analysis.xlsx")

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
str(df)

# Identify likely Likert-type columns
likert_columns <- sapply(df, function(x) any(grepl("4 - agree |3 - neither agree nor disagree|5 - agree strongly|2 - disagree|1 - disagree strongly|not at all important|not too important|somewhat important|very important|1 - not important at all|2 - not very important|3 - somewhat important|4 - important|5 - extremely important|6 - trust|7 - trust completely|neither trust nor distrust|trust somewhat|distrust somewhat|distrust", unique(x))))

# Filter to get names of columns with likely Likert responses
likert_columns_names <- names(df)[likert_columns]

# Revised recoding function excluding "yes" and "no"
recode_likert <- function(x) {
  levels <- c("very important"  ,     "somewhat important" ,  "not too important"  ,  "not at all important",
              
              "1 - not important at all", "2 - not very important", "3 - somewhat important", 
              "4 - important", "5 - extremely important",
              
              
              
              "4 - neither trust nor distrust", "5 - trust somewhat",
              "6 - trust", "2 - distrust", "3 - distrust somewhat", "7 - trust completely",
              "1 - distrust completely",
              "4 - agree" ,                     "3 - neither agree nor disagree", "5 - agree strongly"   ,          "2 - disagree"  ,                
              "1 - disagree strongly" )
  
  # Mapping levels to numeric values
  map <- setNames(c(4,3,2,1,1, 2, 3, 4, 5, 4,5,6,2,3,7,1,4,3,5,2,1), levels)
  factor(x, levels = levels, labels = map[levels], ordered = TRUE)
}

# Convert factor labels to numeric
convert_factor_to_numeric <- function(f) {
  as.numeric(as.character(f))
}

# Copy dataframe to preserve original data
df_likert_recode = df

# Identify likely Likert-type columns again or reuse previously identified columns
likert_columns_names <- names(df)[likert_columns] # Assuming previous identification is reused

# Apply recoding function to each identified Likert column in the copied dataframe
df_likert_recode[likert_columns_names] <- lapply(df_likert_recode[likert_columns_names], function(x) convert_factor_to_numeric(recode_likert(x)))

# Sample inspection of the recoded data structure in the new dataframe
str(df_likert_recode)
colnames(df_likert_recode)



print_unique_values <- function(df) {
  # Using sapply to iterate over each column
  unique_values_list <- sapply(df, function(column) {
    unique_values <- unique(column)  # Get unique values for the column
    # Return a list containing the column name and its unique values
    return(list(UniqueValues = unique_values))
  })
  
  # Print the unique values for each column
  names(unique_values_list) <- colnames(df)  # Set names to column names
  lapply(names(unique_values_list), function(col_name) {
    cat("Unique values in", col_name, ":\n")
    print(unique_values_list[[col_name]])
    cat("\n")  # Add a newline for better readability
  })
}


# Assuming 'df' is your dataframe
print_unique_values(df_likert_recode)


values_to_recode <- c("transgender", "prefer not to answer", "other (specify):")

# Recode these values to 'other'
df_likert_recode$QS2..Gender.Identity[df_likert_recode$QS2..Gender.Identity %in% values_to_recode] <- 'other'



# Study 1
  dv11 <- c(
    "QA2r1..Trust._.Nonprofits"    # Corresponding to trust in nonprofits
    
  )

dv12 <- c(
  
  "QA14..Perceived.Operational.Efficiency"  # Corresponding to perceived efficiency
)

ivs11 <- c(
  "AGEGROUP.._COMPUTED_._.Age.group", 
  "QS2..Gender.Identity",
  "QNB4..Household.Income",
  "USCR.._COMPUTED_._.Region.based.on.US.State",
  "QNB6..Political.Party.Preference",
  "QA10..Awareness.of.Overhead.Costs",
  
  "QA14..Perceived.Operational.Efficiency", # As independent for another context
  "QA9r1..Expenses._.Insurance"                ,                                         
  "QA9r2..Expenses._.Technology"                ,                                        
  "QA9r3..Expenses._.Program.Staff"              ,                                       
  "QA9r4..Expenses._.Fundraising.Staff"           ,                                      
  "QA9r5..Expenses._.Fundraising.Activities"       ,                                     
  "QA9r6..Expenses._.CEO.Salary"                    ,                                    
  "QA9r7..Expenses._.Marketing"                      ,                                   
  "QA9r8..Expenses._.Volunteer.Recruitment"           ,                                  
  "QA9r9..Expenses._.Legal.Services"                   ,                                 
  "QA9r10..Expenses._.Financial.Audit"                  ,                                
  "QA9r11..Expenses._.Advocacy"                          ,                               
  "QA9r12..Expenses._.Board.Meetings", 
  "QA8r1..Importance._.Accreditation"  ,                                                 
  "QA8r2..Importance._.Governance"      ,                                                
  "QA8r3..Importance._.Passion.for.Cause",                                               
  "QA8r4..Importance._.Personal_Family.Use"       ,                                      
  "QA8r5..Importance._.Community.Benefit"          ,                                     
  "QA8r6..Importance._.Effective.Management"        ,                                    
  "QA8r7..Importance._.Achieving.Goals"              ,                                   
  "QA8r8..Importance._.Reputation"                    ,                                  
  "QA8r9..Importance._.Compelling.Stories"             ,                                 
  "QA8r10..Importance._.Leadership.Recognition"         ,                                
  "QA8r11..Importance._.Financial.Transparency"          ,                               
  "QA8r12..Importance._.Community.Impact.Info"            ,                              
  "QA8r13..Importance._.Competitive.Overhead"              ,                             
  "QA8r14..Importance._.Low.Administrative.Costs",
  "QA3r11..Awareness._.Nonprofit.Employment"         ,                                   
  "QA3r12..Awareness._.Volunteer.Participation"       ,                                  
  "QA3r13..Awareness._.Nonprofit.Services"             ,                                 
  "QA3r14..Awareness._.Economic.Contribution"           ,                                
  "QA3r15..Awareness._.Sectors.GDP.Contribution"         ,                               
  "QA3r16..Awareness._.Total.Nonprofits",
                                                    
  "QA2r2..Trust._.Small.Businesses"    ,                                                 
  "QA2r3..Trust._.Large.Corporations"  ,                                                 
  "QA2r4..Trust._.Governments" ,
  "QS8r1..Supported.Organization.Types._.Arts...Culture"                  ,              
  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare"    ,             
  "QS8r3..Supported.Organization.Types._.Health"                            ,            
  "QS8r4..Supported.Organization.Types._.Hospitals"                          ,           
  "QS8r5..Supported.Organization.Types._.Education.and.Research"              ,          
  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief",
  "QS8r7..Supported.Organization.Types._.Religious.Organizations"                       ,
  "QS8r8..Supported.Organization.Types._.Social_Human.Services"                         ,
  "QS8r9..Supported.Organization.Types._.Sports...Recreation"                           ,
  "QS8r10..Supported.Organization.Types._.Universities...Colleges"                      ,
  "QS8r888..Supported.Organization.Types._.Other"
)

ivs12 <- c(
  "AGEGROUP.._COMPUTED_._.Age.group", 
  "QS2..Gender.Identity",
  "QNB4..Household.Income",
  "USCR.._COMPUTED_._.Region.based.on.US.State",
  "QNB6..Political.Party.Preference",
  "QA10..Awareness.of.Overhead.Costs",
  
  
  "QA9r1..Expenses._.Insurance"                ,                                         
  "QA9r2..Expenses._.Technology"                ,                                        
  "QA9r3..Expenses._.Program.Staff"              ,                                       
  "QA9r4..Expenses._.Fundraising.Staff"           ,                                      
  "QA9r5..Expenses._.Fundraising.Activities"       ,                                     
  "QA9r6..Expenses._.CEO.Salary"                    ,                                    
  "QA9r7..Expenses._.Marketing"                      ,                                   
  "QA9r8..Expenses._.Volunteer.Recruitment"           ,                                  
  "QA9r9..Expenses._.Legal.Services"                   ,                                 
  "QA9r10..Expenses._.Financial.Audit"                  ,                                
  "QA9r11..Expenses._.Advocacy"                          ,                               
  "QA9r12..Expenses._.Board.Meetings", 
  "QA8r1..Importance._.Accreditation"  ,                                                 
  "QA8r2..Importance._.Governance"      ,                                                
  "QA8r3..Importance._.Passion.for.Cause",                                               
  "QA8r4..Importance._.Personal_Family.Use"       ,                                      
  "QA8r5..Importance._.Community.Benefit"          ,                                     
  "QA8r6..Importance._.Effective.Management"        ,                                    
  "QA8r7..Importance._.Achieving.Goals"              ,                                   
  "QA8r8..Importance._.Reputation"                    ,                                  
  "QA8r9..Importance._.Compelling.Stories"             ,                                 
  "QA8r10..Importance._.Leadership.Recognition"         ,                                
  "QA8r11..Importance._.Financial.Transparency"          ,                               
  "QA8r12..Importance._.Community.Impact.Info"            ,                              
  "QA8r13..Importance._.Competitive.Overhead"              ,                             
  "QA8r14..Importance._.Low.Administrative.Costs",
  "QA3r11..Awareness._.Nonprofit.Employment"         ,                                   
  "QA3r12..Awareness._.Volunteer.Participation"       ,                                  
  "QA3r13..Awareness._.Nonprofit.Services"             ,                                 
  "QA3r14..Awareness._.Economic.Contribution"           ,                                
  "QA3r15..Awareness._.Sectors.GDP.Contribution"         ,                               
  "QA3r16..Awareness._.Total.Nonprofits",
  "QA2r1..Trust._.Nonprofits",
  "QA2r2..Trust._.Small.Businesses"    ,                                                 
  "QA2r3..Trust._.Large.Corporations"  ,                                                 
  "QA2r4..Trust._.Governments" ,
  "QS8r1..Supported.Organization.Types._.Arts...Culture"                  ,              
  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare"    ,             
  "QS8r3..Supported.Organization.Types._.Health"                            ,            
  "QS8r4..Supported.Organization.Types._.Hospitals"                          ,           
  "QS8r5..Supported.Organization.Types._.Education.and.Research"              ,          
  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief",
  "QS8r7..Supported.Organization.Types._.Religious.Organizations"                       ,
  "QS8r8..Supported.Organization.Types._.Social_Human.Services"                         ,
  "QS8r9..Supported.Organization.Types._.Sports...Recreation"                           ,
  "QS8r10..Supported.Organization.Types._.Universities...Colleges"                      ,
  "QS8r888..Supported.Organization.Types._.Other"
)

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


vars <- c(
  "AGEGROUP.._COMPUTED_._.Age.group", 
  "QS2..Gender.Identity",
  "QNB4..Household.Income",
  "USCR.._COMPUTED_._.Region.based.on.US.State",
  "QNB6..Political.Party.Preference")

df_freq <- create_frequency_tables(df_likert_recode, vars)

df_freq_ivs11 <- create_frequency_tables(df_likert_recode, ivs11)

df_freq_dvs <- create_frequency_tables(df_likert_recode, c(dv11,dv12))

# Convert Gender Identity to factor with specified levels
df_likert_recode$QS2..Gender.Identity <- factor(df_likert_recode$QS2..Gender.Identity, levels = c("male", "female", "other"))

# Convert Age Group to factor with ordered levels
df_likert_recode$AGEGROUP.._COMPUTED_._.Age.group <- factor(df_likert_recode$AGEGROUP.._COMPUTED_._.Age.group, levels = c("under 25", "25-44", "45-64", "65+"))

# Convert US Census Division to factor with specified levels
df_likert_recode$USCD.._COMPUTED_._.Region.based.on.US.State <- factor(df_likert_recode$USCD.._COMPUTED_._.Region.based.on.US.State, levels = c("west south central", "middle atlantic", "pacific", "east north central", "mountain", "east south central", "south atlantic", "west north central", "new england"))

# Convert US Census Region to factor with specified levels
df_likert_recode$USCR.._COMPUTED_._.Region.based.on.US.State <- factor(df_likert_recode$USCR.._COMPUTED_._.Region.based.on.US.State, levels = c("south", "north east", "west", "midwest"))

# Convert Charity Donation Past Year to factor with specified levels
df_likert_recode$QS5..Donated.to.Charity._Past.Year_ <- factor(df_likert_recode$QS5..Donated.to.Charity._Past.Year_, levels = c("no", "yes"))

# Example for categorical responses with "yes", "no", "not sure"
df_likert_recode$QA5r1..Addressing._.Food.Insecurity <- factor(df_likert_recode$QA5r1..Addressing._.Food.Insecurity, levels = c("no", "not sure", "yes"))

# For Employment Status which is more categorical without a specific order
df_likert_recode$QB1..Employment.Status <- factor(df_likert_recode$QB1..Employment.Status, levels = c("retired", "employed full-time", "not employed at present", "employed part-time", "student"))

# Charity Donation Amount
df_likert_recode$QS6..Charity.Donation.Amount._Past.Year_ <- factor(df_likert_recode$QS6..Charity.Donation.Amount._Past.Year_,
                                                      levels = c("less than $500", "$500 to $1000", "$1001 to $2000",
                                                                 "$2001 to $3000", "$3001 to $4000", "$4001 to $5000",
                                                                 "$5000+", "prefer not to answer"))

# Donation Trend
df_likert_recode$QS7..Donation.Trend._Past.Years_ <- factor(df_likert_recode$QS7..Donation.Trend._Past.Years_,
                                              levels = c("down", "same", "up"))

# Impact of Crises on Charitable Services
df_likert_recode$QA4..Impact.of.Crises.on.Charitable.Services <- factor(df_likert_recode$QA4..Impact.of.Crises.on.Charitable.Services,
                                                          levels = c("no", "don't know", "yes - less demand", "yes - more demand"))


response_levels <- c("no", "not sure", "yes")
df_likert_recode$QA5r1..Addressing._.Food.Insecurity <- factor(df_likert_recode$QA5r1..Addressing._.Food.Insecurity, levels = response_levels)
df_likert_recode$QA5r2..Addressing._.Cost.of.Living <- factor(df_likert_recode$QA5r2..Addressing._.Cost.of.Living, levels = response_levels)
df_likert_recode$QA5r3..Addressing._.Homelessness <- factor(df_likert_recode$QA5r3..Addressing._.Homelessness, levels = response_levels)
df_likert_recode$QA5r4..Addressing._.Climate.Change <- factor(df_likert_recode$QA5r4..Addressing._.Climate.Change, levels = response_levels)
df_likert_recode$QA5r5..Addressing._.Healthcare.Needs <- factor(df_likert_recode$QA5r5..Addressing._.Healthcare.Needs, levels = response_levels)
df_likert_recode$QA5r6..Addressing._.Diversity...Inclusion <- factor(df_likert_recode$QA5r6..Addressing._.Diversity...Inclusion, levels = response_levels)
df_likert_recode$QA5r7..Addressing._.Social.Polarization <- factor(df_likert_recode$QA5r7..Addressing._.Social.Polarization, levels = response_levels)
df_likert_recode$QA5r8..Addressing._.Spiritual.Wellbeing <- factor(df_likert_recode$QA5r8..Addressing._.Spiritual.Wellbeing, levels = response_levels)
df_likert_recode$QA15..Opinion.on.Core.Operational.Expenses <- factor(df_likert_recode$QA15..Opinion.on.Core.Operational.Expenses, levels = response_levels)

# Charitable Services as Essential
df_likert_recode$QA6..Charitable.Services.as.Essential <- factor(df_likert_recode$QA6..Charitable.Services.as.Essential, levels = response_levels)

# Expenses - using similar levels for all relevant questions
expense_response_levels <- c("disagree", "not sure", "agree")
df_likert_recode$QA9r1..Expenses._.Insurance <- factor(df_likert_recode$QA9r1..Expenses._.Insurance, levels = expense_response_levels)
df_likert_recode$QA9r2..Expenses._.Technology <- factor(df_likert_recode$QA9r2..Expenses._.Technology, levels = expense_response_levels)
df_likert_recode$QA9r3..Expenses._.Program.Staff <- factor(df_likert_recode$QA9r3..Expenses._.Program.Staff, levels = expense_response_levels)
df_likert_recode$QA9r4..Expenses._.Fundraising.Staff <- factor(df_likert_recode$QA9r4..Expenses._.Fundraising.Staff, levels = expense_response_levels)
df_likert_recode$QA9r5..Expenses._.Fundraising.Activities <- factor(df_likert_recode$QA9r5..Expenses._.Fundraising.Activities, levels = expense_response_levels)
df_likert_recode$QA9r6..Expenses._.CEO.Salary <- factor(df_likert_recode$QA9r6..Expenses._.CEO.Salary, levels = expense_response_levels)
df_likert_recode$QA9r7..Expenses._.Marketing <- factor(df_likert_recode$QA9r7..Expenses._.Marketing, levels = expense_response_levels)
df_likert_recode$QA9r8..Expenses._.Volunteer.Recruitment <- factor(df_likert_recode$QA9r8..Expenses._.Volunteer.Recruitment, levels = expense_response_levels)
df_likert_recode$QA9r9..Expenses._.Legal.Services <- factor(df_likert_recode$QA9r9..Expenses._.Legal.Services, levels = expense_response_levels)
df_likert_recode$QA9r10..Expenses._.Financial.Audit <- factor(df_likert_recode$QA9r10..Expenses._.Financial.Audit, levels = expense_response_levels)
df_likert_recode$QA9r11..Expenses._.Advocacy <- factor(df_likert_recode$QA9r11..Expenses._.Advocacy, levels = expense_response_levels)
df_likert_recode$QA9r12..Expenses._.Board.Meetings <- factor(df_likert_recode$QA9r12..Expenses._.Board.Meetings, levels = expense_response_levels)

# Perception of Overhead Costs
overhead_levels <- c("under 10%", "10-20%", "21-30%", "31-40%", "41-50%", "51-60%", "over 60%")
df_likert_recode$QA11..Perception.of.Overhead.Costs <- factor(df_likert_recode$QA11..Perception.of.Overhead.Costs, levels = overhead_levels)

# Fair Compensation for Charity Workers
df_likert_recode$QA13..Fair.Compensation.for.Charity.Workers <- factor(df_likert_recode$QA13..Fair.Compensation.for.Charity.Workers, levels = response_levels)

# Household Income - Setting levels in a logical order reflecting income ranges
df_likert_recode$QNB4..Household.Income <- factor(df_likert_recode$QNB4..Household.Income,
                                    levels = c("under $50,000", "$50,000 to $69,999", "$70,000 to $99,999",
                                               "$100,000 to $149,999", "$150,000 to $249,999", 
                                               "$250,000 to $500,000", "more than $500,000", 
                                               "prefer not to say"))

# Political Party Preference - Unordered factors as political parties do not inherently have an order
df_likert_recode$QNB6..Political.Party.Preference <- factor(df_likert_recode$QNB6..Political.Party.Preference,
                                              levels = c("democratic party", "republican party", "the green party",
                                                         "constitution party", "natural law party", "libertarians",
                                                         "unsure / prefer not to answer"))

str(df_likert_recode)

# Function to convert all character columns in a dataframe to factors
convert_chars_to_factors <- function(df) {
  # Loop through each column in the dataframe
  for (col_name in names(df)) {
    # Check if the column is of type character
    if (is.character(df[[col_name]])) {
      # Convert the column to a factor
      df[[col_name]] <- factor(df[[col_name]])
    }
  }
  return(df)
}

# Apply the function to your dataframe
df_likert_recode <- convert_chars_to_factors(df_likert_recode)

# Ensure the variable is numeric
df_likert_recode$QA2r1..Trust._.Nonprofits <- as.numeric(as.character(df_likert_recode$QA2r1..Trust._.Nonprofits))


library(fastDummies)

# One-hot encode all factor variables in the dataset
df_encoded <- dummy_cols(
  df_likert_recode, 
  select_columns = names(df_likert_recode)[sapply(df_likert_recode, is.factor)], 
  remove_first_dummy = FALSE,  # Retain all levels without dropping the first
  remove_selected_columns = TRUE # Remove the original factor columns
)

colnames(df_likert_recode)
colnames(df_encoded)

# Get rid of special characters

names(df_encoded) <- gsub(" ", "_", trimws(names(df_encoded)))
names(df_encoded) <- gsub("\\s+", "_", trimws(names(df_encoded), whitespace = "[\\h\\v\\s]+"))
names(df_encoded) <- gsub("\\(", "_", names(df_encoded))
names(df_encoded) <- gsub("\\)", "_", names(df_encoded))
names(df_encoded) <- gsub("\\-", "_", names(df_encoded))
names(df_encoded) <- gsub("/", "_", names(df_encoded))
names(df_encoded) <- gsub("\\\\", "_", names(df_encoded)) 
names(df_encoded) <- gsub("\\?", "", names(df_encoded))
names(df_encoded) <- gsub("\\'", "", names(df_encoded))
names(df_encoded) <- gsub("\\,", "_", names(df_encoded))
names(df_encoded) <- gsub("\\$", "", names(df_encoded))
names(df_encoded) <- gsub("\\+", "_", names(df_encoded))

# Trim all values

# Loop over each column in the dataframe
df_encoded <- data.frame(lapply(df_encoded, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))

colnames(df_encoded)

ivs11 <- c(
#  "AGEGROUP.._COMPUTED_._.Age.group_under_25",
  "AGEGROUP.._COMPUTED_._.Age.group_25_44",
  "AGEGROUP.._COMPUTED_._.Age.group_45_64",
  "AGEGROUP.._COMPUTED_._.Age.group_65_",
#  "QS2..Gender.Identity_male",
  "QS2..Gender.Identity_female",
  "QS2..Gender.Identity_other",
#  "QNB4..Household.Income_under_50_000",
  "QNB4..Household.Income_50_000_to_69_999",
  "QNB4..Household.Income_70_000_to_99_999",
  "QNB4..Household.Income_100_000_to_149_999",
  "QNB4..Household.Income_150_000_to_249_999",
  "QNB4..Household.Income_250_000_to_500_000",
  "QNB4..Household.Income_more_than_500_000",
  "QNB4..Household.Income_prefer_not_to_say",
#  "USCR.._COMPUTED_._.Region.based.on.US.State_north_east",
  "USCR.._COMPUTED_._.Region.based.on.US.State_south",
  "USCR.._COMPUTED_._.Region.based.on.US.State_midwest",
  "USCR.._COMPUTED_._.Region.based.on.US.State_west",
#  "QNB6..Political.Party.Preference_democratic_party",
  "QNB6..Political.Party.Preference_republican_party",
  "QNB6..Political.Party.Preference_the_green_party",
  "QNB6..Political.Party.Preference_constitution_party",
  "QNB6..Political.Party.Preference_natural_law_party",
  "QNB6..Political.Party.Preference_libertarians",
  "QNB6..Political.Party.Preference_unsure___prefer_not_to_answer",
#  "QA10..Awareness.of.Overhead.Costs_no__have_no_idea",
  "QA10..Awareness.of.Overhead.Costs_yes__roughly",
  "QA14..Perceived.Operational.Efficiency",
#  "QA9r1..Expenses._.Insurance_disagree",
  "QA9r1..Expenses._.Insurance_not_sure",
  "QA9r1..Expenses._.Insurance_agree",
#  "QA9r2..Expenses._.Technology_disagree",
  "QA9r2..Expenses._.Technology_not_sure",
  "QA9r2..Expenses._.Technology_agree",
#  "QA9r3..Expenses._.Program.Staff_disagree",
  "QA9r3..Expenses._.Program.Staff_not_sure",
  "QA9r3..Expenses._.Program.Staff_agree",
#  "QA9r4..Expenses._.Fundraising.Staff_disagree",
  "QA9r4..Expenses._.Fundraising.Staff_not_sure",
  "QA9r4..Expenses._.Fundraising.Staff_agree",
#  "QA9r5..Expenses._.Fundraising.Activities_disagree",
  "QA9r5..Expenses._.Fundraising.Activities_not_sure",
  "QA9r5..Expenses._.Fundraising.Activities_agree",
#  "QA9r6..Expenses._.CEO.Salary_disagree",
  "QA9r6..Expenses._.CEO.Salary_not_sure",
  "QA9r6..Expenses._.CEO.Salary_agree",
#  "QA9r7..Expenses._.Marketing_disagree",
  "QA9r7..Expenses._.Marketing_not_sure",
  "QA9r7..Expenses._.Marketing_agree",
#  "QA9r8..Expenses._.Volunteer.Recruitment_disagree",
  "QA9r8..Expenses._.Volunteer.Recruitment_not_sure",
  "QA9r8..Expenses._.Volunteer.Recruitment_agree",
#  "QA9r9..Expenses._.Legal.Services_disagree",
  "QA9r9..Expenses._.Legal.Services_not_sure",
  "QA9r9..Expenses._.Legal.Services_agree",
#  "QA9r10..Expenses._.Financial.Audit_disagree",
  "QA9r10..Expenses._.Financial.Audit_not_sure",
  "QA9r10..Expenses._.Financial.Audit_agree",
#  "QA9r11..Expenses._.Advocacy_disagree",
  "QA9r11..Expenses._.Advocacy_not_sure",
  "QA9r11..Expenses._.Advocacy_agree",
#  "QA9r12..Expenses._.Board.Meetings_disagree",
  "QA9r12..Expenses._.Board.Meetings_not_sure",
  "QA9r12..Expenses._.Board.Meetings_agree",
  "QA8r1..Importance._.Accreditation",
  "QA8r2..Importance._.Governance",
  "QA8r3..Importance._.Passion.for.Cause",
  "QA8r4..Importance._.Personal_Family.Use",
  "QA8r5..Importance._.Community.Benefit",
  "QA8r6..Importance._.Effective.Management",
  "QA8r7..Importance._.Achieving.Goals",
  "QA8r8..Importance._.Reputation",
  "QA8r9..Importance._.Compelling.Stories",
  "QA8r10..Importance._.Leadership.Recognition",
  "QA8r11..Importance._.Financial.Transparency",
  "QA8r12..Importance._.Community.Impact.Info",
  "QA8r13..Importance._.Competitive.Overhead",
  "QA8r14..Importance._.Low.Administrative.Costs",
#  "QA3r11..Awareness._.Nonprofit.Employment_no",
  "QA3r11..Awareness._.Nonprofit.Employment_yes",
#  "QA3r12..Awareness._.Volunteer.Participation_no",
  "QA3r12..Awareness._.Volunteer.Participation_yes",
#  "QA3r13..Awareness._.Nonprofit.Services_no",
  "QA3r13..Awareness._.Nonprofit.Services_yes",
#  "QA3r14..Awareness._.Economic.Contribution_no",
  "QA3r14..Awareness._.Economic.Contribution_yes",
#  "QA3r15..Awareness._.Sectors.GDP.Contribution_no",
  "QA3r15..Awareness._.Sectors.GDP.Contribution_yes",
#  "QA3r16..Awareness._.Total.Nonprofits_no",
  "QA3r16..Awareness._.Total.Nonprofits_yes",
  "QA2r2..Trust._.Small.Businesses",
  "QA2r3..Trust._.Large.Corporations",
  "QA2r4..Trust._.Governments",
#  "QS8r1..Supported.Organization.Types._.Arts...Culture_no_to._arts_._culture",
  "QS8r1..Supported.Organization.Types._.Arts...Culture_arts_._culture",
#  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_no_to._environment_or_animal_welfare",
  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_environment_or_animal_welfare",
#  "QS8r3..Supported.Organization.Types._.Health_no_to._health",
  "QS8r3..Supported.Organization.Types._.Health_health",
 # "QS8r4..Supported.Organization.Types._.Hospitals_no_to._hospitals",
  "QS8r4..Supported.Organization.Types._.Hospitals_hospitals",
#  "QS8r5..Supported.Organization.Types._.Education.and.Research_no_to._education_and_research",
  "QS8r5..Supported.Organization.Types._.Education.and.Research_education_and_research",
#  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_no_to._international_organizations_or_disaster_relief",
  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_international_organizations_or_disaster_relief",
 # "QS8r7..Supported.Organization.Types._.Religious.Organizations_no_to._religious_organizations",
  "QS8r7..Supported.Organization.Types._.Religious.Organizations_religious_organizations",
#  "QS8r8..Supported.Organization.Types._.Social_Human.Services_no_to._social_human_services",
  "QS8r8..Supported.Organization.Types._.Social_Human.Services_social_human_services",
#  "QS8r9..Supported.Organization.Types._.Sports...Recreation_no_to._sports_._recreation",
  "QS8r9..Supported.Organization.Types._.Sports...Recreation_sports_._recreation",
#  "QS8r10..Supported.Organization.Types._.Universities...Colleges_no_to._universities_._colleges",
  "QS8r10..Supported.Organization.Types._.Universities...Colleges_universities_._colleges",
#  "QS8r888..Supported.Organization.Types._.Other_no_to._other__specify_.",
  "QS8r888..Supported.Organization.Types._.Other_other__specify_."
)



ivs12 <- c(
  #  "AGEGROUP.._COMPUTED_._.Age.group_under_25",
  "AGEGROUP.._COMPUTED_._.Age.group_25_44",
  "AGEGROUP.._COMPUTED_._.Age.group_45_64",
  "AGEGROUP.._COMPUTED_._.Age.group_65_",
  #  "QS2..Gender.Identity_male",
  "QS2..Gender.Identity_female",
  "QS2..Gender.Identity_other",
  #  "QNB4..Household.Income_under_50_000",
  "QNB4..Household.Income_50_000_to_69_999",
  "QNB4..Household.Income_70_000_to_99_999",
  "QNB4..Household.Income_100_000_to_149_999",
  "QNB4..Household.Income_150_000_to_249_999",
  "QNB4..Household.Income_250_000_to_500_000",
  "QNB4..Household.Income_more_than_500_000",
  "QNB4..Household.Income_prefer_not_to_say",
  #  "USCR.._COMPUTED_._.Region.based.on.US.State_north_east",
  "USCR.._COMPUTED_._.Region.based.on.US.State_south",
  "USCR.._COMPUTED_._.Region.based.on.US.State_midwest",
  "USCR.._COMPUTED_._.Region.based.on.US.State_west",
  #  "QNB6..Political.Party.Preference_democratic_party",
  "QNB6..Political.Party.Preference_republican_party",
  "QNB6..Political.Party.Preference_the_green_party",
  "QNB6..Political.Party.Preference_constitution_party",
  "QNB6..Political.Party.Preference_natural_law_party",
  "QNB6..Political.Party.Preference_libertarians",
  "QNB6..Political.Party.Preference_unsure___prefer_not_to_answer",
  #  "QA10..Awareness.of.Overhead.Costs_no__have_no_idea",
  "QA10..Awareness.of.Overhead.Costs_yes__roughly",
  
  #  "QA9r1..Expenses._.Insurance_disagree",
  "QA9r1..Expenses._.Insurance_not_sure",
  "QA9r1..Expenses._.Insurance_agree",
  #  "QA9r2..Expenses._.Technology_disagree",
  "QA9r2..Expenses._.Technology_not_sure",
  "QA9r2..Expenses._.Technology_agree",
  #  "QA9r3..Expenses._.Program.Staff_disagree",
  "QA9r3..Expenses._.Program.Staff_not_sure",
  "QA9r3..Expenses._.Program.Staff_agree",
  #  "QA9r4..Expenses._.Fundraising.Staff_disagree",
  "QA9r4..Expenses._.Fundraising.Staff_not_sure",
  "QA9r4..Expenses._.Fundraising.Staff_agree",
  #  "QA9r5..Expenses._.Fundraising.Activities_disagree",
  "QA9r5..Expenses._.Fundraising.Activities_not_sure",
  "QA9r5..Expenses._.Fundraising.Activities_agree",
  #  "QA9r6..Expenses._.CEO.Salary_disagree",
  "QA9r6..Expenses._.CEO.Salary_not_sure",
  "QA9r6..Expenses._.CEO.Salary_agree",
  #  "QA9r7..Expenses._.Marketing_disagree",
  "QA9r7..Expenses._.Marketing_not_sure",
  "QA9r7..Expenses._.Marketing_agree",
  #  "QA9r8..Expenses._.Volunteer.Recruitment_disagree",
  "QA9r8..Expenses._.Volunteer.Recruitment_not_sure",
  "QA9r8..Expenses._.Volunteer.Recruitment_agree",
  #  "QA9r9..Expenses._.Legal.Services_disagree",
  "QA9r9..Expenses._.Legal.Services_not_sure",
  "QA9r9..Expenses._.Legal.Services_agree",
  #  "QA9r10..Expenses._.Financial.Audit_disagree",
  "QA9r10..Expenses._.Financial.Audit_not_sure",
  "QA9r10..Expenses._.Financial.Audit_agree",
  #  "QA9r11..Expenses._.Advocacy_disagree",
  "QA9r11..Expenses._.Advocacy_not_sure",
  "QA9r11..Expenses._.Advocacy_agree",
  #  "QA9r12..Expenses._.Board.Meetings_disagree",
  "QA9r12..Expenses._.Board.Meetings_not_sure",
  "QA9r12..Expenses._.Board.Meetings_agree",
  "QA8r1..Importance._.Accreditation",
  "QA8r2..Importance._.Governance",
  "QA8r3..Importance._.Passion.for.Cause",
  "QA8r4..Importance._.Personal_Family.Use",
  "QA8r5..Importance._.Community.Benefit",
  "QA8r6..Importance._.Effective.Management",
  "QA8r7..Importance._.Achieving.Goals",
  "QA8r8..Importance._.Reputation",
  "QA8r9..Importance._.Compelling.Stories",
  "QA8r10..Importance._.Leadership.Recognition",
  "QA8r11..Importance._.Financial.Transparency",
  "QA8r12..Importance._.Community.Impact.Info",
  "QA8r13..Importance._.Competitive.Overhead",
  "QA8r14..Importance._.Low.Administrative.Costs",
  #  "QA3r11..Awareness._.Nonprofit.Employment_no",
  "QA3r11..Awareness._.Nonprofit.Employment_yes",
  #  "QA3r12..Awareness._.Volunteer.Participation_no",
  "QA3r12..Awareness._.Volunteer.Participation_yes",
  #  "QA3r13..Awareness._.Nonprofit.Services_no",
  "QA3r13..Awareness._.Nonprofit.Services_yes",
  #  "QA3r14..Awareness._.Economic.Contribution_no",
  "QA3r14..Awareness._.Economic.Contribution_yes",
  #  "QA3r15..Awareness._.Sectors.GDP.Contribution_no",
  "QA3r15..Awareness._.Sectors.GDP.Contribution_yes",
  #  "QA3r16..Awareness._.Total.Nonprofits_no",
  "QA3r16..Awareness._.Total.Nonprofits_yes",
  "QA2r1..Trust._.Nonprofits",
  "QA2r2..Trust._.Small.Businesses",
  "QA2r3..Trust._.Large.Corporations",
  "QA2r4..Trust._.Governments",
  #  "QS8r1..Supported.Organization.Types._.Arts...Culture_no_to._arts_._culture",
  "QS8r1..Supported.Organization.Types._.Arts...Culture_arts_._culture",
  #  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_no_to._environment_or_animal_welfare",
  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_environment_or_animal_welfare",
  #  "QS8r3..Supported.Organization.Types._.Health_no_to._health",
  "QS8r3..Supported.Organization.Types._.Health_health",
  # "QS8r4..Supported.Organization.Types._.Hospitals_no_to._hospitals",
  "QS8r4..Supported.Organization.Types._.Hospitals_hospitals",
  #  "QS8r5..Supported.Organization.Types._.Education.and.Research_no_to._education_and_research",
  "QS8r5..Supported.Organization.Types._.Education.and.Research_education_and_research",
  #  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_no_to._international_organizations_or_disaster_relief",
  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_international_organizations_or_disaster_relief",
  # "QS8r7..Supported.Organization.Types._.Religious.Organizations_no_to._religious_organizations",
  "QS8r7..Supported.Organization.Types._.Religious.Organizations_religious_organizations",
  #  "QS8r8..Supported.Organization.Types._.Social_Human.Services_no_to._social_human_services",
  "QS8r8..Supported.Organization.Types._.Social_Human.Services_social_human_services",
  #  "QS8r9..Supported.Organization.Types._.Sports...Recreation_no_to._sports_._recreation",
  "QS8r9..Supported.Organization.Types._.Sports...Recreation_sports_._recreation",
  #  "QS8r10..Supported.Organization.Types._.Universities...Colleges_no_to._universities_._colleges",
  "QS8r10..Supported.Organization.Types._.Universities...Colleges_universities_._colleges",
  #  "QS8r888..Supported.Organization.Types._.Other_no_to._other__specify_.",
  "QS8r888..Supported.Organization.Types._.Other_other__specify_."
)


# Main function to fit OLS and process results with standardized coefficients
library(ggplot2)
library(broom)
library(dplyr)

fit_ols_and_format <- function(data, predictors, response_vars, save_plots = FALSE) {
  ols_results_list <- list()
  formulas <- sapply(response_vars, function(rv) as.formula(paste(rv, "~", paste(predictors, collapse = " + "))))
  
  for (i in seq_along(response_vars)) {
    response_var <- response_vars[i]
    formula <- formulas[[i]]
    
    cat("\nProcessing response variable:", response_var, "\n")
    
    tryCatch({
      # Fit the OLS model
      lm_model <- lm(formula, data = data)
      model_summary <- summary(lm_model)
      r_squared <- model_summary$r.squared
      f_statistic <- model_summary$fstatistic[1]
      p_value <- pf(f_statistic, model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE)
      
      # Calculate standard deviations for numeric predictors and response variable
      numeric_predictors <- predictors[sapply(data[predictors], is.numeric)]
      sd_predictors <- sapply(numeric_predictors, function(col) sd(data[[col]], na.rm = TRUE))
      sd_response <- sd(data[[response_var]], na.rm = TRUE)
      
      # Standardize numeric predictors and response variable
      standardized_data <- data %>%
        mutate(across(all_of(numeric_predictors), scale, .names = "std_{col}")) %>%
        mutate(std_response = scale(.data[[response_var]]))
      
      # Adjust formula for standardized variables
      std_predictors <- c(
        paste0("std_", numeric_predictors),   # Standardized numeric predictors
        predictors[sapply(data[predictors], is.factor)]  # Leave factors unmodified
      )
      
      std_formula <- as.formula(paste("std_response ~", paste(std_predictors, collapse = " + ")))
      
      # Fit the model with standardized variables
      std_model <- lm(std_formula, data = standardized_data)
      std_coefs <- broom::tidy(std_model)
      
      # Extract tidy output from the lm model and add model metrics
      ols_results <- broom::tidy(lm_model) %>%
        mutate(ResponseVariable = response_var,
               R_Squared = r_squared,
               F_Statistic = f_statistic,
               P_Value = p_value,
               SD_Response = sd_response) %>%
        left_join(std_coefs %>% select(term, std_estimate = estimate), by = "term")
      
      # Add standard deviations of predictors
      ols_results <- ols_results %>%
        mutate(SD_Predictors = ifelse(term %in% names(sd_predictors), sd_predictors[term], NA))
      
      # Store the results in a list
      ols_results_list[[response_var]] <- ols_results
      
      # Save diagnostic plots if requested
      if (save_plots) {
        plot_dir <- file.path(getwd(), "plots")
        if (!dir.exists(plot_dir)) dir.create(plot_dir)
        
        # Residual plot
        res_plot <- ggplot(data = broom::augment(lm_model), aes(.fitted, .resid)) +
          geom_point() +
          geom_hline(yintercept = 0, linetype = "dashed") +
          labs(title = paste("Residual Plot for", response_var),
               x = "Fitted Values",
               y = "Residuals") +
          theme_minimal()
        
        # Save the plot
        ggsave(filename = file.path(plot_dir, paste0(response_var, "_residual_plot.png")),
               plot = res_plot, width = 8, height = 6)
      }
      
      cat("Successfully processed response variable:", response_var, "\n")
      
    }, error = function(e) {
      cat("Error encountered for response variable:", response_var, "\n")
      cat("Formula:", deparse(formula), "\n")
      cat("Error message:", conditionMessage(e), "\n")
    })
  }
  
  # Combine all results into a single data frame
  combined_results <- do.call(rbind, ols_results_list)
  
  return(combined_results)
}

colnames(df_encoded)

# Run the modified function
df_ols_dv11 <- fit_ols_and_format(
  data = df_encoded,
  predictors = ivs11,  # Replace with actual predictor names
  response_vars = dv11,  # Replace with actual response variable names
  
  save_plots = TRUE
)

df_ols_dv12 <- fit_ols_and_format(
  data = df_encoded,
  predictors = ivs12,  
  response_vars = dv12,  
  
  save_plots = TRUE
)


# Study 1
dv21 <- "QA15..Opinion.on.Core.Operational.Expenses"

# Merge the column into df_encoded from df_likert_recoded
df_encoded <- df_encoded %>%
  left_join(df_likert_recode %>%
              select(record..Record.number, QA15..Opinion.on.Core.Operational.Expenses),
            by = "record..Record.number")



colnames(df_encoded)


ivs21 <- c(
  # Perceptions of Charitable Services
#  "QA6..Charitable.Services.as.Essential_no",
  "QA6..Charitable.Services.as.Essential_not_sure",
  "QA6..Charitable.Services.as.Essential_yes",
  
  # Importance of Nonprofits
  "QA1r1..Importance._.Nonprofits",
  
  # Trust in Nonprofits
  "QA2r1..Trust._.Nonprofits",
  
  # Cause Support Across Regions and Political Affiliation
  "QS8r1..Supported.Organization.Types._.Arts...Culture_arts_._culture",
#  "QS8r1..Supported.Organization.Types._.Arts...Culture_no_to._arts_._culture",
  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_environment_or_animal_welfare",
#  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_no_to._environment_or_animal_welfare",
  "QS8r3..Supported.Organization.Types._.Health_health",
#  "QS8r3..Supported.Organization.Types._.Health_no_to._health",
  "QS8r4..Supported.Organization.Types._.Hospitals_hospitals",
#  "QS8r4..Supported.Organization.Types._.Hospitals_no_to._hospitals",
  "QS8r5..Supported.Organization.Types._.Education.and.Research_education_and_research",
#  "QS8r5..Supported.Organization.Types._.Education.and.Research_no_to._education_and_research",
  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_international_organizations_or_disaster_relief",
#  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_no_to._international_organizations_or_disaster_relief",
  "QS8r7..Supported.Organization.Types._.Religious.Organizations_religious_organizations",
#  "QS8r7..Supported.Organization.Types._.Religious.Organizations_no_to._religious_organizations",
  "QS8r8..Supported.Organization.Types._.Social_Human.Services_social_human_services",
#  "QS8r8..Supported.Organization.Types._.Social_Human.Services_no_to._social_human_services",
  "QS8r9..Supported.Organization.Types._.Sports...Recreation_sports_._recreation",
#  "QS8r9..Supported.Organization.Types._.Sports...Recreation_no_to._sports_._recreation",
  "QS8r10..Supported.Organization.Types._.Universities...Colleges_universities_._colleges",
#  "QS8r10..Supported.Organization.Types._.Universities...Colleges_no_to._universities_._colleges",
  
  # Impact of Crises on Nonprofits
#  "QA4..Impact.of.Crises.on.Charitable.Services_no",
  "QA4..Impact.of.Crises.on.Charitable.Services_dont_know",
  "QA4..Impact.of.Crises.on.Charitable.Services_yes___less_demand",
  "QA4..Impact.of.Crises.on.Charitable.Services_yes___more_demand",
  
  # Perception of Nonprofit Activity on Key Issues
#  "QA5r1..Addressing._.Food.Insecurity_no",
  "QA5r1..Addressing._.Food.Insecurity_not_sure",
  "QA5r1..Addressing._.Food.Insecurity_yes",
 # "QA5r2..Addressing._.Cost.of.Living_no",
  "QA5r2..Addressing._.Cost.of.Living_not_sure",
  "QA5r2..Addressing._.Cost.of.Living_yes",
#  "QA5r3..Addressing._.Homelessness_no",
  "QA5r3..Addressing._.Homelessness_not_sure",
  "QA5r3..Addressing._.Homelessness_yes",
#  "QA5r4..Addressing._.Climate.Change_no",
  "QA5r4..Addressing._.Climate.Change_not_sure",
  "QA5r4..Addressing._.Climate.Change_yes",
#  "QA5r5..Addressing._.Healthcare.Needs_no",
  "QA5r5..Addressing._.Healthcare.Needs_not_sure",
  "QA5r5..Addressing._.Healthcare.Needs_yes",
#  "QA5r6..Addressing._.Diversity...Inclusion_no",
  "QA5r6..Addressing._.Diversity...Inclusion_not_sure",
  "QA5r6..Addressing._.Diversity...Inclusion_yes",
#  "QA5r7..Addressing._.Social.Polarization_no",
  "QA5r7..Addressing._.Social.Polarization_not_sure",
  "QA5r7..Addressing._.Social.Polarization_yes",
#  "QA5r8..Addressing._.Spiritual.Wellbeing_no",
  "QA5r8..Addressing._.Spiritual.Wellbeing_not_sure",
  "QA5r8..Addressing._.Spiritual.Wellbeing_yes",
  
  # Demographics
#  "AGEGROUP.._COMPUTED_._.Age.group_under_25",
  "AGEGROUP.._COMPUTED_._.Age.group_25_44",
  "AGEGROUP.._COMPUTED_._.Age.group_45_64",
  "AGEGROUP.._COMPUTED_._.Age.group_65_",
#  "QS2..Gender.Identity_male",
  "QS2..Gender.Identity_female",
  "QS2..Gender.Identity_other",
#  "QNB4..Household.Income_under_50_000",
  "QNB4..Household.Income_50_000_to_69_999",
  "QNB4..Household.Income_70_000_to_99_999",
  "QNB4..Household.Income_100_000_to_149_999",
  "QNB4..Household.Income_150_000_to_249_999",
  "QNB4..Household.Income_250_000_to_500_000",
  "QNB4..Household.Income_more_than_500_000",
  "QNB4..Household.Income_prefer_not_to_say",
  "USCR.._COMPUTED_._.Region.based.on.US.State_south",
#  "USCR.._COMPUTED_._.Region.based.on.US.State_north_east",
  "USCR.._COMPUTED_._.Region.based.on.US.State_west",
  "USCR.._COMPUTED_._.Region.based.on.US.State_midwest",
#  "QNB6..Political.Party.Preference_democratic_party",
  "QNB6..Political.Party.Preference_republican_party",
  "QNB6..Political.Party.Preference_the_green_party",
  "QNB6..Political.Party.Preference_constitution_party",
  "QNB6..Political.Party.Preference_natural_law_party",
  "QNB6..Political.Party.Preference_libertarians",
  "QNB6..Political.Party.Preference_unsure___prefer_not_to_answer"
)



str(df_likert_recode)

library(VGAM)

multinomial_regression_vgam <- function(data, dependent_var, independent_vars) {
  
  # Calculate standard deviations for numeric predictors and dependent variable
  numeric_independent_vars <- independent_vars[sapply(data[independent_vars], is.numeric)]
  sd_predictors <- sapply(numeric_independent_vars, function(col) sd(data[[col]], na.rm = TRUE))
  sd_response <- sd(as.numeric(data[[dependent_var]]), na.rm = TRUE)
  
  # Create the formula for the multinomial logistic regression model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit the multinomial logistic regression model using vglm
  model <- vglm(formula, family = multinomial(refLevel = "no"), data = data, na.action = na.omit)
  
  # Print the summary of the model
  summary_model <- summary(model, coef = TRUE)
  print(summary_model)
  
  # Extract coefficients and compute standard errors, z-values, and p-values
  coefficients <- coef(summary_model)
  
  # Create a dataframe to store coefficient results
  coef_df <- as.data.frame(coefficients)
  
  # Calculate Odds Ratios from the log odds (coefficients)
  coef_df$OddsRatios <- exp(coef_df$Estimate)
  
  # Add standard deviations of predictors and response variable
  coef_df$SD_Predictors <- NA
  coef_df$SD_Response <- sd_response
  for (var in names(sd_predictors)) {
    coef_df$SD_Predictors[grep(var, rownames(coef_df))] <- sd_predictors[var]
  }
  
  # Extract model fit statistics
  deviance <- slot(model, "criterion")$deviance
  log_likelihood <- slot(model, "criterion")$loglikelihood
  residual_deviance <- slot(model, "ResSS")
  df_residual <- slot(model, "df.residual")
  df_total <- slot(model, "df.total")
  iter <- slot(model, "iter")
  rank <- slot(model, "rank")
  
  # Create a dataframe to store model fit statistics
  fit_stats_df <- data.frame(
    Deviance = deviance,
    LogLikelihood = log_likelihood,
    ResidualDeviance = residual_deviance,
    DF_Residual = df_residual,
    DF_Total = df_total,
    Iterations = iter,
    Rank = rank
  )
  
  # Returning a list of two dataframes
  return(list(Coefficients = coef_df, FitStatistics = fit_stats_df))
}


model_results <- multinomial_regression_vgam(df_encoded, dv21, ivs21)
df_mlg_21 <- model_results$Coefficients
df_mlg_21fit <- model_results$FitStatistics


colnames(df_likert_recode)

# Perform the join to add the desired column from df_likert_recode to df_encoded
library(dplyr)

str(df_likert_recode)

dv31 <- c(
  
  "QS5..Donated.to.Charity._Past.Year_"
)

colnames(df_likert_recode)

ivs31 <- c(
# Donation Amount
# "QS6..Charity.Donation.Amount._Past.Year__less_than_500",
#  "QS6..Charity.Donation.Amount._Past.Year__500_to_1000",
# "QS6..Charity.Donation.Amount._Past.Year__1001_to_2000",
#  "QS6..Charity.Donation.Amount._Past.Year__2001_to_3000",
#  "QS6..Charity.Donation.Amount._Past.Year__3001_to_4000",
#  "QS6..Charity.Donation.Amount._Past.Year__4001_to_5000",
#  "QS6..Charity.Donation.Amount._Past.Year__5000_",
#  "QS6..Charity.Donation.Amount._Past.Year__prefer_not_to_answer",
#  "QS6..Charity.Donation.Amount._Past.Year__NA",
  
  # Trends in Donation Behavior
#  "QS7..Donation.Trend._Past.Years__down",
#  "QS7..Donation.Trend._Past.Years__same",
#  "QS7..Donation.Trend._Past.Years__up",
#  "QS7..Donation.Trend._Past.Years__NA",
  
  # Types of Organizations Supported
  "QS8r1..Supported.Organization.Types._.Arts...Culture_arts_._culture",
#  "QS8r1..Supported.Organization.Types._.Arts...Culture_no_to._arts_._culture",
  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_environment_or_animal_welfare",
#  "QS8r2..Supported.Organization.Types._.Environment.or.Animal.Welfare_no_to._environment_or_animal_welfare",
 "QS8r3..Supported.Organization.Types._.Health_health",
#  "QS8r3..Supported.Organization.Types._.Health_no_to._health",
"QS8r4..Supported.Organization.Types._.Hospitals_hospitals",
#  "QS8r4..Supported.Organization.Types._.Hospitals_no_to._hospitals",
  "QS8r5..Supported.Organization.Types._.Education.and.Research_education_and_research",
#  "QS8r5..Supported.Organization.Types._.Education.and.Research_no_to._education_and_research",
  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_international_organizations_or_disaster_relief",
#  "QS8r6..Supported.Organization.Types._.International.Organizations.or.Disaster.Relief_no_to._international_organizations_or_disaster_relief",
 "QS8r7..Supported.Organization.Types._.Religious.Organizations_religious_organizations",
#  "QS8r7..Supported.Organization.Types._.Religious.Organizations_no_to._religious_organizations",
  "QS8r8..Supported.Organization.Types._.Social_Human.Services_social_human_services",
#  "QS8r8..Supported.Organization.Types._.Social_Human.Services_no_to._social_human_services",
  "QS8r9..Supported.Organization.Types._.Sports...Recreation_sports_._recreation",
#  "QS8r9..Supported.Organization.Types._.Sports...Recreation_no_to._sports_._recreation",
  "QS8r10..Supported.Organization.Types._.Universities...Colleges_universities_._colleges",
#  "QS8r10..Supported.Organization.Types._.Universities...Colleges_no_to._universities_._colleges",
  
  # Factors Influencing Donation Decisions
  "QA8r1..Importance._.Accreditation",
  "QA8r2..Importance._.Governance",
  "QA8r3..Importance._.Passion.for.Cause",
  "QA8r4..Importance._.Personal_Family.Use",
  "QA8r5..Importance._.Community.Benefit",
  "QA8r6..Importance._.Effective.Management",
  "QA8r7..Importance._.Achieving.Goals",
  "QA8r8..Importance._.Reputation",
  "QA8r9..Importance._.Compelling.Stories",
  "QA8r10..Importance._.Leadership.Recognition",
  "QA8r11..Importance._.Financial.Transparency",
  "QA8r12..Importance._.Community.Impact.Info",
  "QA8r13..Importance._.Competitive.Overhead",
  "QA8r14..Importance._.Low.Administrative.Costs",
  
  # Demographics
#  "AGEGROUP.._COMPUTED_._.Age.group_under_25",
  "AGEGROUP.._COMPUTED_._.Age.group_25_44",
  "AGEGROUP.._COMPUTED_._.Age.group_45_64",
  "AGEGROUP.._COMPUTED_._.Age.group_65_",
#  "QS2..Gender.Identity_male",
  "QS2..Gender.Identity_female",
  "QS2..Gender.Identity_other",
#  "QNB4..Household.Income_under_50_000",
  "QNB4..Household.Income_50_000_to_69_999",
  "QNB4..Household.Income_70_000_to_99_999",
  "QNB4..Household.Income_100_000_to_149_999",
  "QNB4..Household.Income_150_000_to_249_999",
  "QNB4..Household.Income_250_000_to_500_000",
  "QNB4..Household.Income_more_than_500_000",
  "QNB4..Household.Income_prefer_not_to_say",
  "USCR.._COMPUTED_._.Region.based.on.US.State_south",
#  "USCR.._COMPUTED_._.Region.based.on.US.State_north_east",
  "USCR.._COMPUTED_._.Region.based.on.US.State_west",
  "USCR.._COMPUTED_._.Region.based.on.US.State_midwest",
#  "QNB6..Political.Party.Preference_democratic_party",
  "QNB6..Political.Party.Preference_republican_party",
  "QNB6..Political.Party.Preference_the_green_party",
  "QNB6..Political.Party.Preference_constitution_party",
  "QNB6..Political.Party.Preference_natural_law_party",
  "QNB6..Political.Party.Preference_libertarians"#  "QNB6..Political.Party.Preference_unsure___prefer_not_to_answer"
)

# Perform the join to add the desired column from df_likert_recode to df_encoded
library(dplyr)

df_encoded <- df_encoded %>%
  left_join(
    df_likert_recode %>% 
      select(record..Record.number, QS5..Donated.to.Charity._Past.Year_),
    by = "record..Record.number"
  )


library(VGAM)

binary_logistic_vglm <- function(data, dependent_var, independent_vars) {
  
  # Ensure the dependent variable is binary
  if (!is.factor(data[[dependent_var]])) {
    stop("Dependent variable must be a binary factor.")
  }
  
  if (length(levels(data[[dependent_var]])) != 2) {
    stop("Dependent variable must have exactly two levels.")
  }
  
  # Calculate standard deviations for numeric predictors and dependent variable
  numeric_independent_vars <- independent_vars[sapply(data[independent_vars], is.numeric)]
  sd_predictors <- sapply(numeric_independent_vars, function(col) sd(data[[col]], na.rm = TRUE))
  sd_response <- ifelse(is.numeric(as.numeric(data[[dependent_var]])), 
                        sd(as.numeric(data[[dependent_var]]), na.rm = TRUE), NA)
  
  # Create the formula for the binary logistic regression model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit the binary logistic regression model using vglm with logitlink()
  model <- vglm(formula, family = binomialff(link = logitlink), data = data, na.action = na.omit)
  
  # Print the summary of the model
  summary_model <- summary(model)
  print(summary_model)
  
  # Extract coefficients and compute standard errors, z-values, and p-values
  coef_table <- coef(summary_model)
  coef_df <- data.frame(
    Predictor = rownames(coef_table),
    Estimate = coef_table[, "Estimate"],
    StdError = coef_table[, "Std. Error"],
    ZValue = coef_table[, "z value"],
    PValue = coef_table[, "Pr(>|z|)"]
  )
  
  # Calculate Odds Ratios from the log odds (coefficients)
  coef_df$OddsRatios <- exp(coef_df$Estimate)
  
  # Add standard deviations of predictors and response variable
  coef_df$SD_Predictors <- NA
  coef_df$SD_Response <- sd_response
  for (var in names(sd_predictors)) {
    coef_df$SD_Predictors[grep(var, coef_df$Predictor)] <- sd_predictors[var]
  }
  
  # Extract model fit statistics
  deviance <- model@criterion["deviance"]
  log_likelihood <- model@criterion["loglikelihood"]
  
  # Create a dataframe to store model fit statistics
  fit_stats_df <- data.frame(
    Deviance = deviance,
    LogLikelihood = log_likelihood,
    Iterations = model@iter,
    Rank = model@rank
  )
  
  # Returning a list of two dataframes
  return(list(Coefficients = coef_df, FitStatistics = fit_stats_df))
}


model_results2 <- binary_logistic_vglm(df_encoded, dv31, ivs31)
df_mlg_31 <- model_results2$Coefficients
df_mlg_31fit <- model_results2$FitStatistics

colnames(df_encoded)

library(dplyr)
library(tidyr)

# Function to diagnose missing data
diagnose_missing_data <- function(df, vars) {
  missing_summary <- df %>%
    select(all_of(vars)) %>%
    summarise(across(everything(), ~ sum(is.na(.)))) %>%
    pivot_longer(cols = everything(), names_to = "Variable", values_to = "Missing_Values")
  
  total_rows <- nrow(df)
  
  missing_summary <- missing_summary %>%
    mutate(Percentage_Missing = (Missing_Values / total_rows) * 100)
  
  return(missing_summary)
}

missing_data_summary <- diagnose_missing_data(df_encoded, ivs11)

str(df_likert_recode)


library(car)
library(dplyr)

# Function to calculate VIF
calculate_vif <- function(data, dependent_var, independent_vars) {
  # Check if dependent variable and independent variables exist in the dataset
  if (!(dependent_var %in% colnames(data)) || any(!independent_vars %in% colnames(data))) {
    stop("Dependent or independent variables are not present in the dataset.")
  }
  
  # Create a formula for the model
  formula <- as.formula(paste(dependent_var, "~", paste(independent_vars, collapse = " + ")))
  
  # Fit a linear model
  lm_model <- lm(formula, data = data)
  
  # Calculate VIF
  vif_values <- vif(lm_model)
  
  # Convert to a dataframe
  vif_df <- data.frame(
    Predictor = names(vif_values),
    VIF = vif_values
  )
  
  # Identify problematic predictors with VIF > 5 (common threshold)
  problematic_predictors <- vif_df %>%
    filter(VIF > 5) %>%
    arrange(desc(VIF))
  
  if (nrow(problematic_predictors) > 0) {
    cat("\nProblematic predictors with high VIF:\n")
    print(problematic_predictors)
  } else {
    cat("\nNo multicollinearity issues detected (VIF <= 5 for all predictors).\n")
  }
  
  return(vif_df)
}

# List of dependent and independent variables as objects
dv_list <- list(
  list(dv = dv11, ivs = ivs11),
  list(dv = dv12, ivs = ivs12),
  list(dv = "QA15..Opinion.on.Core.Operational.Expenses_yes", ivs = ivs21),
  list(dv = "QS5..Donated.to.Charity._Past.Year__yes", ivs = ivs31)
)

# Iterate over all dependent variables and calculate VIF
vif_results <- lapply(dv_list, function(vars) {
  cat("\nCalculating VIF for dependent variable:", vars$dv, "\n")
  vif_result <- calculate_vif(df_encoded, vars$dv, vars$ivs)
  vif_result$DependentVariable <- vars$dv
  return(vif_result)
})

# Combine results into a single dataframe
vif_summary <- do.call(rbind, vif_results)

# Display results
print(vif_summary)


# Function to calculate standardized estimates and errors
calculate_standardized_metrics <- function(df, estimate_col_name = "Estimate", stderr_col_name = "StdError") {
  # Check if necessary columns are present
  required_columns <- c(estimate_col_name, stderr_col_name, "SD_Predictors", "SD_Response")
  missing_columns <- setdiff(required_columns, colnames(df))
  if (length(missing_columns) > 0) {
    stop(paste("The following required columns are missing in the dataset:", paste(missing_columns, collapse = ", ")))
  }
  
  # Calculate standardized estimates
  df$Standardized_Estimate <- ifelse(
    !is.na(df$SD_Predictors) & !is.na(df$SD_Response),
    df[[estimate_col_name]] * (df$SD_Predictors / df$SD_Response),
    NA
  )
  
  # Calculate standardized errors
  df$Standardized_Error <- ifelse(
    !is.na(df$SD_Predictors) & !is.na(df$SD_Response),
    df[[stderr_col_name]] * (df$SD_Predictors / df$SD_Response),
    NA
  )
  
  return(df)
}

# Apply the function to each dataset with appropriate column names
df_ols_dv11 <- calculate_standardized_metrics(df_ols_dv11, estimate_col_name = "estimate", stderr_col_name = "std.error")
df_ols_dv12 <- calculate_standardized_metrics(df_ols_dv12, estimate_col_name = "estimate", stderr_col_name = "std.error")
df_mlg_21 <- calculate_standardized_metrics(df_mlg_21, estimate_col_name = "Estimate", stderr_col_name = "Std. Error")
df_mlg_31 <- calculate_standardized_metrics(df_mlg_31, estimate_col_name = "Estimate", stderr_col_name = "StdError")


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
  "Sample Chars" = df_freq,
  "Model 1.1" = df_ols_dv11,
  "Model 1.2" = df_ols_dv12,
  "Model 2.1" = df_mlg_21,
  "Model 2.1 Fit" = df_mlg_21fit,
  "Model 3.1" = df_mlg_31,
  "Model 3.1 Fit" = df_mlg_31fit,
  "Collinearity Assessment" = vif_summary
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")



library(ggplot2)

# Function to create dot-and-whisker plots
create_dot_whisker_plot <- function(data, 
                                    predictor_col = "Predictor", 
                                    std_estimate_col = "Standardized_Estimate", 
                                    std_error_col = "Standardized_Error") {
  
  # Check if required columns exist
  required_cols <- c(predictor_col, std_estimate_col, std_error_col)
  if (!all(required_cols %in% colnames(data))) {
    stop("Required columns are missing in the dataset.")
  }
  
  # Create the plot
  plot <- ggplot(data, aes(x = !!sym(std_estimate_col), y = !!sym(predictor_col))) +
    geom_point(color = "blue") +
    geom_errorbarh(aes(xmin = !!sym(std_estimate_col) - !!sym(std_error_col), 
                       xmax = !!sym(std_estimate_col) + !!sym(std_error_col)),
                   height = 0.2, color = "black") +
    labs(
      
      x = "Standardized Coefficient",
      y = "Predictor"
    ) +
    theme_minimal() +
    theme(
      strip.text = element_text(size = 10, face = "bold"),
      axis.text.y = element_text(size = 8),
      plot.margin = margin(t = 10, r = 10, b = 10, l = 10),
      plot.background = element_rect(fill = "white"),
      panel.background = element_rect(fill = "white"),
      legend.position = "none" # Remove legend
    ) +
    coord_cartesian(clip = "off")
  
  return(plot)
}

# Apply the function to each dataset
plot_ols_dv11 <- create_dot_whisker_plot(df_ols_dv11, predictor_col = "term", 
                                         std_estimate_col = "Standardized_Estimate", 
                                         std_error_col = "Standardized_Error")

plot_ols_dv12 <- create_dot_whisker_plot(df_ols_dv12, predictor_col = "term", 
                                         std_estimate_col = "Standardized_Estimate", 
                                         std_error_col = "Standardized_Error")

# For df_mlg_21 and df_mlg_31, reset the Predictor column from rownames if necessary
df_mlg_21$Predictor <- rownames(df_mlg_21)

df_mlg_21 <- df_mlg_21[!df_mlg_21[["Predictor"]] %in% c("QNB4..Household.Income_more_than_500_000:1", "QNB4..Household.Income_more_than_500_000:2"), ]

plot_mlg_21 <- create_dot_whisker_plot(df_mlg_21, predictor_col = "Predictor", 
                                       std_estimate_col = "Standardized_Estimate", 
                                       std_error_col = "Standardized_Error")

plot_mlg_31 <- create_dot_whisker_plot(df_mlg_31, predictor_col = "Predictor", 
                                       std_estimate_col = "Standardized_Estimate", 
                                       std_error_col = "Standardized_Error")

# Display plots
print(plot_ols_dv11)
print(plot_ols_dv12)
print(plot_mlg_21)
print(plot_mlg_31)

# Save the plot to an A4-sized file
ggsave(
  filename = "dot_whisker_plot_11.png", 
  plot = plot_ols_dv11, 
  width = 11.69, # Adjusted width to fit the content
  height = 11.69, # A4 height in inches
  units = "in",
  dpi = 300
)

# Save the plot to an A4-sized file
ggsave(
  filename = "dot_whisker_plot_12.png", 
  plot = plot_ols_dv12, 
  width = 11.69, # Adjusted width to fit the content
  height = 11.69, # A4 height in inches
  units = "in",
  dpi = 300
)

# Save the plot to an A4-sized file
ggsave(
  filename = "dot_whisker_plot_21.png", 
  plot = plot_mlg_21, 
  width = 11.69, # Adjusted width to fit the content
  height = 11.69, # A4 height in inches
  units = "in",
  dpi = 300
)

ggsave(
  filename = "dot_whisker_plot_31.png", 
  plot = plot_mlg_31, 
  width = 11.69, # Adjusted width to fit the content
  height = 11.69, # A4 height in inches
  units = "in",
  dpi = 300
)

