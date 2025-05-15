library(readxl)
library(dplyr)
library(openxlsx)

# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/pkerwood/Transfer")

# List all Excel files
files <- list.files(pattern = "\\.xlsx$")

# Initialize lists to store dataframes
dfs_mentees <- list()
dfs_mentors <- list()

# Load each file into the appropriate list and standardize column names
for (file in files) {
  df <- readxl::read_excel(file)
  
  # Standardize column names by removing spaces, trimming, and converting to lowercase
  colnames(df) <- gsub(" ", "", tolower(trimws(colnames(df))))
  
  # Assign dataframes to mentee or mentor lists
  if (grepl("mentee", file, ignore.case = TRUE)) {
    dfs_mentees[[file]] <- df
  } else if (grepl("mentor", file, ignore.case = TRUE)) {
    dfs_mentors[[file]] <- df
  }
}

# Function to find columns appearing in at least three dataframes and record their first appearances
find_columns_with_min_frequency <- function(dfs, min_count = 3) {
  # Aggregate all column names with the dataframe index they appear in
  column_list <- lapply(seq_along(dfs), function(i) {
    setNames(rep(names(dfs[[i]]), each = 1), rep(i, length(names(dfs[[i]]))))
  })
  all_columns <- unlist(column_list)
  
  # Count the frequency of each column name
  column_counts <- table(all_columns)
  
  # Filter column names that appear at least 'min_count' times
  frequent_columns <- names(column_counts[column_counts >= min_count])
  
  # Determine the first appearance order for the frequent columns
  first_appearance_order <- sapply(frequent_columns, function(col) {
    min(as.numeric(names(all_columns[all_columns == col])))
  })
  
  # Return the columns sorted by their first appearance
  frequent_columns_sorted <- frequent_columns[order(first_appearance_order)]
  
  return(frequent_columns_sorted)
}

# Get columns that appear at least in three dataframes for both mentees and mentors
frequent_columns <- union(
  find_columns_with_min_frequency(dfs_mentees),
  find_columns_with_min_frequency(dfs_mentors)
)

# Function to combine dataframes ensuring all frequent columns are included
combine_dataframes <- function(dfs, columns) {
  dfs_combined <- bind_rows(lapply(dfs, function(df) {
    # Ensure all frequent columns are in the dataframe, add missing ones as NA
    missing_cols <- setdiff(columns, names(df))
    for (col in missing_cols) {
      df[[col]] <- NA  # Append missing columns without reordering
    }
    # Select and order columns as per the frequent columns list
    df <- df[, columns, drop = FALSE]
    return(df)
  }), .id = "ID")
  
  return(dfs_combined)
}

# Combine mentees and mentors dataframes
combined_mentees <- combine_dataframes(dfs_mentees, frequent_columns)
combined_mentors <- combine_dataframes(dfs_mentors, frequent_columns)

# Add source column
combined_mentees$Source <- "Mentee"
combined_mentors$Source <- "Mentor"


# Optionally, combine mentees and mentors into a single dataframe
combined_all <- bind_rows(combined_mentees, combined_mentors)

# Export the combined dataframe to an Excel file
write.xlsx(combined_all, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/pkerwood/Transfer/combined_data.xlsx", row.names = FALSE)

# Print completion message
cat("All dataframes have been combined using columns that appear in at least three dataframes and exported to combined_data.xlsx.\n")




library(dplyr)

# List of columns to remove from both combined_mentees and combined_mentors
columns_to_remove <- c(
  "responsetype", "stagedate(utc)", "tags", 
  "let'sgettheboringbutnecessarystuffoutofthewayfirst.pleaseconfirmyouhavereadourdatapolicy(foundonelevateme.co)andgiveus(elevate)yourpermissiontocollectandprocessyourdataforthepurposesofprovidingyouwithmentoringandcommunicatingwithyouabouttheprogramme.youcanwithdrawatanytimeandyouareabletoaskusforacopyofthedataweholdonyoubyemailingusatinfo@elevateme.co.",
  "other...43", "other...44", 
  "let'sgettheboringbutnecessarystuffoutofthewayfirst.pleaseconfirmyouhavereadour[datapolic](https://www.elevateme.co/privacy-policy)yandgiveus(elevate)yourpermissiontocollectandprocessyourdataforthepurposesofprovidingyouwithmentoringandcommunicatingwithyouabouttheprogramme.youcanwithdrawatanytimeandyouareabletoaskusforacopyofthedataweholdonyoubyemailingusat[info@elevateme.co](mailto:info@elevateme.co)",
  "no,i'drathergoincognitothanks", "other...58",
  "let'sgettheboringbutnecessarystuffoutofthewayfirst.pleaseconfirmyouhavereadourdatapolicy(foundonelevateme.co)andgiveus(elevate)yourpermissiontocollectandprocessyourdataforthepurposesofprovidingyouwithmentoringandcommunicatingwithyouabouttheprogramme.youcanwithdrawatanytimeandyouareabletoaskusforacopyofthedataweholdonyoubyemailingusatinfo@elevateme.co",
  "onascaleof1-10where1isnotatalland10istotallypositive,howconfidentdoyoufeelaboutthefuture?", "other...61", "other...62", 
  "other...63", "other...64", "Source", "networkid", "startdate(utc)", "submitdate(utc)"
)

# Remove specified columns from combined_mentees
combined_mentees <- combined_mentees %>%
  select(-all_of(columns_to_remove))

# Remove specified columns from combined_mentors
combined_mentors <- combined_mentors %>%
  select(-all_of(columns_to_remove))

# Optionally, you can combine mentees and mentors into a single dataframe after column removal
combined_all <- bind_rows(combined_mentees, combined_mentors)

# Write the cleaned combined dataframe to an Excel file
write.xlsx(combined_all, "C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/pkerwood/Transfer/cleaned_combined_data.xlsx", row.names = FALSE)

colnames(combined_mentees)
colnames(combined_mentors)

library(dplyr)

# Function to extract year and concatenate names
process_names_and_year <- function(df) {
  df %>%
    mutate(
      Year = substr(ID, 1, 4)
    )
}

# Apply the transformation function to combined_mentees and combined_mentors
combined_mentees <- process_names_and_year(combined_mentees)
combined_mentors <- process_names_and_year(combined_mentors)


library(dplyr)

# Function to rename columns by inserting underscores and improving readability
rename_columns <- function(df) {
  names(df) <- names(df) %>%
    gsub("andareyoureadytotakecontroloftherelationship_settheagendaandmakethemostofthisopportunity?", "and_are_you_ready_to_take_control_of_the_relationship_set_the_agenda_and_make_the_most_of_this_opportunity", .) %>%
    gsub("andwhatcompanydoyouworkfororareyoufreelanceorjuststartingout?", "and_what_company_do_you_work_for_or_are_you_freelance_or_just_starting_out", .) %>%
    gsub("anythingelseyou'dlikeustoknowaboutbeforeweletyouloose?", "anything_else_you'd_like_us_to_know_about_before_we_let_you_loose", .) %>%
    gsub("canyoucommitto6x1hoursessionsinthenext6monthsandtoansweringoursurveyattheend?", "can_you_commit_to_6_x_1_hour_sessions_in_the_next_6_months_and_to_answering_our_survey_at_the_end", .) %>%
    gsub("careerprogression", "career_progression", .) %>%
    gsub("confidence", "confidence", .) %>%
    gsub("conflictmanagement", "conflict_management", .) %>%
    gsub("creativityandsellingyourideasin", "creativity_and_selling_your_ideas_in", .) %>%
    gsub("howlonghaveyoubeenintheindustry?", "how_long_have_you_been_in_the_industry", .) %>%
    gsub("howtosellabusiness", "how_to_sell_a_business", .) %>%
    gsub("howtostartabusiness", "how_to_start_a_business", .) %>%
    gsub("leadership", "leadership", .) %>%
    gsub("makingitasafreelancer", "making_it_as_a_freelancer", .) %>%
    gsub("movingroleswithinoroutsidetheindustry", "moving_roles_within_or_outside_the_industry", .) %>%
    gsub("returningtoworkafteracareerbreak", "returning_to_work_after_a_career_break", .) %>%
    gsub("teamdevelopment", "team_development", .) %>%
    gsub("teamrecruitment", "team_recruitment", .) %>%
    gsub("what'syourcurrentjobtitle?", "what_is_your_current_job_title", .) %>%
    gsub("what'syourfirstname?", "what_is_your_first_name", .) %>%
    gsub("whatdoyouhopetogetfrommentoring?", "what_do_you_hope_to_get_from_mentoring", .) %>%
    gsub("whatittakestorunasuccessfulbusiness", "what_it_takes_to_run_a_successful_business", .) %>%
    gsub("worklifebalance", "work_life_balance", .) %>%
    gsub("wouldyoubeinterestedinbecomingamentor?", "would_you_be_interested_in_becoming_a_mentor", .) %>%
    gsub("andhowdoyoufeelabouttheovertimeyouwork?", "and_how_do_you_feel_about_the_overtime_you_work", .) %>%
    gsub("andyoursurname?", "and_your_surname", .) %>%
    gsub("arethereclearopportunitiesforyoutoprogress/gainpromotionwithinyourcurrentbusiness?", "are_there_clear_opportunities_for_you_to_progress_gain_promotion_within_your_current_business", .) %>%
    gsub("doyoufeelrecognisedandvaluedfortheworkthatyoudo?", "do_you_feel_recognised_and_valued_for_the_work_that_you_do", .) %>%
    gsub("haveyoubeenprovidedwithadequatetrainingorpersonaldevelopmentinordertofulfilyourcurrentrolefully?", "have_you_been_provided_with_adequate_training_or_personal_development_in_order_to_fulfil_your_current_role_fully", .) %>%
    gsub("howoftendoyouworkovertime(thatistimebeyondyourcontractedhours)?", "how_often_do_you_work_overtime_that_is_time_beyond_your_contracted_hours", .) %>%
    gsub("industryknowledge/insight", "industry_knowledge_insight", .) %>%
    gsub("jobskills(negotiation/communication/presentationetc)", "job_skills_negotiation_communication_presentation_etc", .) %>%
    gsub("whatcouldweofferfromelevatethatcouldhelpyou?", "what_could_we_offer_from_elevate_that_could_help_you", .) %>%
    gsub("whenwasthelasttimeyouhadaraise?", "when_was_the_last_time_you_had_a_raise", .) %>%
    gsub("whenwasthelasttimeyouwerepromoted?", "when_was_the_last_time_you_were_promoted", .) %>%
    gsub("andwhatcompanydoyouworkfororareyouabusinessownerorafreelancer?", "and_what_company_do_you_work_for_or_are_you_a_business_owner_or_a_freelancer", .) %>%
    gsub("howlonghaveyouworkedintheeventsindustry?", "how_long_have_you_worked_in_the_events_industry", .) %>%
    gsub("jobskills(negotiation/presentation/communicationetc)", "job_skills_negotiation_presentation_communication_etc", .) %>%
    gsub("whatdoyouhopetogetfrombeingmentoredbysomeone?", "what_do_you_hope_to_get_from_being_mentored_by_someone", .) %>%
    gsub("yes,youcanusemycompanyname", "yes_you_can_use_my_company_name", .) %>%
    gsub("yes,youcanusemyname", "yes_you_can_use_my_name", .) %>%
    gsub("andwhatcompanydoyouworkfororareyoufreelance?", "and_what_company_do_you_work_for_or_are_you_freelance", .) %>%
    gsub("whatdoyouhopetogetfrommentoringsomeone?", "what_do_you_hope_to_get_from_mentoring_someone", .) %>%
    gsub("wouldyouliketobementored?", "would_you_like_to_be_mentored", .) %>%
    gsub("haveyoumentoredbefore?", "have_you_mentored_before", .) %>%
    gsub("howdidyouhearaboutelevate?", "how_did_you_hear_about_elevate", .) %>%
    gsub("whatcouldweofferthatcouldhelpyourightnow?", "what_could_we_offer_that_could_help_you_right_now", .) %>%
    gsub("doyouhavementoringorcoachingexperience?", "do_you_have_mentoring_or_coaching_experience", .) %>%
    gsub("year", "year", .)
  
  return(df)
}

# Apply the renaming function to both dataframes
combined_mentees <- rename_columns(combined_mentees)
combined_mentors <- rename_columns(combined_mentors)

# STATISTICS

colnames(combined_mentees)
colnames(combined_mentors)


# Step 2: Consolidate company columns in combined_mentees and combined_mentors
combined_mentors <- combined_mentors %>%
  mutate(
    consolidated_company = coalesce(`and_what_company_do_you_work_for_or_are_you_freelance?`,
                                    `and_what_company_do_you_work_for_or_are_you_a_business_owner_or_a_freelancer?`,
                                    `and_what_company_do_you_work_for_or_are_you_freelance_or_just_starting_out?`),
    `and_what_company_do_you_work_for_or_are_you_freelance_or_just_starting_out?` = NULL,
    `and_what_company_do_you_work_for_or_are_you_a_business_owner_or_a_freelancer?` = NULL,
    `and_what_company_do_you_work_for_or_are_you_freelance?` = NULL
  )

combined_mentees <- combined_mentees %>%
  mutate(
    consolidated_company = coalesce(`and_what_company_do_you_work_for_or_are_you_freelance?`,
                                    `and_what_company_do_you_work_for_or_are_you_a_business_owner_or_a_freelancer?`,
                                    `and_what_company_do_you_work_for_or_are_you_freelance_or_just_starting_out?`),
    `and_what_company_do_you_work_for_or_are_you_freelance_or_just_starting_out?` = NULL,
    `and_what_company_do_you_work_for_or_are_you_a_business_owner_or_a_freelancer?` = NULL,
    `and_what_company_do_you_work_for_or_are_you_freelance?` = NULL
  )




library(dplyr)
library(tidyr)

# Combine first and last names into a single 'mentor_name' column, ensure it is in lowercase and trimmed
combined_mentors <- combined_mentors %>%
  mutate(
    mentor_name = tolower(trimws(ifelse(is.na(`what_is_your_first_name?`), 
                                        `what_is_your_first_name?`, 
                                        paste(`what_is_your_first_name?`, `and_your_surname?`, sep = " ")))),
    consolidated_company = tolower(trimws(consolidated_company)),
    Year = as.integer(Year)  # Corrected from 'year' to 'Year'
  ) %>%
  select(-c(`what_is_your_first_name?`, `and_your_surname?`))  # Remove original name columns to avoid duplication

# Combine first and last names into a single 'mentor_name' column, ensure it is in lowercase and trimmed
combined_mentees <- combined_mentees %>%
  mutate(
    mentee_name = tolower(trimws(ifelse(is.na(`what_is_your_first_name?`), 
                                        `what_is_your_first_name?`, 
                                        paste(`what_is_your_first_name?`, `and_your_surname?`, sep = " ")))),
    consolidated_company = tolower(trimws(consolidated_company)),
    Year = as.integer(Year)  # Corrected from 'year' to 'Year'
  ) %>%
  select(-c(`what_is_your_first_name?`, `and_your_surname?`))  # Remove original name columns to avoid duplication

# Clean mentor_name in combined_mentors by removing ' na' at the end of names
combined_mentors <- combined_mentors %>%
  mutate(
    mentor_name = gsub("\\s*na$", "", mentor_name)  # Remove ' na' at the end of the name
  )

# Clean mentee_name in combined_mentees by removing ' na' at the end of names
combined_mentees <- combined_mentees %>%
  mutate(
    mentee_name = gsub("\\s*na$", "", mentee_name)  # Remove ' na' at the end of the name
  )

library(dplyr)

# Calculate years of mentoring service, both at individual and company levels
mentor_company_participation <- combined_mentors %>%
  group_by(mentor_name, consolidated_company) %>%
  summarize(
    total_years = n_distinct(Year),
    first_year = min(Year),
    last_year = max(Year),
    .groups = 'drop'
  ) %>%
  mutate(duration_category = case_when(
    total_years == 1 ~ "One year only",
    total_years > 1 ~ paste(total_years, "years")
  ))

# Separate analysis for just mentors
mentor_participation <- combined_mentors %>%
  arrange(mentor_name, Year) %>%  # Ensure the data is ordered to retrieve the latest correctly
  group_by(mentor_name) %>%
  summarize(
    total_years = n_distinct(Year),
    first_year = min(Year),
    last_year = max(Year),
    latest_industry_experience = last(`how_long_have_you_been_in_the_industry?`),
    latest_events_industry_experience = last(`how_long_have_you_worked_in_the_events_industry?`),
    .groups = 'drop'
  ) %>%
  # Fill any NA in latest industry experience with latest events industry experience if available
  mutate(
    consolidated_experience = coalesce(latest_industry_experience, latest_events_industry_experience),
    duration_category = case_when(
      total_years == 1 ~ "One year only",
      total_years > 1 ~ paste(total_years, "years")
    )
  )

# Company participation with total years calculated
company_participation <- combined_mentors %>%
  group_by(consolidated_company) %>%
  summarize(
    total_mentors = n_distinct(mentor_name),
    total_years = sum(n_distinct(Year, mentor_name)),
    average_years_per_mentor = total_years / total_mentors,
    participation_years = n_distinct(Year),
    .groups = 'drop'
  ) %>%
  arrange(desc(total_mentors))


colnames(combined_mentors)



# 1. Assess mentee return in subsequent years
mentee_participation <- combined_mentees %>%
  group_by(mentee_name) %>%  
  summarize(
    total_participation_years = n_distinct(Year),
    .groups = 'drop'
  ) %>%
  mutate(
    is_returning = ifelse(total_participation_years > 1, "Returning", "One-time")
  )


# 3. Evaluate and rank companies based on participation
company_mentee_participation <- combined_mentees %>%
  group_by(consolidated_company) %>%
  summarize(
    total_mentees = n_distinct(mentee_name),
    participation_years = n_distinct(Year),
    .groups = 'drop'
  ) %>%
  arrange(desc(total_mentees))

# Merging mentor participation data for a complete company ranking
complete_company_mentee_participation <- company_mentee_participation %>%
  full_join(
    combined_mentors %>%
      group_by(consolidated_company) %>%
      summarize(
        total_mentors = n_distinct(mentor_name),
        mentor_participation_years = n_distinct(Year),
        .groups = 'drop'
      ),
    by = "consolidated_company"
  ) %>%
  mutate(
    total_participants = coalesce(total_mentees, 0L) + coalesce(total_mentors, 0L)
  ) %>%
  arrange(desc(total_participants))


#Mentees that become Mentors

# Normalize names in both datasets (if not already done)
combined_mentees <- combined_mentees %>%
  mutate(
    normalized_name = tolower(gsub("\\s+", "", mentee_name)),  # Remove spaces and convert to lowercase
    original_mentee_name = mentee_name  # Keep track of the original mentee names
  )

combined_mentors <- combined_mentors %>%
  mutate(
    normalized_name = tolower(gsub("\\s+", "", mentor_name)),  # Remove spaces and convert to lowercase
    original_mentor_name = mentor_name  # Keep track of the original mentor names
  )

# Identify mentees who become mentors and retain their original names
mentees_to_mentors <- combined_mentees %>%
  filter(normalized_name %in% combined_mentors$normalized_name) %>%
  distinct(normalized_name, .keep_all = TRUE) %>%
  # Join with the mentors dataset to get years they mentored and their original names
  inner_join(combined_mentors %>% select(normalized_name, Year, original_mentor_name) %>% distinct(), 
             by = "normalized_name") %>%
  # Keep only those instances where the mentorship year is later than the mentee year
  filter(Year.x < Year.y) %>%
  # Organize and rename for clarity
  rename(
    mentee_year = Year.x,
    mentor_year = Year.y
  ) %>%
  # Select the fields to display, including original names
  select(original_mentee_name, mentee_year, original_mentor_name, mentor_year)

# Display the result
print(mentees_to_mentors)

colnames(combined_mentees)

# List of columns to process
columns_to_analyze <- c(
  "career_progression", "confidence", "conflict_management", "creativity_and_selling_your_ideas_in", 
  "how_to_sell_a_business", "how_to_start_a_business", "leadership", "making_it_as_a_freelancer", 
  "moving_roles_within_or_outside_the_industry", "returning_to_work_after_a_career_break", 
  "team_development", "team_recruitment", "what_it_takes_to_run_a_successful_business", 
  "work_life_balance", "industry_knowledge_insight", "jobskills(negotiation/presentation/communicationetc)"
)

# Function to calculate frequency and percentage
calculate_frequencies <- function(df, column) {
  # Count non-NA values (assumed to be the selected responses)
  n_selected <- sum(!is.na(df[[column]]))
  # Calculate percentage
  percent_selected <- (n_selected / nrow(df)) * 100
  # Create a table with the results
  result <- data.frame(
    Question = column,
    N = n_selected,
    Percent = round(percent_selected, 2)
  )
  return(result)
}

# Apply the function to each column and combine results
frequency_tables_mentors <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_frequencies(combined_mentors, col)))

# Apply the function to each column and combine results
frequency_tables_mentees <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_frequencies(combined_mentees, col)))

# Function to calculate frequencies by Year
calculate_frequencies_by_year <- function(df, column) {
  result <- df %>%
    group_by(Year) %>%
    summarize(
      N = sum(!is.na(!!sym(column))),
      Percent = round((N / n()) * 100, 2),
      .groups = 'drop'
    )
  result$Question <- column
  return(result)
}

# Apply the function to each column for mentors and mentees, grouped by Year
frequency_tables_mentors_by_year <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_frequencies_by_year(combined_mentors, col)))
frequency_tables_mentees_by_year <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_frequencies_by_year(combined_mentees, col)))



# List of columns to analyze for both mentees and mentors
columns_to_analyze <- c(
  "and_how_do_you_feel_about_the_overtime_you_work?", 
  "do_you_feel_recognised_and_valued_for_the_work_that_you_do?", 
  "have_you_been_provided_with_adequate_training_or_personal_development_in_order_to_fulfil_your_current_role_fully?", 
  "would_you_like_to_be_mentored?", 
  "have_you_mentored_before?", 
  "how_did_you_hear_about_elevate?", 
  "what_could_we_offer_that_could_help_you_right_now?"
)

# Function to calculate frequencies for each column
calculate_response_frequencies <- function(df, column) {
  # Create a frequency table for the column
  freq_table <- table(df[[column]], useNA = "ifany")
  
  # Convert the frequency table to a data frame
  result <- as.data.frame(freq_table, stringsAsFactors = FALSE)
  colnames(result) <- c("Response", "N")
  
  # Calculate the percentage
  result$Percent <- round((result$N / sum(result$N, na.rm = TRUE)) * 100, 2)
  
  # Add the question name for reference
  result$Question <- column
  
  return(result)
}

# Apply the function to each column for mentees
frequency_tables_mentees2 <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_response_frequencies(combined_mentees, col)))

# Apply the function to each column for mentors
frequency_tables_mentors2 <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_response_frequencies(combined_mentors, col)))


# Function to calculate response frequencies by Year
calculate_response_frequencies_by_year <- function(df, column) {
  result <- df %>%
    group_by(Year) %>%
    count(!!sym(column)) %>%
    mutate(Percent = round((n / sum(n)) * 100, 2)) %>%
    ungroup() %>%
    rename(Response = !!sym(column))
  result$Question <- column
  return(result)
}

# Apply the function to each column for mentors and mentees, grouped by Year
frequency_tables_mentees2_by_year <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_response_frequencies_by_year(combined_mentees, col)))
frequency_tables_mentors2_by_year <- do.call(rbind, lapply(columns_to_analyze, function(col) calculate_response_frequencies_by_year(combined_mentors, col)))










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
  "Mentees Data" = combined_mentees,
  "Mentors Data" = combined_mentors,
  "Mentors Statistics" = mentor_participation,
  "Company Participation" = company_participation,
  "Mentee Participation" = mentee_participation,
  "Company Mentee Participation" = complete_company_mentee_participation,
  "Mentees to Mentors" = mentees_to_mentors,
  "Mentor Frequency by Year" = frequency_tables_mentors_by_year,
  "Mentee Frequency by Year" = frequency_tables_mentees_by_year,
  "Mentor Response Frequencies" = frequency_tables_mentors2,
  "Mentee Response Frequencies" = frequency_tables_mentees2
)


# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")
