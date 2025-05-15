library(googlesheets4)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/trackrecord/Third Wave/GO-Like Data")

# Authenticate with Google Sheets
#gs4_auth()  # Uncomment this line if authentication is needed

# Specify the Google Sheet URL
sheet_url <- "https://docs.google.com/spreadsheets/d/17cJGGDI50xzg7Xq8e6f7On57NYqD0TCu1e0xOcC1RqM/edit?gid=738801321#gid=738801321k"

# Get all sheet names
sheet_names <- sheet_names(sheet_url)

# Function to read sheet and standardize columns
read_and_standardize_sheet <- function(sheet_name, sheet_url, all_columns) {
  data <- read_sheet(sheet_url, sheet = sheet_name)
  missing_cols <- setdiff(all_columns, names(data))
  data[missing_cols] <- NA  # Add missing columns as NA
  data <- data[all_columns]  # Reorder columns to match the standard
  data$SheetName <- sheet_name
  return(data)
}

# First, determine all possible columns
all_columns <- unique(unlist(lapply(sheet_names, function(name) names(read_sheet(sheet_url, sheet = name)))))

# Read and combine data ensuring all data frames have the same structure
all_data <- do.call(rbind, lapply(sheet_names, read_and_standardize_sheet, sheet_url = sheet_url, all_columns = all_columns))

library(dplyr)
library(tidyr)
# Function to recode the responses
recode_responses <- function(response) {
  case_when(
    grepl("Strongly Disagree", response) ~ 1,
    grepl("Disagree", response) ~ 2,
    
    grepl("Strongly Agree", response) ~ 5,
    grepl("Agree", response) ~ 4,
    TRUE ~ NA_real_
  )
}

# Clean column names to remove hidden characters and excess spaces
colnames(all_data) <- gsub("\\s+", " ", colnames(all_data))  # Replace multiple spaces with a single space
colnames(all_data) <- trimws(colnames(all_data))  # Remove leading and trailing whitespaces

# Reapply the recoding function with the corrected column names
df_recoded <- all_data %>%
  mutate(
    `This function/market has a clear and inspiring vision` = recode_responses(`This function/market has a clear and inspiring vision`),
    `This function/market has a clear offering and value proposition` = recode_responses(`This function/market has a clear offering and value proposition`),
    `This function/market has the leadership team to deliver effectively to achieve its potential` = recode_responses(`This function/market has the leadership team to deliver effectively to achieve its potential`),
    `This function/market is and will continue to be a significant contributor to Clearwater's growth` = recode_responses(`This function/market is and will continue to be a significant contributor to Clearwater's growth`),
    `This function has the team and expertise to be consistently credible with clients/prospects` = recode_responses(`This function has the team and expertise to be consistently credible with clients/prospects`)
  )

all_columns

# Employee Data

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/trackrecord/Third Wave/GO-Like Data")

library(openxlsx)
df_employee <- read.xlsx("Org Effectiveness and Succession Planning_Round 1 5.15.24 2 (1).xlsx", sheet = "Employees")

df_recoded <- df_recoded %>%
  mutate(`Please add your email so that we know you have completed the survey. Your answers will remain confidential.` = 
           trimws(`Please add your email so that we know you have completed the survey. Your answers will remain confidential.`))

df_employee <- df_employee %>%
  mutate(Email = trimws(Email))

# Ensure unique email addresses in df_employee
df_employee <- df_employee %>%
  distinct(Email, .keep_all = TRUE)

df_final <- left_join(df_recoded, df_employee, by = c("Please add your email so that we know you have completed the survey. Your answers will remain confidential." = "Email"))


# Rescale the 5-point scale to a 0-100 scale
rescale_to_100 <- function(value) {
  ((value - 1) / 4) * 100
}

# Rescale the new columns to 100 and update column names in df_final
df_final <- df_final %>%
  mutate(
    `This function/market has a clear and inspiring vision` = rescale_to_100(`This function/market has a clear and inspiring vision`),
    `This function/market has a clear offering and value proposition` = rescale_to_100(`This function/market has a clear offering and value proposition`),
    `This function/market has the leadership team to deliver effectively to achieve its potential` = rescale_to_100(`This function/market has the leadership team to deliver effectively to achieve its potential`),
    `This function/market is and will continue to be a significant contributor to Clearwater's growth` = rescale_to_100(`This function/market is and will continue to be a significant contributor to Clearwater's growth`),
    `This function has the team and expertise to be consistently credible with clients/prospects` = rescale_to_100(`This function has the team and expertise to be consistently credible with clients/prospects`)
  )

# Calculate the averages by SheetName and rename to Group
group_averages <- df_final %>%
  group_by(SheetName) %>%
  summarise(
    `Clear and Inspiring Vision` = mean(`This function/market has a clear and inspiring vision`, na.rm = TRUE),
    `Clear Offering and Value Proposition` = mean(`This function/market has a clear offering and value proposition`, na.rm = TRUE),
    `Leadership Team to Deliver Effectively` = mean(`This function/market has the leadership team to deliver effectively to achieve its potential`, na.rm = TRUE),
    `Impact and Contribution to Growth and Innovation` = mean(`This function/market is and will continue to be a significant contributor to Clearwater's growth`, na.rm = TRUE),
    `Team and Resources to be Credible` = mean(`This function has the team and expertise to be consistently credible with clients/prospects`, na.rm = TRUE)
  ) %>%
  rename(Group = SheetName)



library(ggplot2)
library(tidyr)

# Convert the data to long format for easier plotting with ggplot2
group_averages_long <- group_averages %>%
  pivot_longer(cols = -Group, names_to = "Metric", values_to = "Value")

# Function to create and save plots
create_plot <- function(metric_name) {
  p <- ggplot(group_averages_long %>% filter(Metric == metric_name), aes(x = Group, y = round(Value, 0), fill = Group)) +
    geom_col() +
    geom_text(aes(label = round(Value, 0)), vjust = -0.3) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 90, hjust = 1),
          legend.position = "none") +
    labs(title = paste("Comparative Column Plot for", metric_name), x = "Group", y = metric_name)
  
  print(p)
  # Save the plot
  ggsave(paste0(metric_name, "_plot.png"), plot = p, width = 10, height = 6)
}

# List of metrics to plot
metrics <- unique(group_averages_long$Metric)

# Create and save plots for each metric
lapply(metrics, create_plot)


# Cut by GRoup

library(dplyr)
library(tidyr)
library(ggplot2)


# Filter the dataframe and create a new column for the group
df_final <- df_final %>%
  mutate(Group_Type = ifelse(`Org.Family` == '500 - S&M', 'Sales', 'Others'))

# Calculate the averages by SheetName and Group_Type and include sample size
# Calculate the averages by SheetName and Group_Type and include sample size
group_averages_byorg <- df_final %>%
  group_by(SheetName, Group_Type) %>%
  summarise(
    `Clear and Inspiring Vision` = mean(`This function/market has a clear and inspiring vision`, na.rm = TRUE),
    `Clear Offering and Value Proposition` = mean(`This function/market has a clear offering and value proposition`, na.rm = TRUE),
    `Leadership Team to Deliver Effectively` = mean(`This function/market has the leadership team to deliver effectively to achieve its potential`, na.rm = TRUE),
    `Impact and Contribution to Growth and Innovation` = mean(`This function/market is and will continue to be a significant contributor to Clearwater's growth`, na.rm = TRUE),
    `Team and Resources to be Credible` = mean(`This function has the team and expertise to be consistently credible with clients/prospects`, na.rm = TRUE),
    Sample_Size = n()  # Count of responses for each group
  ) %>%
  rename(Group = SheetName)  # Renaming SheetName to Group for clarity

# Convert the data to long format for easier plotting with ggplot2
group_averages_long <- group_averages_byorg %>%
  pivot_longer(
    cols = -c(Group, Group_Type, Sample_Size),  # Exclude these from the pivoting process
    names_to = "Metric", 
    values_to = "Value"
  ) %>%
  filter(!is.na(Group_Type))  # Exclude records where Group_Type is NA to clean up the data


# Function to create and save plots
create_plot <- function(metric_name) {
  p <- ggplot(group_averages_long %>% filter(Metric == metric_name, !is.na(Value)), aes(x = Group, y = round(Value, 0), fill = Group_Type)) +
    geom_col(position = "dodge") +
    geom_text(aes(label = round(Value, 0)), vjust = -0.3, position = position_dodge(width = 0.9)) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 90, hjust = 1),
          legend.position = "top") +
    labs(title = paste("Comparative Column Plot for", metric_name), x = "Group", y = metric_name, fill = "Group Type")
  
  print(p)
  # Save the plot
  ggsave(paste0(metric_name, "_plotByOrg.png"), plot = p, width = 10, height = 6)
}

# List of metrics to plot
metrics <- unique(group_averages_long$Metric)

# Create and save plots for each metric
lapply(metrics, create_plot)

# Display the resulting dataframe with sample sizes
print(group_averages_byorg)


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

data_list <- list(
  "Final Data" = df_final,
  "Overall Scores" = group_averages,
  "Scores by Role" = group_averages_byorg
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Go_Like_Data.xlsx")










# Additional Request


# Specify the Google Sheet URL
sheet_url <- "https://docs.google.com/spreadsheets/d/17cJGGDI50xzg7Xq8e6f7On57NYqD0TCu1e0xOcC1RqM/edit?gid=738801321#gid=738801321k"

# Get all sheet names
sheet_names <- sheet_names(sheet_url)

# Function to read sheet, adjust for dislocation, and standardize columns
read_adjust_and_segment <- function(sheet_name, sheet_url) {
  data <- read_sheet(sheet_url, sheet = sheet_name)
  
  if (sheet_name %in% c("Final Data", "Overall Scores", "Scores by Role")) {
    # Adjust for dislocated columns
    names(data)[-1] <- names(data)[-1]  # Shift column names to the right
    names(data)[1] <- "Segment"  # Set the first column as 'Segment'
    data <- data[, !is.na(names(data))]  # Remove any columns that might have turned NA due to name shifting
  }
  
  # Return segmented data if required
  if ("Segment" %in% names(data)) {
    return(data %>% group_by(Segment))
  } else {
    return(data)
  }
}

# Read and process data from each sheet, focusing on those needing adjustment
data_list <- lapply(sheet_names, read_adjust_and_segment, sheet_url = sheet_url)




library(googlesheets4)
library(dplyr)
library(tidyr)
library(ggplot2)


# Function to recode textual responses to numeric scores
recode_responses <- function(response) {
  case_when(
    grepl("Strongly Disagree", response) ~ 1,
    grepl("Disagree", response) ~ 2,
    grepl("Strongly Agree", response) ~ 5,
    grepl("Agree", response) ~ 4,
    TRUE ~ NA_real_
  )
}

# Function to rescale values from a 5-point scale to a 0-100 scale
rescale_to_100 <- function(value) {
  ((value - 1) / 4) * 100
}

# Function to process data based on whether it needs to be segmented by the first column
process_data <- function(data, sheet_name, segment = FALSE) {
  # Add sheet name as a column for identification
  data <- data %>%
    mutate(across(starts_with("This function"), recode_responses)) %>%
    mutate(across(starts_with("This function"), rescale_to_100)) %>%
    mutate(SheetName = sheet_name)  # Ensure SheetName is preserved through processing
  
  if (segment) {
    # Dynamically identify and use the first column for grouping
    first_col_name <- names(data)[1]
    grouped_data <- data %>%
      group_by(across(all_of(first_col_name)), .add = TRUE) %>%
      summarise(across(starts_with("This function"), mean, na.rm = TRUE), .groups = 'drop')
    overall_data <- data %>%
      group_by(SheetName) %>%
      summarise(across(starts_with("This function"), mean, na.rm = TRUE), .groups = 'drop')
    return(bind_rows(grouped_data, overall_data))
  } else {
    return(data %>%
             group_by(SheetName) %>%
             summarise(across(starts_with("This function"), mean, na.rm = TRUE), .groups = 'drop'))
  }
}

# Fetch all sheet names
sheet_names <- sheet_names(sheet_url)

# Apply processing based on the index of the dataframe in the list
combined_data <- bind_rows(lapply(seq_along(sheet_names), function(i) {
  sheet_data <- read_sheet(sheet_url, sheet = sheet_names[i])
  process_data(sheet_data, sheet_names[i], segment = i %in% c(6, 7, 8))
}))

# Optionally, visualize or further process combined_data
print(combined_data)


library(ggplot2)
library(tidyr)
library(dplyr)

# Assuming 'combined_data' is your DataFrame
# Convert data to long format, while preserving all other columns for potential grouping
combined_data_long <- combined_data %>%
  pivot_longer(cols = starts_with("This function"), names_to = "Question", values_to = "Score")

# Function to create plots by grouping variable
create_plots <- function(data, grouping_var, title_suffix) {
  filtered_data <- data %>%
    drop_na(!!sym(grouping_var), Score)  # Dropping NAs based on grouping variable dynamically
  
  p <- ggplot(filtered_data, aes_string(x = grouping_var, y = "Score", color = grouping_var)) +
    geom_point(position = position_dodge(0.5), size = 3) +
    geom_errorbar(aes(ymin = Score - sd(Score), ymax = Score + sd(Score)), width = 0.2, position = position_dodge(0.5)) +
    facet_wrap(~Question, scales = "free_y") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1)) +
    labs(title = paste("Dot and Whisker Plot of Scores by", title_suffix), x = title_suffix, y = "Score")
  
  print(p)
}

# Create plots for each sheet name
create_plots(combined_data_long, "SheetName", "Sheet Name")












# Function to create plots by grouping variable
create_plots <- function(data, grouping_var, title_suffix) {
  filtered_data <- data %>%
    drop_na(!!sym(grouping_var), Score)  # Dropping NAs based on grouping variable dynamically
  
  p <- ggplot(filtered_data, aes(x = .data[[grouping_var]], y = Score, color = .data[[grouping_var]])) +
    geom_point(position = position_dodge(0.5), size = 3) +
    geom_errorbar(aes(ymin = Score - sd(Score), ymax = Score + sd(Score)), width = 0.2, position = position_dodge(0.5)) +
    facet_wrap(~Question, scales = "free_y") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1)) +
    labs(title = paste("Dot and Whisker Plot of Scores by", title_suffix), x = title_suffix, y = "Score")
  
  print(p)
}

# Create plots for each sheet name
create_plots(combined_data_long, "SheetName", "Sheet Name")












# Function to create individual plots for each question by a given grouping variable
create_individual_plots <- function(data, grouping_var, title_suffix) {
  unique_questions <- unique(data$Question)
  
  for (question in unique_questions) {
    filtered_data <- data %>%
      filter(Question == question) %>%
      drop_na(!!sym(grouping_var), Score)  # Dropping NAs for the current grouping variable and scores
    
    # Create the plot
    p <- ggplot(filtered_data, aes(x = .data[[grouping_var]], y = Score, color = .data[[grouping_var]])) +
      geom_point(position = position_dodge(0.5), size = 3) +
      geom_errorbar(aes(ymin = Score - sd(Score), ymax = Score + sd(Score)), width = 0.2, position = position_dodge(0.5)) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1),
            legend.position = "none") +  # Hide the legend
      labs(title = paste(question, "-", title_suffix), x = title_suffix, y = "Score")
    
    # Generate a filename
    filename <- paste0(gsub("[^A-Za-z0-9]", "", question), "_", gsub("[^A-Za-z0-9]", "", title_suffix), ".png")
    
    # Save the plot
    ggsave(filename, plot = p, width = 10, height = 8)
  }
}

# Create individual plots for each question, grouped by 'SheetName'
create_individual_plots(combined_data_long, "SheetName", "Sheet Name")






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

data_list <- list(
  "Data - per Function/Category" = combined_data
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "Go_Like_Data_v2.xlsx")
