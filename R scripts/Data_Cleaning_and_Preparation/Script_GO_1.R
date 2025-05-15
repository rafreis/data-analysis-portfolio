library(googlesheets4)

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/trackrecord")

#gs4_auth()  # This will open a browser window for you to authenticate

# Specify the Google Sheet URL
sheet_url <- "https://https://docs.google.com/spreadsheets/d/17OkVv2pdhRM3VVosx_mVywjBl5GfPQ0Pe8TC8pRvIwg"

# Get all sheet names
sheet_names <- sheet_names(sheet_url)

# Function to read data from a single sheet and add the sheet name as a column
read_sheet_with_name <- function(sheet_name, sheet_url) {
  data <- read_sheet(sheet_url, sheet = sheet_name)
  data$SheetName <- sheet_name
  return(data)
}

# Read data from all sheets except the last one and combine into a single dataframe
all_data <- do.call(rbind, lapply(sheet_names, read_sheet_with_name, sheet_url = sheet_url))


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

# Apply the recoding function to the relevant columns and overwrite them
df_recoded <- all_data %>%
  mutate(
    `This GO has a compelling vision` = recode_responses(`This GO has a compelling vision`),
    `This GO has a clear offering and value proposition` = recode_responses(`This GO has a clear offering and value proposition`),
    `This GO has the leadership team to achieve its growth potential` = recode_responses(`This GO has the leadership team to achieve its growth potential`),
    `This GO has the expertise to be credible with Clients` = recode_responses(`This GO has the expertise to be credible with Clients`),
    `This GO has the team and resources to achieve its goals` = recode_responses(`This GO has the team and resources to achieve its goals`)
  )

# Employee Data

setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/trackrecord")

library(openxlsx)
df_employee <- read.xlsx("Org Effectiveness and Succession Planning_Round 1 5.15.24 2 (1).xlsx", sheet = "Employees")

df_recoded <- df_recoded %>%
  mutate(`Please insert your email so that we know you have completed this survey. Your answers will remain confidential.` = 
           trimws(`Please insert your email so that we know you have completed this survey. Your answers will remain confidential.`))

df_employee <- df_employee %>%
  mutate(Email = trimws(Email))

# Ensure unique email addresses in df_employee
df_employee <- df_employee %>%
  distinct(Email, .keep_all = TRUE)

df_final <- left_join(df_recoded, df_employee, by = c("Please insert your email so that we know you have completed this survey. Your answers will remain confidential." = "Email"))


# Rescale the 5-point scale to a 0-100 scale
rescale_to_100 <- function(value) {
  ((value - 1) / 4) * 100
}

df_final <- df_final %>%
  mutate(
    `This GO has a compelling vision` = rescale_to_100(`This GO has a compelling vision`),
    `This GO has a clear offering and value proposition` = rescale_to_100(`This GO has a clear offering and value proposition`),
    `This GO has the leadership team to achieve its growth potential` = rescale_to_100(`This GO has the leadership team to achieve its growth potential`),
    `This GO has the expertise to be credible with Clients` = rescale_to_100(`This GO has the expertise to be credible with Clients`),
    `This GO has the team and resources to achieve its goals` = rescale_to_100(`This GO has the team and resources to achieve its goals`)
  )

# Calculate the averages by SheetName and rename to Group
group_averages <- df_final %>%
  group_by(SheetName) %>%
  summarise(
    `Compelling Vision` = mean(`This GO has a compelling vision`, na.rm = TRUE),
    `Clear offering and value proposition` = mean(`This GO has a clear offering and value proposition`, na.rm = TRUE),
    `Leadership team to achieve its growth potential` = mean(`This GO has the leadership team to achieve its growth potential`, na.rm = TRUE),
    `Expertise to be credible with Clients` = mean(`This GO has the expertise to be credible with Clients`, na.rm = TRUE),
    `Team and resources to achieve its goals` = mean(`This GO has the team and resources to achieve its goals`, na.rm = TRUE)
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
group_averages_byorg <- df_final %>%
  group_by(SheetName, Group_Type) %>%
  summarise(
    `Compelling Vision` = mean(`This GO has a compelling vision`, na.rm = TRUE),
    `Clear offering and value proposition` = mean(`This GO has a clear offering and value proposition`, na.rm = TRUE),
    `Leadership team to achieve its growth potential` = mean(`This GO has the leadership team to achieve its growth potential`, na.rm = TRUE),
    `Expertise to be credible with Clients` = mean(`This GO has the expertise to be credible with Clients`, na.rm = TRUE),
    `Team and resources to achieve its goals` = mean(`This GO has the team and resources to achieve its goals`, na.rm = TRUE),
    Sample_Size = n()
  ) %>%
  rename(Group = SheetName)

# Convert the data to long format for easier plotting with ggplot2
group_averages_long <- group_averages_byorg %>%
  pivot_longer(cols = -c(Group, Group_Type, Sample_Size), names_to = "Metric", values_to = "Value") %>%
  filter(!is.na(Group_Type))  # Remove NA cases in Group_Type

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
  "Final GO Data" = df_final,
  "Overall GO Scores" = group_averages,
  "GO Scores by Role" = group_averages_byorg
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "GoData.xlsx")