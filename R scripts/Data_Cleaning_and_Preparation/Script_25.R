# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/elizli/")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df_survey <- read.xlsx("research data (1) (1)_std.xlsx", sheet = "outcome data")

df_interview <- read.xlsx("research data (1) (1)_std.xlsx", sheet = "interviews")

str(df_survey)

# Get rid of special characters

names(df_survey) <- gsub(" ", "_", trimws(names(df_survey)))
names(df_survey) <- gsub("\\s+", "_", trimws(names(df_survey), whitespace = "[\\h\\v\\s]+"))
names(df_survey) <- gsub("\\(", "_", names(df_survey))
names(df_survey) <- gsub("\\)", "_", names(df_survey))
names(df_survey) <- gsub("\\-", "_", names(df_survey))
names(df_survey) <- gsub("/", "_", names(df_survey))
names(df_survey) <- gsub("\\\\", "_", names(df_survey)) 
names(df_survey) <- gsub("\\?", "", names(df_survey))
names(df_survey) <- gsub("\\'", "", names(df_survey))
names(df_survey) <- gsub("\\,", "_", names(df_survey))
names(df_survey) <- gsub("\\$", "", names(df_survey))
names(df_survey) <- gsub("\\+", "_", names(df_survey))

# Trim all values

# Loop over each column in the dataframe
df_survey <- data.frame(lapply(df_survey, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))


names(df_interview) <- gsub(" ", "_", trimws(names(df_interview)))
names(df_interview) <- gsub("\\s+", "_", trimws(names(df_interview), whitespace = "[\\h\\v\\s]+"))
names(df_interview) <- gsub("\\(", "_", names(df_interview))
names(df_interview) <- gsub("\\)", "_", names(df_interview))
names(df_interview) <- gsub("\\-", "_", names(df_interview))
names(df_interview) <- gsub("/", "_", names(df_interview))
names(df_interview) <- gsub("\\\\", "_", names(df_interview)) 
names(df_interview) <- gsub("\\?", "", names(df_interview))
names(df_interview) <- gsub("\\'", "", names(df_interview))
names(df_interview) <- gsub("\\,", "_", names(df_interview))
names(df_interview) <- gsub("\\$", "", names(df_interview))
names(df_interview) <- gsub("\\+", "_", names(df_interview))

# Trim all values

# Loop over each column in the dataframe
df_interview <- data.frame(lapply(df_interview, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))


library(moments)

# Define the numeric variables
numeric_vars <- c("Birthweight._g_", "Age", "Parity", "BMI")

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  results <- data.frame(Variable = character(), Mean = numeric(), Median = numeric(),
                        SEM = numeric(), SD = numeric(), Skewness = numeric(), Kurtosis = numeric(),
                        stringsAsFactors = FALSE)
  
  for (var in desc_vars) {
    variable_data <- data[[var]]
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean(variable_data, na.rm = TRUE),
      Median = median(variable_data, na.rm = TRUE),
      SEM = sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data))),
      SD = sd(variable_data, na.rm = TRUE),
      Skewness = skewness(variable_data, na.rm = TRUE),
      Kurtosis = kurtosis(variable_data, na.rm = TRUE)
    ))
  }
  
  return(results)
}

# Calculate descriptive statistics for numeric variables
df_descriptive_stats <- calculate_descriptive_stats(df_survey, numeric_vars)


# Convert the data to a long format suitable for plotting
df_long <- pivot_longer(df_survey, cols = numeric_vars, names_to = "variable", values_to = "value")

# Create boxplots for each variable on a single figure with separate y-axes
plot <- ggplot(df_long, aes(x = factor(1), y = value, fill = variable)) +
  geom_boxplot(outlier.colour = "red", outlier.shape = 19) +
  stat_summary(fun = mean, geom = "point", shape = 23, size = 3, fill = "white") +  # Add mean as white dot
  facet_wrap(~variable, scales = "free_y") +  # Allows each plot to have a free y-axis
  scale_fill_brewer(palette = "Set3") +  # Using a color palette for distinction
  theme_minimal() +
  theme(
    strip.background = element_rect(fill = "#D55E00", colour = "#CC79A7", size = 1, linetype = "solid"),
    strip.text = element_text(size = 12, color = "white"),
    panel.spacing = unit(2, "lines"),  # Adjust spacing between facets
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank()
  )

# Print the plot to the R console
print(plot)

# Optionally, save the plot
ggsave("Faceted_Boxplots_with_Means.png", plot = plot, width = 12, height = 8, dpi = 300)

# Define the categorical variables
categorical_vars <- c("On.Aspirin", "PIH_PET", "Chronic.HTN", "Ges.at.delivery",
                      "IOL._reason", "Mode.of.delivery", "IUGR", "Apgars",
                      "IUD_neonatal.death", "method.of.conception", "Ethnicity",
                      "English.1st.language", "Smoker")

# Function to create frequency tables
create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list()
  
  for (category in categories) {
    data[[category]] <- factor(data[[category]])
    counts <- table(data[[category]])
    freq_table <- data.frame(
      Category = rep(category, length(counts)),
      Level = names(counts),
      Count = as.integer(counts),
      Percentage = (counts / sum(counts)) * 100
    )
    all_freq_tables[[category]] <- freq_table
  }
  
  do.call(rbind, all_freq_tables)
}

# Generate frequency tables for categorical variables
df_freq_tables <- create_frequency_tables(df_survey, categorical_vars)



# Define categorical variables
categorical_vars <- c("Method.of.conception", "Ethnicity", "Education", "English.1st.language", 
                      "Religion", "Smoker", "Have.you.had.high.blood.pressure.or.pre_eclampsia.in.pregnancy.before",
                      "Have.you.ever.had.high.blood.pressure.before.pregnancy",
                      "Are.you.taking.or.have.you.been.advised.to.start.taking.Aspirin.150mg.every.evening.in.this.pregnancy",
                      "Did.you.also.decline.the.combined.screening.test.which.screens.for.Downs._Edwards.and.Pataus.syndromes",
                      "Do.you.feel.that.you.understand.what.pre_eclampsia.is",
                      "Have.any.of.your.friends.or.family.had.pre_eclampsia",
                      "Did.you.feel.that.you.understood.how.the.pre_eclampsia.screening.would.be.done.at.your.appointment",
                      "Did.you.understand.what.the.test.results.would.be.and.what.would.be.offered",
                      "Were.you.aware.of.the.use.of.aspirin.in.pregnancy.to.reduce.the.chance.of.pre_eclampsia",
                      "was.the.pre_eclampsia.screening.discussed.with.you.before.the.scan.appointment.or.any.information.provided.",
                      "Had.you.already.decided.whether.to.accept.or.decline.the.pre_eclampsia.screening.before.your.scan.appointment")

# Generate frequency tables for categorical variables
df_freq_tables_interview <- create_frequency_tables(df_interview, categorical_vars)







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
  
  "Freq Table Interview" = df_freq_tables_interview, 
  "Freq Table Survey" = df_freq_tables,
  "Descriptive Stats" = df_descriptive_stats
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")
