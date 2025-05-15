setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/matthewhill11")

# If SPSS
library(haven)

# Declaring variables

vars_waves <- c("wave22", "wave25")

var_constituencies <- c("pconW22", "pconW25")

var_approval <- c("likeLabW22", "likeLabW25")

var_id <- "id"

# Combine all the columns into one vector
columns_to_import <- c(var_id, vars_waves, var_constituencies, var_approval)

# Read the SPSS file with only the specified columns
df <- read_sav("BES2019_W25_Panel_v25.2.sav", col_select = all_of(columns_to_import))



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

# Declaring variables

vars_waves <- c("wave22", "wave25")

var_constituencies <- c("pconW22", "pconW25")

var_approval <- "likeLab"

var_id <- "id"

# Filter the dataset for cases where wave22 and wave25 both equal 1
df_wave_filtered <- df %>%
  filter(wave22 == 1 & wave25 == 1)

# Filter the dataset for specific combinations of likeLabW22 and likeLabW25 values
df_pcon_filtered <- df_wave_filtered %>%
  filter((pconW22 %in% c(302, 303, 304) & pconW25 %in% c(302, 303, 304)) |
           (pconW22 %in% c(324, 180, 477) & pconW25 %in% c(324, 180, 477)))

df_no_dontknow <- df_pcon_filtered %>%
  filter(likeLabW22 != 9999 & likeLabW25 != 9999)

# Add a column to flag the treatment groups based on pconW25 values
df_no_dontknow <- df_no_dontknow %>%
  mutate(treatment_group = if_else(pconW25 %in% c(302, 303, 304), "Yes", "No"))

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
  "Cleaned Data" = df_no_dontknow
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "CleanedData.xlsx")
