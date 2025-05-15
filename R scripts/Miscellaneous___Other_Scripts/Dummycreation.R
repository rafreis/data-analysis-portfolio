# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/morgan")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("CleanData.xlsx")

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

colnames(df)


library(fastDummies)

vars <- c("What.is.your.age"                                                                   ,                                                               
 "What.is.your.gender._.Selected.Choice"                                                         ,                                                    
 "Which.best.describes.the.area.where.you.spent.the.majority.of.your.childhood._.Selected.Choice"   ,                                                 
 "What.is.your.annual.household.income"                                                              ,                                                
 "What.is.your.religious.affiliation._.Selected.Choice"                                               ,                                               
 "Is.the.religion.you.currently.identify.with.the.same.as.the.religion.you.were.raised.in."            ,                                              
 "How.often.do.you.participate.in.religious.or.spiritual.practices._e.g._.prayer_.meditation_.worship.services_" ,                                    
 "How.important.is.religion.or.spirituality.in.your.daily.life"                                                   ,                                   
"How.important.is.religion.or.spirituality.in.your.daily.life"                                                     ,                                 
 "Which.of.the.following.best.describes.your.political.ideology._.Selected.Choice"                                  ,                                 
 "What.is.your.current.marital.status"                                                                               ,                                
 "What.is.the.highest.level.of.education.you.have.completed"                                                          ,                               
 "What.is.your.current.employment.status"                                                                              ,                              
 "What.is.your.ethnicity._.Selected.Choice"                                                                             ,                             
 "What.is.your.sexual.orientation._.Selected.Choice")


df_recoded <- dummy_cols(df, select_columns = vars)


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
  "Data" = df_recoded
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "CleanData_dummy.xlsx")