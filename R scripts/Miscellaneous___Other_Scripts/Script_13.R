# Load the readr library
library(readr)

# Define the file paths
certificates_path <- "C:/Users/rafre/Downloads/certificates.csv"
pp_complete_path <- "C:/Users/rafre/Downloads/pp-complete.csv"

# Define a function to read the file in chunks
read_in_chunks <- function(file_path, chunk_size = 1000000) {
  data_list <- list()
  
  callback <- function(x, pos) {
    data_list <<- append(data_list, list(x))
  }
  
  read_csv_chunked(file_path, callback = callback, chunk_size = chunk_size)
  
  # Combine the chunks into a single data frame
  final_data <- do.call(rbind, data_list)
  return(final_data)
}

# Load the CSV files in chunks
certificates_df <- read_in_chunks(certificates_path)
pp_complete_df <- read_in_chunks(pp_complete_path)

# Display the first few rows of each data frame to verify the data has been loaded correctly
cat("Certificates Data Frame:\n")
print(head(certificates_df))

cat("\nPP-Complete Data Frame:\n")
print(head(pp_complete_df))