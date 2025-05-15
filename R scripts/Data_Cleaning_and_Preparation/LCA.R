


## FORMULA TO REVERT THE ENCODING ##

get_columnwise_encoding_mapping <- function(encoded_data, original_data) {
  encoding_mapping <- list()
  
  for (i in 1:ncol(encoded_data)) {
    encoding_mapping[[colnames(encoded_data)[i]]] <- list()
    unique_values <- unique(encoded_data[[i]])
    for (val in unique_values) {
      # get the row indexes where the encoded value equals the current unique value
      idx <- which(encoded_data[[i]] == val)
      # get the corresponding original values
      original_values <- unique(original_data[[i]][idx])
      # add the mapping to the list
      encoding_mapping[[colnames(encoded_data)[i]]][[as.character(val)]] <- original_values
    }
  }
  
  return(encoding_mapping)
}

# Now you can use it as follows:
columnwise_encoding_mapping <- get_columnwise_encoding_mapping(encoded_data_numeric, combined_data)

# EXPORT THE ENCODING MAPPING #

# Load the required package
library(openxlsx)

# Function to flatten the encoding mapping into a data frame
flatten_encoding_mapping <- function(encoding_mapping) {
  flat_mapping <- data.frame()
  
  for (colname in names(encoding_mapping)) {
    for (encoded_value in names(encoding_mapping[[colname]])) {
      original_values <- encoding_mapping[[colname]][[encoded_value]]
      flat_mapping <- rbind(flat_mapping, 
                            data.frame(Column = colname,
                                       EncodedValue = encoded_value,
                                       OriginalValue = paste(original_values, collapse = ", ")))
    }
  }
  
  return(flat_mapping)
}

# Flatten the encoding mapping
flat_mapping <- flatten_encoding_mapping(columnwise_encoding_mapping)

# Write the flattened encoding mapping to an Excel file
write.xlsx(flat_mapping, "EncodingMapping.xlsx")

#### PROFILE ANALYSIS ####

# Extract the most likely class membership for each respondent
most_likely_class <- lca_result$predclass

# Add the most likely class membership to the original encoded data
encoded_data_with_class <- cbind(encoded_data_numeric, Cluster = most_likely_class)

# Create a function to compute the mode of a vector
compute_mode <- function(x) {
  ux <- unique(x)
  ux[which.max(tabulate(match(x, ux)))]
}

# Group by class and compute the mode for each variable
class_profiles <- encoded_data_with_class %>%
  group_by(Cluster) %>%
  summarize(across(starts_with("Q"), compute_mode))

# Print class profiles
print(class_profiles)

# Count the number of individuals in each class
class_counts <- table(most_likely_class)
print(class_counts)

# Calculate class proportions
class_proportions <- class_counts / sum(class_counts)
print(class_proportions)

# Create a data frame with Cluster and encoded data
class_responses <- cbind(Cluster = most_likely_class, encoded_data_numeric)

# Reshape the data to long format for easier analysis
class_responses_long <- class_responses %>%
  pivot_longer(starts_with("Q"), names_to = "Question", values_to = "Response")

# Calculate the frequency of each response for each class
class_response_frequencies <- class_responses_long %>%
  mutate(Question = fct_inorder(Question)) %>%
  group_by(Cluster, Question, Response) %>%
  summarize(Frequency = n(), .groups = "drop") %>%
  arrange(Cluster, Question, desc(Frequency))

# Print the frequencies
print(class_response_frequencies)

# Calculate proportions of each response for each class and question
class_response_proportions <- class_responses_long %>%
  mutate(Question = fct_inorder(Question)) %>%
  group_by(Cluster, Question) %>%
  count(Response) %>%
  mutate(Prop = n / sum(n)) %>%
  arrange(Cluster, Question, desc(Prop))

# Create table for class counts
class_counts_df <- as.data.frame(class_counts)
colnames(class_counts_df) <- c("Class", "Count")

# Create table for class proportions
class_proportions_df <- as.data.frame(class_proportions)
colnames(class_proportions_df) <- c("Class", "Proportion")

# Create an 'ID' variable that records the order of the rows
class_profiles$ID <- seq_len(nrow(class_profiles))

# Reshape the data into a 'long' format
class_profiles_long <- melt(class_profiles, id.vars = c("Cluster", "ID"))

# Use the 'ID' variable to sort the long-format data
class_profiles_long <- class_profiles_long[order(class_profiles_long$ID), ]

# Remove the 'ID' variable as it is no longer needed
class_profiles_long$ID <- NULL

# Create an 'ID' variable that records the order of the rows
class_profiles_long$ID <- seq_len(nrow(class_profiles_long))

# Merge with flat_mapping for class profiles
class_profiles_long <- merge(class_profiles_long, flat_mapping, 
                             by.x = c("variable", "value"), 
                             by.y = c("Column", "EncodedValue"),
                             all.x = TRUE)

# Merge and reapply ordering for class_response_frequencies
class_response_frequencies <- merge(class_response_frequencies, flat_mapping, 
                                    by.x = c("Question", "Response"), 
                                    by.y = c("Column", "EncodedValue"),
                                    all.x = TRUE)

class_response_frequencies <- class_response_frequencies %>%
  arrange(Cluster, match(Question, paste("Q", 1:36, "Combined", sep = "")), desc(Frequency))


# Merge and reapply ordering for class_response_proportions
class_response_proportions <- merge(class_response_proportions, flat_mapping, 
                                    by.x = c("Question", "Response"), 
                                    by.y = c("Column", "EncodedValue"),
                                    all.x = TRUE)

class_response_proportions <- class_response_proportions %>%
  arrange(Cluster, match(Question, paste("Q", 1:36, "Choice", sep = "")), desc(Prop))

# Use the 'ID' variable to sort the merged data
class_profiles_long <- class_profiles_long[order(class_profiles_long$ID), ]

# Remove the 'ID' variable as it is no longer needed
class_profiles_long$ID <- NULL

# Reshape the data frame to a wide format
class_profiles_wide <- reshape(class_profiles_long,
                               idvar = "Cluster",
                               timevar = "variable",
                               direction = "wide")

# Columns to keep
keep_cols <- c("Cluster", grep("OriginalValue", names(class_profiles_wide), value = TRUE))

# Subset the data frame
class_profiles_wide <- class_profiles_wide[, keep_cols]

# Rename the columns to reflect the original question labels
names(class_profiles_wide) <- gsub("OriginalValue", "", names(class_profiles_wide))

# Set row names to NULL
rownames(class_profiles_wide) <- NULL

# Remove the 'OriginalValue.' prefix from column names
names(class_profiles_wide) <- gsub("^OriginalValue.", "", names(class_profiles_wide))

##VISUALIZATIONS

### Bar plot for the count of members in each cluster

library(RColorBrewer)

ggplot(class_counts_df, aes(x = as.factor(Class), y = Count, fill = as.factor(Class))) +
  geom_bar(stat = "identity") +
  labs(title = "Cluster Sizes", x = "Cluster", y = "Count") +
  theme_minimal() +
  scale_fill_brewer() +
  guides(fill=guide_legend(title="Classes"))


str(filtered_data$Response)
filtered_data$Response <- as.factor(filtered_data$Response)
filtered_data$Cluster <- as.factor(filtered_data$Cluster)

### Proportional Bar Plots for Selected Questions

# Select a few key questions for illustration, say Q1Choice, Q2Choice, and Q3Choice
selected_questions <- c("Q1Choice", "Q2Choice", "Q3Choice")

# Filter the data for these questions
filtered_data <- class_response_proportions %>% filter(Question %in% selected_questions)

# Generate the bar plots
ggplot(filtered_data, aes(x = as.factor(Cluster), y = Prop, fill = Response)) +
  geom_bar(stat = "identity", position = "dodge") +
  facet_wrap(~ Question, ncol = 1) +
  labs(title = "Response Proportions for Selected Questions", x = "Cluster", y = "Proportion") +
  theme_minimal() +
  scale_fill_gradient(low = "blue", high = "red")

## EXPORT TO EXCEL ##

# Load the required package
library(writexl)

# Create a named list of data frames
dfs_to_export <- list(
  ClassCounts = class_counts_df,
  ClassProportions = class_proportions_df,
  ClassProfilesWide = class_profiles_wide,
  ClassResponseProportions = class_response_proportions,
  FitIndices = fit_indices
)

# Write the data frames to Excel
write_xlsx(dfs_to_export, path = "analysis_results_onlyChoice.xlsx")

