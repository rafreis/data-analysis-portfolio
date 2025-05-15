# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/ama_vasquez")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("CleanedData.xlsx")


df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}), stringsAsFactors = FALSE)  # Ensure that character data does not convert to factors

library(dplyr)
library(tidyr)

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

# Filter the groups based on 'Only3Sessions'
group_no <- df %>% filter(Only3Sessions == "No")
group_yes <- df %>% filter(Only3Sessions == "Yes")


# Function to calculate descriptive statistics
calculate_descriptives <- function(data) {
  data %>%
    gather(key = "variable", value = "value", -Participant, -Only3Sessions) %>%
    group_by(variable) %>%
    summarise(
      mean = mean(value, na.rm = TRUE),
      median = median(value, na.rm = TRUE),
      sd = sd(value, na.rm = TRUE),
      range_min = min(value, na.rm = TRUE),
      range_max = max(value, na.rm = TRUE)
    ) %>%
    mutate(
      range = paste(range_min, "-", range_max)
    ) %>%
    select(-range_min, -range_max)
}

# Apply function to both groups
desc_stats_no <- calculate_descriptives(group_no)
desc_stats_yes <- calculate_descriptives(group_yes)



perform_wilcoxon <- function(data, pre_post_pairs) {
  results <- lapply(pre_post_pairs, function(pair) {
    pre <- data[[pair[1]]]
    post <- data[[pair[2]]]
    
    if (all(is.na(pre)) | all(is.na(post))) {
      return(data.frame(variable = pair[1], V = NA, p_value = NA, effect_size = NA))
    } else {
      test <- wilcox.test(pre, post, paired = TRUE, exact = FALSE, correct = TRUE)
      
      # Assuming large sample size and calculating Z
      N <- length(na.omit(pre - post))  # N should be the total number of observations possible for pairing
      Z <- (test$statistic - (N*(N+1)/4)) / sqrt((N*(N+1)*(2*N+1)/24))
      r = Z / sqrt(N)  # Correct effect size calculation
      
      data.frame(
        variable = pair[1],
        V = test$statistic,
        p_value = test$p.value,
        effect_size = r
      )
    }
  })
  return(do.call(rbind, results))
}




# Define pairs of variables for comparison
variable_pairs <- list(
  c("PreTotalMAIA", "PostTotalMAIA"),
  c("PreNoticing", "PostNoticing"),
  c("PreNoDistract", "PostNoDistract"),
  c("PreNotWorry", "PostNotWorry"),
  c("PreAttenReg", "PostAttenReg"),
  c("PreEmoAware", "PostEmoAware"),
  c("PreSelfReg", "PostSelfReg"),
  c("PreBodyListen", "PostBodyListen"),
  c("PreTrusting", "PostTrusting"),
  c("PreLEASTotal", "PostLEASTotal"),
  c("PreSelf", "PostSelf"),
  c("PreOther", "PostOther"),
  c("PreSBCTotal", "PostSBCTotal"),
  c("PreBodyAware", "PostBodyAware"),
  c("PreBodyDiss", "PostBodyDiss")
)


# Apply the function to both groups
test_results_no <- perform_wilcoxon(group_no, variable_pairs)
test_results_yes <- perform_wilcoxon(group_yes, variable_pairs)


# Boxplots


# Data preparation: melting the dataframe to long format for easier plotting
library(tidyr)
library(ggplot2)
library(dplyr)

# Reshape the dataframe and prepare for plotting
df_long <- df %>%
  pivot_longer(
    cols = c("PreTotalMAIA", "PostTotalMAIA", "PreNoticing", "PostNoticing",
             "PreNoDistract", "PostNoDistract", "PreNotWorry", "PostNotWorry",
             "PreAttenReg", "PostAttenReg", "PreEmoAware", "PostEmoAware",
             "PreSelfReg", "PostSelfReg", "PreBodyListen", "PostBodyListen",
             "PreTrusting", "PostTrusting", "PreSBCTotal", "PostSBCTotal",
             "PreBodyAware", "PostBodyAware", "PreBodyDiss", "PostBodyDiss",
             "PreLEASTotal", "PostLEASTotal", "PreSelf", "PostSelf",
             "PreOther", "PostOther"),
    names_to = "variable",
    values_to = "score"
  ) %>%
  mutate(
    Scale = case_when(
      variable %in% c("PreTotalMAIA", "PostTotalMAIA") ~ "Total MAIA",
      variable %in% c("PreNoticing", "PostNoticing",
                      "PreNoDistract", "PostNoDistract", "PreNotWorry", "PostNotWorry",
                      "PreAttenReg", "PostAttenReg", "PreEmoAware", "PostEmoAware",
                      "PreSelfReg", "PostSelfReg", "PreBodyListen", "PostBodyListen",
                      "PreTrusting", "PostTrusting") ~ "MAIA",
      variable %in% c("PreSBCTotal", "PostSBCTotal",
                      "PreBodyAware", "PostBodyAware", "PreBodyDiss", "PostBodyDiss") ~ "SBC",
      variable %in% c("PreLEASTotal", "PostLEASTotal", "PreSelf", "PostSelf",
                      "PreOther", "PostOther") ~ "LEAS"
    ),
    Period = ifelse(grepl("Pre", variable), "Pre", "Post"),
    Variable = sub("Pre|Post", "", variable)
  )

dodge_width <- 0.9  # Adjust this value as needed for optimal spacing

# Plot and save Total MAIA scores
total_maia_plot <- ggplot(df_long %>% filter(Scale == "Total MAIA"), aes(x = Variable, y = score, fill = Period)) +
  geom_boxplot(position = position_dodge(dodge_width)) +
  labs(title = "Boxplots for Total MAIA Scores", x = "Variable", y = "Score") +
  theme(axis.text.x = element_text(hjust = 1))
ggsave("Total_MAIA_Boxplots.png", plot = total_maia_plot, width = 10, height = 6, dpi = 300)

# Plot and save MAIA variables (excluding total MAIA)
maia_plot <- ggplot(df_long %>% filter(Scale == "MAIA"), aes(x = Variable, y = score, fill = Period)) +
  geom_boxplot(position = position_dodge(dodge_width)) +
  labs(title = "Boxplots for MAIA Variables (Excluding Total)", x = "Variable", y = "Score") +
  theme(axis.text.x = element_text(hjust = 1))
ggsave("MAIA_Boxplots.png", plot = maia_plot, width = 10, height = 6, dpi = 300)

# Plot and save LEAS variables
leas_plot <- ggplot(df_long %>% filter(Scale == "LEAS"), aes(x = Variable, y = score, fill = Period)) +
  geom_boxplot(position = position_dodge(dodge_width)) +
  labs(title = "Boxplots for LEAS Variables", x = "Variable", y = "Score") +
  theme(axis.text.x = element_text(hjust = 1))
ggsave("LEAS_Boxplots.png", plot = leas_plot, width = 10, height = 6, dpi = 300)

# Plot and save SBC variables
sbc_plot <- ggplot(df_long %>% filter(Scale == "SBC"), aes(x = Variable, y = score, fill = Period)) +
  geom_boxplot(position = position_dodge(dodge_width)) +
  labs(title = "Boxplots for SBC Variables", x = "Variable", y = "Score") +
  theme(axis.text.x = element_text(hjust = 1))
ggsave("SBC_Boxplots.png", plot = sbc_plot, width = 10, height = 6, dpi = 300)

# Display the plots
print(total_maia_plot)
print(maia_plot)
print(leas_plot)
print(sbc_plot)



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
  "Desc Stats" = desc_stats_no,
  "Desc Stats - 3sessions"  = desc_stats_yes, 
  "Wilcoxon" = test_results_no,
  "Wilcoxon - 3sessions" = test_results_yes
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")


library(tidyr)
library(ggplot2)
library(dplyr)

# Assuming df is your original dataframe
# Reshape the dataframe and prepare for plotting
df_long <- df %>%
  pivot_longer(
    cols = c("PreTotalMAIA", "PostTotalMAIA", "PreNoticing", "PostNoticing",
             "PreNoDistract", "PostNoDistract", "PreNotWorry", "PostNotWorry",
             "PreAttenReg", "PostAttenReg", "PreEmoAware", "PostEmoAware",
             "PreSelfReg", "PostSelfReg", "PreBodyListen", "PostBodyListen",
             "PreTrusting", "PostTrusting", "PreSBCTotal", "PostSBCTotal",
             "PreBodyAware", "PostBodyAware", "PreBodyDiss", "PostBodyDiss",
             "PreLEASTotal", "PostLEASTotal", "PreSelf", "PostSelf",
             "PreOther", "PostOther"),
    names_to = "variable",
    values_to = "score"
  ) %>%
  mutate(
    Scale = case_when(
      variable %in% c("PreTotalMAIA", "PostTotalMAIA") ~ "Total MAIA",
      variable %in% c("PreNoticing", "PostNoticing",
                      "PreNoDistract", "PostNoDistract", "PreNotWorry", "PostNotWorry",
                      "PreAttenReg", "PostAttenReg", "PreEmoAware", "PostEmoAware",
                      "PreSelfReg", "PostSelfReg", "PreBodyListen", "PostBodyListen",
                      "PreTrusting", "PostTrusting") ~ "MAIA",
      variable %in% c("PreSBCTotal", "PostSBCTotal",
                      "PreBodyAware", "PostBodyAware", "PreBodyDiss", "PostBodyDiss") ~ "SBC",
      variable %in% c("PreLEASTotal", "PostLEASTotal", "PreSelf", "PostSelf",
                      "PreOther", "PostOther") ~ "LEAS"
    ),
    Period = ifelse(grepl("Pre", variable), "Pre", "Post"),
    Variable = sub("Pre|Post", "", variable),
    Only3Sessions = recode(Only3Sessions, "No" = "All Sessions", "Yes" = "Only 3 Sessions")
  )


# Function to generate and save plots for a given session status
create_plots_for_session <- function(df_long, session_status) {
  # Filter for the specific session status
  session_data <- df_long %>% filter(Only3Sessions == session_status)
  
  # Generate and save boxplots for Total MAIA scores
  total_maia_plot <- ggplot(session_data %>% filter(Scale == "Total MAIA"), aes(x = Variable, y = score, fill = Period)) +
    geom_boxplot(position = position_dodge(0.9)) +
    labs(title = paste("Boxplots for Total MAIA Scores -", session_status), x = "Variable", y = "Score") +
    theme(axis.text.x = element_text(hjust = 1))
  ggsave(paste("Total_MAIA_Boxplots_", session_status, ".png", sep=""), plot = total_maia_plot, width = 10, height = 6, dpi = 300)
  
  # Generate and save boxplots for MAIA variables (excluding Total MAIA)
  maia_plot <- ggplot(session_data %>% filter(Scale == "MAIA"), aes(x = Variable, y = score, fill = Period)) +
    geom_boxplot(position = position_dodge(0.9)) +
    labs(title = paste("Boxplots for MAIA Variables -", session_status), x = "Variable", y = "Score") +
    theme(axis.text.x = element_text(hjust = 1))
  ggsave(paste("MAIA_Boxplots_", session_status, ".png", sep=""), plot = maia_plot, width = 10, height = 6, dpi = 300)
  
  # Generate and save boxplots for LEAS variables
  leas_plot <- ggplot(session_data %>% filter(Scale == "LEAS"), aes(x = Variable, y = score, fill = Period)) +
    geom_boxplot(position = position_dodge(0.9)) +
    labs(title = paste("Boxplots for LEAS Variables -", session_status), x = "Variable", y = "Score") +
    theme(axis.text.x = element_text(hjust = 1))
  ggsave(paste("LEAS_Boxplots_", session_status, ".png", sep=""), plot = leas_plot, width = 10, height = 6, dpi = 300)
  
  # Generate and save boxplots for SBC variables
  sbc_plot <- ggplot(session_data %>% filter(Scale == "SBC"), aes(x = Variable, y = score, fill = Period)) +
    geom_boxplot(position = position_dodge(0.9)) +
    labs(title = paste("Boxplots for SBC Variables -", session_status), x = "Variable", y = "Score") +
    theme(axis.text.x = element_text(hjust = 1))
  ggsave(paste("SBC_Boxplots_", session_status, ".png", sep=""), plot = sbc_plot, width = 10, height = 6, dpi = 300)
}

# Generate and save plots for 'No' and 'Yes'
create_plots_for_session(df_long, "All Sessions")
create_plots_for_session(df_long, "Only 3 Sessions")





# Create a dataset without filtering Only3Sessions
full_group <- df

# Define only the total score pairs
total_score_pairs <- list(
  c("PreTotalMAIA", "PostTotalMAIA"),
  c("PreSBCTotal", "PostSBCTotal"),
  c("PreLEASTotal", "PostLEASTotal")
)

# Calculate medians and Wilcoxon results
summary_total_scores <- lapply(total_score_pairs, function(pair) {
  pre <- full_group[[pair[1]]]
  post <- full_group[[pair[2]]]
  med_pre <- median(pre, na.rm = TRUE)
  med_post <- median(post, na.rm = TRUE)
  diff <- med_post - med_pre
  
  if (all(is.na(pre)) | all(is.na(post))) {
    return(data.frame(
      Variable = pair[1],
      Median_Pre = NA,
      Median_Post = NA,
      Difference = NA,
      V = NA,
      p_value = NA
    ))
  } else {
    test <- wilcox.test(pre, post, paired = TRUE, exact = FALSE, correct = TRUE)
    data.frame(
      Variable = gsub("Pre", "", pair[1]),
      Median_Pre = med_pre,
      Median_Post = med_post,
      Difference = diff,
      V = as.numeric(test$statistic),
      p_value = test$p.value
    )
  }
})

summary_total_scores_df <- do.call(rbind, summary_total_scores)

library(openxlsx)


# Save to Excel or print
write.xlsx(summary_total_scores_df, "Summary_TotalScores_Median_Wilcoxon.xlsx", overwrite = TRUE)
print(summary_total_scores_df)
