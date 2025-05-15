

library(openxlsx)
library(dplyr)
library(tidyr)

# Function to clean column names
clean_dataframe <- function(df) {
  # Clean column names
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
  
  # Trim and convert character values to lowercase
  df <- data.frame(lapply(df, function(x) {
    if (is.character(x)) {
      x <- trimws(x)  # Trim whitespace
      x <- tolower(x) # Convert to lowercase
    }
    return(x)
  }))
  
  return(df)
}

# Function to pivot dataset into wide format
pivot_pre_post <- function(df, id_cols, time_col) {
  df %>%
    pivot_wider(
      names_from = !!sym(time_col),   # Use the Time column (Pre/Post) to create wide format
      values_from = -all_of(c(id_cols, time_col)), # All other columns are pivoted
      names_sep = "_"
    )
}

# Load the Excel file
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/lemonroze")
sheet_names <- getSheetNames("Clean_Data.xlsx")

# Read and clean data
dfs <- list()
for (sheet in sheet_names) {
  dfs[[sheet]] <- clean_dataframe(read.xlsx("Clean_Data.xlsx", sheet = sheet))
}

# Filter datasets
dfs <- dfs[!grepl("Revised", names(dfs)) | grepl("PANAS Pre \\(Revised\\)|PANAS Post \\(Revised\\)", names(dfs))]
dfs <- dfs[!grepl("PANAS Pre$|PANAS Post$", names(dfs))]

# Integrate datasets with Time column
df_SWC <- rbind(cbind(dfs$SWCPre, Time = "Pre"), cbind(dfs$`SWC Post`, Time = "Post"))
df_SD_Teacher <- rbind(cbind(dfs$`S&DPre teacher`, Time = "Pre"), cbind(dfs$`S&DPost teacher`, Time = "Post"))
df_SD_Parent <- rbind(cbind(dfs$`S&D Pre Parent`, Time = "Pre"), cbind(dfs$`S&D Post Parent`, Time = "Post"))
df_PANAS <- rbind(cbind(dfs$`PANAS Pre (Revised)`, Time = "Pre"), cbind(dfs$`PANAS Post (Revised)`, Time = "Post"))

# Pivot datasets to wide format
id_cols <- c("Ref", "SN")  # Define identifier columns
df_SWC_wide <- pivot_pre_post(df_SWC, id_cols, time_col = "Time")
df_SD_Teacher_wide <- pivot_pre_post(df_SD_Teacher, id_cols, time_col = "Time")
df_SD_Parent_wide <- pivot_pre_post(df_SD_Parent, id_cols, time_col = "Time")
df_PANAS_wide <- pivot_pre_post(df_PANAS, id_cols, time_col = "Time")



library(psych)


reliability_analysis <- function(data, scales, method = "sum") {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Time = character(),
    Scale = character(),
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Create a copy of the data to store the results
  data_with_scales <- data
  
  # Process each scale
  for (scale in names(scales)) {
    for (time in c("Pre", "Post")) {
      # Identify columns for the specific time (e.g., *_Pre, *_Post)
      time_cols <- paste0(scales[[scale]], "_", time)
      
      # Check if the columns exist in the dataframe
      if (!all(time_cols %in% colnames(data))) {
        warning(paste("Some columns for scale", scale, "and time", time, "are missing. Skipping."))
        next
      }
      
      # Extract the subset of columns for the time point
      subset_data <- data[time_cols]
      
      # Remove items with no variance
      no_variance_items <- sapply(subset_data, function(x) var(x, na.rm = TRUE) == 0)
      subset_data <- subset_data[, !no_variance_items, drop = FALSE]
      
      # Skip scale if all items have no variance
      if (ncol(subset_data) == 0) {
        warning(paste("All items in scale", scale, "at time", time, "have no variance. Skipping."))
        next
      }
      
      # Calculate Cronbach's Alpha
      alpha_results <- psych::alpha(subset_data)
      alpha_val <- alpha_results$total$raw_alpha
      
      # Calculate descriptive statistics for each item
      for (item in colnames(subset_data)) {
        item_data <- subset_data[[item]]
        results <- rbind(results, data.frame(
          Time = time,
          Scale = scale,
          Variable = item,
          Mean = mean(item_data, na.rm = TRUE),
          SEM = sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data))),
          StDev = sd(item_data, na.rm = TRUE),
          ITC = alpha_results$item.stats[colnames(subset_data) == item, "raw.r"],
          Alpha = NA
        ))
      }
      
      # Calculate the total score for the scale
      if (method == "sum") {
        scale_total <- rowSums(subset_data, na.rm = TRUE)
      } else if (method == "average") {
        scale_total <- rowMeans(subset_data, na.rm = TRUE)
      } else {
        stop("Invalid method. Use 'sum' or 'average'.")
      }
      
      # Add the total score for the scale and time point
      data_with_scales[[paste0(scale, "_Total_", time)]] <- scale_total
      
      # Add total score statistics to the results
      results <- rbind(results, data.frame(
        Time = time,
        Scale = scale,
        Variable = paste0(scale, "_Total"),
        Mean = mean(scale_total, na.rm = TRUE),
        SEM = sd(scale_total, na.rm = TRUE) / sqrt(sum(!is.na(scale_total))),
        StDev = sd(scale_total, na.rm = TRUE),
        ITC = NA,
        Alpha = alpha_val
      ))
    }
  }
  
  return(list(data_with_scales = data_with_scales, statistics = results))
}





# Define scales for all datasets
scales_PANAS <- list(
  "Positive Affect" = c("joyful", "cheerful", "happy", "lively", "proud", "energentic", "delighted", "excited", "active", "strong", "calm", "interested"),
  "Negative Affect" = c("scared", "sad", "mad", "gloomy", "upset", "lonely", "shamed", "frightened", "disgusted", "guilty", "nervous", "jittery")
)

scales_SD_Parent <- list(
  "Prosocial Behavior" = c("consid", "shares", "caring", "kind", "helpout"),  
  "Hyperactivity" = c("restles", "fidgety", "distrac", "attends.", "reflect."),           
  "Emotional Symptoms" = c("worries", "unhappy", "afraid", "clingy", "somatic"),      
  "Peer Problems" = c("loner", "friend..", "popular.", "bullied", "oldbest"),  
  "Conduct Problems" = c("tantrum", "fights", "lies", "steals", "obeys.."),
  "Total Difficulties" = c(
    "restles", "fidgety", "distrac", "attends.", "reflect.",  # Hyperactivity
    "worries", "unhappy", "afraid", "clingy", "somatic",      # Emotional Symptoms
    "loner", "friend..", "popular.", "bullied", "oldbest",    # Peer Problems
    "tantrum", "fights", "lies", "steals", "obeys.."          # Conduct Problems
  )
)
colnames(df_SD_Parent)
scales_SD_Teacher <- scales_SD_Parent  # Same structure as SD_Parent

scales_SWC <- list(
  "Sense of Worth" = c("SOW1", "SOW2", "SOW3"),
  "Competence Development" = c("CD1", "CD2", "CD3"),
  "Support System" = c("SOS1", "SOS2", "SOS3"),
  "Sense of Purpose" = c("SOP1", "SOP2", "SOP3", "SOP4", "SOP5", "SOP6"),
  "Autonomy" = c("A1", "A2", "A3")
)

# Perform reliability analysis for each dataset
result_PANAS <- reliability_analysis(df_PANAS_wide, scales_PANAS, method = "sum")
result_SD_Parent <- reliability_analysis(df_SD_Parent_wide, scales_SD_Parent, method = "sum")
result_SD_Teacher <- reliability_analysis(df_SD_Teacher_wide, scales_SD_Teacher, method = "sum")
result_SWC <- reliability_analysis(df_SWC_wide, scales_SWC, method = "sum")

# Extract reliability statistics
statistics_PANAS <- result_PANAS$statistics
statistics_SD_Parent <- result_SD_Parent$statistics
statistics_SD_Teacher <- result_SD_Teacher$statistics
statistics_SWC <- result_SWC$statistics

# Extract the updated dataframes with summed/averaged scale scores
df_PANAS_with_scales <- result_PANAS$data_with_scales
df_SD_Parent_with_scales <- result_SD_Parent$data_with_scales
df_SD_Teacher_with_scales <- result_SD_Teacher$data_with_scales
df_SWC_with_scales <- result_SWC$data_with_scales


# Define output file names
output_files <- list(
  PANAS = "PANAS_with_Scales.csv",
  SD_Parent = "SD_Parent_with_Scales.csv",
  SD_Teacher = "SD_Teacher_with_Scales.csv",
  SWC = "SWC_with_Scales.csv"
)

# Export each dataframe to CSV
write.csv(df_PANAS_with_scales, output_files$PANAS, row.names = FALSE)
write.csv(df_SD_Parent_with_scales, output_files$SD_Parent, row.names = FALSE)
write.csv(df_SD_Teacher_with_scales, output_files$SD_Teacher, row.names = FALSE)
write.csv(df_SWC_with_scales, output_files$SWC, row.names = FALSE)

# Print confirmation
cat("Dataframes exported to CSV files:\n")
print(output_files)

