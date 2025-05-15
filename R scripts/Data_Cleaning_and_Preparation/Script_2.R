# ---- Data Loading ----
# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2025/ginlogicband")

# Read data from a CSV file
df_student <- read.csv("Student Uncoded.csv")
df_interviewer <- read.csv("Interviewer Uncoded.csv")


clean_dataframe <- function(df) {
  # Clean column names
  names(df) <- gsub(" ", "_", trimws(names(df)))
  names(df) <- gsub("\\s+", "_", names(df))
  names(df) <- gsub("\\(", "_", names(df))
  names(df) <- gsub("\\)", "_", names(df))
  names(df) <- gsub("-", "_", names(df))
  names(df) <- gsub("/", "_", names(df))
  names(df) <- gsub("\\\\", "_", names(df))
  names(df) <- gsub("\\?", "", names(df))
  names(df) <- gsub("'", "", names(df))
  names(df) <- gsub(",", "_", names(df))
  names(df) <- gsub("\\$", "", names(df))
  names(df) <- gsub("\\+", "_", names(df))
  
  # Clean character values
  df <- data.frame(lapply(df, function(x) {
    if (is.character(x)) {
      x <- tolower(trimws(x))
    }
    return(x)
  }), stringsAsFactors = FALSE)
  
  return(df)
}


df_student <- clean_dataframe(df_student)
df_interviewer <- clean_dataframe(df_interviewer)

str(df_student)
str(df_interviewer)


# FUNCTION 2: Recode Likert scale to numeric
recode_likert <- function(x) {
  recode_map <- c("strongly agree" = 5,
                  "agree" = 4,
                  "neither agree nor disagree" = 3,
                  "disagree" = 2,
                  "strongly disagree" = 1,
                  "yes" = 1,
                  "no" = 0)
  if (is.character(x)) {
    return(as.numeric(recode_map[x]))
  } else {
    return(x)
  }
}

# FUNCTION 3: Select only Likert variables
tidy_select_likert <- function(df) {
  # Remove columns with "please_explain" or "how_do_you_think" etc.
  exclude_patterns <- c("explain", "how", "improve")
  keep <- !grepl(paste(exclude_patterns, collapse = "|"), names(df))
  return(df[, keep])
}

# FUNCTION 4: Apply full cleaning pipeline
tidy_prepare_data <- function(df) {

  df <- tidy_select_likert(df)
  df[] <- lapply(df, recode_likert)
  return(df)
}


df_student_clean <- tidy_prepare_data(df_student)
df_interviewer_clean <- tidy_prepare_data(df_interviewer)

str(df_student_clean)
str(df_interviewer_clean)

describe_character_columns <- function(df) {
  char_vars <- names(df)[sapply(df, is.character) | sapply(df, is.factor)]
  for (var in char_vars) {
    cat("\nFrequency table for:", var, "\n")
    print(freq_table(df[[var]]))
  }
}


describe_character_columns(df_student)
describe_character_columns(df_interviewer)

mapping_table <- data.frame(
  student_var = c(
    "X2..I.am.satisfied.with.the.way.that.the.portfolio.interview.process.was.conducted.",
    "X4..The.portfolio.interview.process.and.the.questions.asked.were.relevant.to.my.personal.and.professional.development.",
    "X6..The.portfolio.interview.process.will.help.me.to.engage.more.in.the.portfolio.process.and.clinical.experiences.",
    "X8..The.portfolio.interview.process.has.encouraged.me.to.think.more.deeply.about.what.I.put.into.my.portfolio.reflections.",
    "X14..The.formality.of.the.portfolio.interview.process.made.me.feel.more.prepared.for.professional.environments.",
    "X18..The.portfolio.interview.process.is.a.beneficial.form.of.assessment.at.this.stage.in.my.career.",
    "X20..Do.you.think.that.the.portfolio.interview.should.be.a.standard.part.of.the.Year.4.assessment."
  ),
  interviewer_var = c(
    "X13..I.was.satisfied.with.the.portfolio.interview.process.and.its.impact.on.student.engagement.and.learning.",
    "X3..The.portfolio.interview.process.and.the.questions.asked.were.relevant.to.student.personal.and.professional.development.",
    "X5..The.portfolio.interview.process.is.a.useful.exercise.in.getting.students.to.engage.more.in.the.portfolio.process.overall.",
    "X7..The.portfolio.interview.process.is.a.useful.exercise.in.encouraging.students.to.think.more.deeply.about.what.they.put.into.their.portfolio.reflections.",
    "X9..The.formality.of.the.portfolio.interview.process.is.important.in.fostering.student.professionalism.",
    "X11..The.portfolio.interview.process.is.a.beneficial.form.of.assessment.at.this.stage.in.their.career.",
    "X15..Do.you.think.that.the.portfolio.interview.should.be.a.standard.part.of.the.Year.4.assessment."
  ),
  concept = c(
    "Satisfaction with process",
    "Relevance to development",
    "Engagement with portfolio",
    "Depth of reflection",
    "Professionalism fostered",
    "Beneficial assessment",
    "Should be part of assessment"
  ),
  stringsAsFactors = FALSE
)


# FUNCTION: Compute mean, median, SD for all numeric variables
descriptive_stats <- function(df) {
  numeric_vars <- names(df)[sapply(df, is.numeric)]
  result <- data.frame(
    Variable = numeric_vars,
    Mean = sapply(df[numeric_vars], mean, na.rm = TRUE),
    Median = sapply(df[numeric_vars], median, na.rm = TRUE),
    SD = sapply(df[numeric_vars], sd, na.rm = TRUE)
  )
  return(result)
}

# APPLY to student data
student_descriptives <- descriptive_stats(df_student_clean)

# View descriptives
student_descriptives

# FUNCTION: Calculate SEM
sem <- function(x) sd(x, na.rm = TRUE) / sqrt(sum(!is.na(x)))

# FUNCTION: Comparison + descriptives loop
compare_groups <- function(student_df, interviewer_df, mapping_table) {
  results <- data.frame()
  
  for (i in 1:nrow(mapping_table)) {
    s_var <- mapping_table$student_var[i]
    i_var <- mapping_table$interviewer_var[i]
    concept <- mapping_table$concept[i]
    
    student_scores <- student_df[[s_var]]
    interviewer_scores <- interviewer_df[[i_var]]
    
    # Descriptive stats
    s_mean <- mean(student_scores, na.rm = TRUE)
    s_median <- median(student_scores, na.rm = TRUE)
    s_sd <- sd(student_scores, na.rm = TRUE)
    s_sem <- sem(student_scores)
    
    i_mean <- mean(interviewer_scores, na.rm = TRUE)
    i_median <- median(interviewer_scores, na.rm = TRUE)
    i_sd <- sd(interviewer_scores, na.rm = TRUE)
    i_sem <- sem(interviewer_scores)
    
    # Mann-Whitney U test
    test <- wilcox.test(student_scores, interviewer_scores, exact = FALSE)
    
    # Store result
    row <- data.frame(
      Concept = concept,
      Student_Mean = s_mean,
      Student_Median = s_median,
      Student_SD = s_sd,
      Student_SEM = s_sem,
      Interviewer_Mean = i_mean,
      Interviewer_Median = i_median,
      Interviewer_SD = i_sd,
      Interviewer_SEM = i_sem,
      W_Statistic = test$statistic,
      p_Value = test$p.value
    )
    
    results <- rbind(results, row)
  }
  
  return(results)
}

# RUN the comparison
comparison_results <- compare_groups(df_student_clean, df_interviewer_clean, mapping_table)

# View results
comparison_results


# FUNCTION: Spearman correlations with tidy output
library(Hmisc) # for rcorr

compute_spearman_correlations <- function(df) {
  numeric_data <- df[, sapply(df, is.numeric)]
  corr <- rcorr(as.matrix(numeric_data), type = "spearman")
  
  # Get lower triangle of matrix
  corr_results <- data.frame()
  vars <- colnames(corr$r)
  for (i in 2:length(vars)) {
    for (j in 1:(i - 1)) {
      row <- data.frame(
        Var1 = vars[i],
        Var2 = vars[j],
        Rho = corr$r[i, j],
        p_value = corr$P[i, j]
      )
      corr_results <- rbind(corr_results, row)
    }
  }
  return(corr_results)
}

# RUN
student_correlations <- compute_spearman_correlations(df_student_clean)

# View
student_correlations


# FUNCTION: Replace long names with short codes
shorten_variable_names <- function(names_vector) {
  short_names <- names_vector
  short_names <- gsub("X1..Which.year.of.medical.school.are.you.currently.in.", "Year", short_names)
  short_names <- gsub("X2..I.am.satisfied.with.the.way.that.the.portfolio.interview.process.was.conducted.", "Satisfaction", short_names)
  short_names <- gsub("X4..The.portfolio.interview.process.and.the.questions.asked.were.relevant.to.my.personal.and.professional.development.", "Relevance", short_names)
  short_names <- gsub("X6..The.portfolio.interview.process.will.help.me.to.engage.more.in.the.portfolio.process.and.clinical.experiences.", "Engagement", short_names)
  short_names <- gsub("X8..The.portfolio.interview.process.has.encouraged.me.to.think.more.deeply.about.what.I.put.into.my.portfolio.reflections.", "Reflection", short_names)
  short_names <- gsub("X10..The.portfolio.interview.process.made.me.more.confident.in.identifying.my.learning.needs.and.goals.", "LearningNeeds", short_names)
  short_names <- gsub("X12..The.portfolio.interview.process.made.me.more.confident.in.presenting.my.experiences.and.goals.", "ConfidencePresenting", short_names)
  short_names <- gsub("X14..The.formality.of.the.portfolio.interview.process.made.me.feel.more.prepared.for.professional.environments.", "ProfessionalismPrepared", short_names)
  short_names <- gsub("X16..The.portfolio.interview.process.and.the.pre.interview.training.helped.to.increase.my.sense.of.professionalism.", "ProfessionalismTraining", short_names)
  short_names <- gsub("X18..The.portfolio.interview.process.is.a.beneficial.form.of.assessment.at.this.stage.in.my.career.", "BeneficialAssessment", short_names)
  short_names <- gsub("X20..Do.you.think.that.the.portfolio.interview.should.be.a.standard.part.of.the.Year.4.assessment.", "StandardAssessment", short_names)
  
  return(short_names)
}



# Heatmap of correlations
library(ggplot2)
library(reshape2)

plot_correlation_heatmap <- function(df) {
  numeric_data <- df[, sapply(df, is.numeric)]
  corr_matrix <- cor(numeric_data, method = "spearman", use = "pairwise.complete.obs")
  colnames(corr_matrix) <- shorten_variable_names(colnames(corr_matrix))
  rownames(corr_matrix) <- shorten_variable_names(rownames(corr_matrix))
  
  # Remove diagonal
  diag(corr_matrix) <- NA
  
  melted_corr <- melt(corr_matrix)
  
  ggplot(melted_corr, aes(Var1, Var2, fill = value)) +
    geom_tile(color = "white") +
    scale_fill_gradient2(low = "blue", high = "red", mid = "white", 
                         midpoint = 0, limit = c(-1,1), space = "Lab",
                         name = "Spearman\nCorrelation") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1),
          axis.title.x = element_blank(),
          axis.title.y = element_blank()) +
    coord_fixed()
}



# PLOT
plot_correlation_heatmap(df_student_clean)


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
  "Comparison" = comparison_results, 
  "Correlations" = student_correlations
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")

plot_student_boxplots <- function(df) {
  numeric_data <- df[, sapply(df, is.numeric)]
  
  # Rename columns for plotting
  original_names <- colnames(numeric_data)
  short_names <- shorten_variable_names(original_names)
  
  # Exclude unwanted variables
  exclude_vars <- c("StandardAssessment", "Year")
  keep_indices <- !short_names %in% exclude_vars
  
  numeric_data <- numeric_data[, keep_indices]
  colnames(numeric_data) <- short_names[keep_indices]
  
  # Reshape for plotting
  long_data <- pivot_longer(as.data.frame(numeric_data), 
                            cols = everything(), 
                            names_to = "Variable", 
                            values_to = "Score")
  
  # Plot
  ggplot(long_data, aes(x = Variable, y = Score)) +
    geom_boxplot(fill = "skyblue", color = "black") +
    theme_minimal() +
    labs(title = "Distribution of Student Ratings by Item", 
         x = "", 
         y = "Score (1-5)") +
    theme(axis.text.x = element_text(angle = 45, hjust = 1),
          plot.title = element_text(hjust = 0.5)) 
}

# Call the function
plot_student_boxplots(df_student_clean)
