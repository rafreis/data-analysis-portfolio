library(ggplot2)
library(dplyr)
library(tidyr)
library(stringr)
library(haven)

# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/mlanjab2023")

# Import the SPSS file
df <- read_sav("Job_Search_Platform_Survey_May_15__2024_10.20 (1).sav")

# List of labelled columns to convert to factors
labelled_columns <- c("Q1.2", "Q2.1", "Q2.3", "Q3.2", "Q3.5", 
                      "Q3.7_1", "Q3.7_2", "Q3.7_3", "Q3.7_4", "Q3.7_5", 
                      "Q3.7_6", "Q3.7_7", "Q3.7_8", "Q3.7_9", "Q3.7_10", 
                      "Q3.7_11", "Q3.7_12", "Q3.7_14", "Q4.3_1", "Q4.3_2", 
                      "Q4.3_3", "Q4.3_5", "Q5.1", "Q5.4", "Q5.6", "Q5.8", 
                      "Q5.9", "Q5.11_1", "Q5.11_2", "Q5.11_3", "Q5.11_4", 
                      "Q5.11_5", "Q5.11_6", "Q5.11_7", "Q5.11_8", "Q5.11_9", 
                      "Q5.11_10", "Q5.11_11", "Q5.11_12", "Q5.11_14", "Q5.11_13", 
                      "Q6.1", "Q6.3")

# List of numeric columns to convert to numeric
numeric_columns <- c("Q2.4_1", "Q2.4_2", "Q2.4_3", "Q2.4_4", "Q2.4_5", 
                     "Q2.5_1", "Q2.5_2", "Q2.5_3", "Q2.5_4", "Q2.5_5", 
                     "Q3.3_1", "Q3.3_2", "Q3.3_3", "Q3.3_4", "Q3.4_1", 
                     "Q3.4_2", "Q3.4_3", "Q3.4_4", "Q3.4_5", "Q3.6", 
                     "sum_q3.7", "Q3.8_1", "Q3.8_2", "Q3.8_3", "Q3.8_4", 
                     "Q3.8_5", "Q3.8_6", "Q3.8_7", "Q3.8_8", "Q3.8_9", 
                     "Q3.8_10", "Q3.8_11", "Q3.8_12", "Q4.1_1", "Q4.1_2", 
                     "Q4.1_3", "Q4.1_4", "Q4.1_5", "Q4.1_6", "Q4.2_1", 
                     "Q4.2_2", "Q4.2_3", "Q4.2_4", "Q4.2_5", "Q4.2_6", 
                     "Q4.4_1", "Q4.4_2", "Q4.4_3", "Q4.4_4", "Q4.4_5", 
                     "Q4.4_6", "Q4.5_1", "Q4.5_2", "Q4.5_3", "Q4.5_4", 
                     "Q4.5_5", "Q4.5_6", "Q4.6_1", "Q4.6_2", "Q4.6_3", 
                     "Q4.6_4", "Q4.6_5", "Q4.6_6", "Q5.2_1", "Q5.2_2", 
                     "Q5.2_3", "Q5.2_4", "Q5.3_1", "Q5.3_2", "Q5.3_3", 
                     "Q5.3_4", "Q5.3_5", "Q5.5_1", "Q5.5_2", "Q5.5_3", 
                     "Q5.5_4", "Q5.5_5", "Q5.5_6", "Q5.10_1", "Q5.10_2", 
                     "Q5.10_3", "Q5.10_4", "Q5.10_5", "Q5.10_6", "Q5.10_7", 
                     "Q5.10_8", "Q5.10_9", "Q5.10_10", "Q5.10_11", "Q5.10_12", 
                     "Q5.10_13", "Q5.10_14", "sum_q5.11", "Q5.12", "Q5.13")

# Convert labelled columns to factors and numeric columns to numeric
df <- df %>%
  mutate(across(all_of(labelled_columns), ~ as_factor(.))) %>%
  mutate(across(all_of(numeric_columns), as.numeric))

# List of columns to invert scales
invert_columns <- c("Q4.4_1", "Q4.4_2", "Q4.4_3", "Q4.4_4", "Q4.4_5", "Q4.4_6",
                    "Q4.5_1", "Q4.5_2", "Q4.5_3", "Q4.5_4", "Q4.5_5", "Q4.5_6",
                    "Q4.6_1", "Q4.6_2", "Q4.6_3", "Q4.6_4", "Q4.6_5", "Q4.6_6")

# Function to invert scale 1-5 to 5-1
invert_scale <- function(x) {
  if (is.numeric(x)) {
    return(6 - x)
  } else {
    return(x)
  }
}

# Invert scales for specified columns
df <- df %>%
  mutate(across(all_of(invert_columns), invert_scale))

# Extract question labels from the provided list and create a named vector
labels_list <- c(
  "Q2.4_1" = "How useful are these services in helping you find an internship? - Online job platforms (e.g., LinkedIn, Indeed)",
  "Q2.4_2" = "How useful are these services in helping you find an internship? - Career services provided by the college",
  "Q2.4_3" = "How useful are these services in helping you find an internship? - Personal connections or networking",
  "Q2.4_4" = "How useful are these services in helping you find an internship? - Finding alumni connections",
  "Q2.4_5" = "How useful are these services in helping you find an internship? - Other (Please specify)",
  "Q2.5_1" = "How challenging were the following during your internship search process? - Lack of response from employers",
  "Q2.5_2" = "How challenging were the following during your internship search process? - Difficulty finding relevant opportunities",
  "Q2.5_3" = "How challenging were the following during your internship search process? - Not having the proper experience for a job",
  "Q2.5_4" = "How challenging were the following during your internship search process? - Finding alumni connections",
  "Q2.5_5" = "How challenging were the following during your internship search process? - Other (Please specify)",
  "Q3.3_1" = "How important are these goals for you when you are pursuing a job? - Financial independence",
  "Q3.3_2" = "How important are these goals for you when you are pursuing a job? - Gaining work experience",
  "Q3.3_3" = "How important are these goals for you when you are pursuing a job? - Networking opportunities",
  "Q3.3_4" = "How important are these goals for you when you are pursuing a job? - Other (Please specify)",
  "Q3.4_1" = "How useful are these services in helping you find a job? - Online platforms (e.g., LinkedIn, Indeed)",
  "Q3.4_2" = "How useful are these services in helping you find a job? - Career services provided by the college",
  "Q3.4_3" = "How useful are these services in helping you find a job? - Personal connections or networking",
  "Q3.4_4" = "How useful are these services in helping you find a job? - Alumni connections",
  "Q3.4_5" = "How useful are these services in helping you find a job? - Other (Please specify)",
  "Q3.6" = "How helpful were these resources in aligning with your career goals and aspirations?",
  "sum_q3.7" = "Number of Interests (Part-Time Jobs)",
  "Q3.8_1" = "How concerned are you about the following challenges when looking for jobs? - Your lack of experience",
  "Q3.8_2" = "How concerned are you about the following challenges when looking for jobs? - Time commitment",
  "Q3.8_3" = "How concerned are you about the following challenges when looking for jobs? - Limited Network and Connections",
  "Q3.8_4" = "How concerned are you about the following challenges when looking for jobs? - Navigating Job Search Platforms",
  "Q3.8_5" = "How concerned are you about the following challenges when looking for jobs? - Lack of Career Guidance",
  "Q3.8_6" = "How concerned are you about the following challenges when looking for jobs? - Uncertainty About Career Interests",
  "Q3.8_7" = "How concerned are you about the following challenges when looking for jobs? - Competitive Job Market",
  "Q3.8_8" = "How concerned are you about the following challenges when looking for jobs? - Transportation and Location Issues",
  "Q3.8_9" = "How concerned are you about the following challenges when looking for jobs? - Finding Quality Opportunities",
  "Q3.8_10" = "How concerned are you about the following challenges when looking for jobs? - Mostly Unpaid Internships",
  "Q3.8_11" = "How concerned are you about the following challenges when looking for jobs? - Low pay for entry level jobs",
  "Q3.8_12" = "How concerned are you about the following challenges when looking for jobs? - Lack of professional skills",
  "Q4.1_1" = "How familiar are you with the following job-searching methods? - Networking",
  "Q4.1_2" = "How familiar are you with the following job-searching methods? - Online job platforms (e.g., LinkedIn, Indeed, Handshake)",
  "Q4.1_3" = "How familiar are you with the following job-searching methods? - Company websites",
  "Q4.1_4" = "How familiar are you with the following job-searching methods? - Career services offered by your school",
  "Q4.1_5" = "How familiar are you with the following job-searching methods? - Social media",
  "Q4.1_6" = "How familiar are you with the following job-searching methods? - Referrals from friends or family",
  "Q4.2_1" = "How useful do you think the following features would be in a job application app? - Education section",
  "Q4.2_2" = "How useful do you think the following features would be in a job application app? - A place to post comments",
  "Q4.2_3" = "How useful do you think the following features would be in a job application app? - Skills Assessment",
  "Q4.2_4" = "How useful do you think the following features would be in a job application app? - Advanced search",
  "Q4.2_5" = "How useful do you think the following features would be in a job application app? - Mentorship Program",
  "Q4.2_6" = "How useful do you think the following features would be in a job application app? - Viewing job postings",
  "Q4.4_1" = "Rate each of the following features of LinkedIn. - User-Friendly Interface",
  "Q4.4_2" = "Rate each of the following features of LinkedIn. - Job Recommendation Accuracy",
  "Q4.4_3" = "Rate each of the following features of LinkedIn. - Advanced Search Filters",
  "Q4.4_4" = "Rate each of the following features of LinkedIn. - Company Insights and Reviews",
  "Q4.4_5" = "Rate each of the following features of LinkedIn. - Networking Opportunities",
  "Q4.4_6" = "Rate each of the following features of LinkedIn. - Education aspects",
  "Q4.5_1" = "Rate each of the following features of Indeed. - User-Friendly Interface",
  "Q4.5_2" = "Rate each of the following features of Indeed. - Job Recommendation Accuracy",
  "Q4.5_3" = "Rate each of the following features of Indeed. - Advanced Search Filters",
  "Q4.5_4" = "Rate each of the following features of Indeed. - Company Insights and Reviews",
  "Q4.5_5" = "Rate each of the following features of Indeed. - Networking Opportunities",
  "Q4.5_6" = "Rate each of the following features of Indeed. - Education aspects",
  "Q4.6_1" = "Rate each of the following features of Handshake. - User-Friendly Interface",
  "Q4.6_2" = "Rate each of the following features of Handshake. - Job Recommendation Accuracy",
  "Q4.6_3" = "Rate each of the following features of Handshake. - Advanced Search Filters",
  "Q4.6_4" = "Rate each of the following features of Handshake. - Company Insights and Reviews",
  "Q4.6_5" = "Rate each of the following features of Handshake. - Networking Opportunities",
  "Q4.6_6" = "Rate each of the following features of Handshake. - Education aspects",
  "Q5.2_1" = "How important are these goals for you when you are pursuing a job? - Financial independence",
  "Q5.2_2" = "How important are these goals for you when you are pursuing a job? - Gaining work experience",
  "Q5.2_3" = "How important are these goals for you when you are pursuing a job? - Networking opportunities",
  "Q5.2_4" = "How important are these goals for you when you are pursuing a job? - Other (Please specify)",
  "Q5.3_1" = "How useful are the following methods to find part-time/full-time job opportunities? - Online platforms (e.g., LinkedIn, Indeed)",
  "Q5.3_2" = "How useful are the following methods to find part-time/full-time job opportunities? - Word of mouth (friends, family, acquaintances)",
  "Q5.3_3" = "How useful are the following methods to find part-time/full-time job opportunities? - Local businesses or establishments",
  "Q5.3_4" = "How useful are the following methods to find part-time/full-time job opportunities? - School resources",
  "Q5.3_5" = "How useful are the following methods to find part-time/full-time job opportunities? - Other (Please specify)",
  "Q5.5_1" = "How important are the following factors when you're picking a college? - Academic programs offered",
  "Q5.5_2" = "How important are the following factors when you're picking a college? - Location",
  "Q5.5_3" = "How important are the following factors when you're picking a college? - Cost of tuition and fees",
  "Q5.5_4" = "How important are the following factors when you're picking a college? - Reputation/prestige",
  "Q5.5_5" = "How important are the following factors when you're picking a college? - Campus culture",
  "Q5.5_6" = "How important are the following factors when you're picking a college? - Other (Please specify)",
  "Q5.10_1" = "How concerned are you about the following challenges when looking for jobs? - Your lack of experience",
  "Q5.10_2" = "How concerned are you about the following challenges when looking for jobs? - Time commitment",
  "Q5.10_3" = "How concerned are you about the following challenges when looking for jobs? - Limited Network and Connections",
  "Q5.10_4" = "How concerned are you about the following challenges when looking for jobs? - Navigating Job Search Platforms",
  "Q5.10_5" = "How concerned are you about the following challenges when looking for jobs? - Lack of Career Guidance",
  "Q5.10_6" = "How concerned are you about the following challenges when looking for jobs? - Uncertainty About Career Interests",
  "Q5.10_7" = "How concerned are you about the following challenges when looking for jobs? - Competitive Job Market",
  "Q5.10_8" = "How concerned are you about the following challenges when looking for jobs? - Transportation and Location Issues",
  "Q5.10_9" = "How concerned are you about the following challenges when looking for jobs? - Finding Quality Opportunities",
  "Q5.10_10" = "How concerned are you about the following challenges when looking for jobs? - Unsure where to look for jobs",
  "Q5.10_11" = "How concerned are you about the following challenges when looking for jobs? - Mostly Unpaid Internships",
  "Q5.10_12" = "How concerned are you about the following challenges when looking for jobs? - Lack of professional skills",
  "Q5.10_13" = "How concerned are you about the following challenges when looking for jobs? - Lack of response from employers",
  "Q5.10_14" = "How concerned are you about the following challenges when looking for jobs? - Other (Please specify)",
  "Q5.11_1" = "What types of Part-Time Jobs are you interested in? - Cashier positions",
  "Q5.11_2" = "What types of Part-Time Jobs are you interested in? - Sales associate in Local store or mall",
  "Q5.11_3" = "What types of Part-Time Jobs are you interested in? - Barista at a coffee shop",
  "Q5.11_4" = "What types of Part-Time Jobs are you interested in? - Fast food restaurant roles",
  "Q5.11_5" = "What types of Part-Time Jobs are you interested in? - Waiter/Busboy",
  "Q5.11_6" = "What types of Part-Time Jobs are you interested in? - Host/Hostess at restaurant",
  "Q5.11_7" = "What types of Part-Time Jobs are you interested in? - Tutoring",
  "Q5.11_8" = "What types of Part-Time Jobs are you interested in? - Office Position",
  "Q5.11_9" = "What types of Part-Time Jobs are you interested in? - Delivery Jobs",
  "Q5.11_10" = "What types of Part-Time Jobs are you interested in? - Baby sitting",
  "Q5.11_11" = "What types of Part-Time Jobs are you interested in? - Pet care",
  "Q5.11_12" = "What types of Part-Time Jobs are you interested in? - Internships",
  "Q5.11_14" = "What types of Part-Time Jobs are you interested in? - Other",
  "Q5.11_13" = "What types of Part-Time Jobs are you interested in? - I am not looking for a part time job at this time",
  "sum_q5.11" = "Number of Interests (Part-Time Jobs)",
  "Q5.12" = "How likely are you to use a job application app recommended or integrated by your high school?",
  "Q5.13" = "How interested would you be in taking a free 15 minute course offered by an app to get certified and make yourself better for jobs/internships?"
)

# CHarts
library(ggplot2)
library(dplyr)
library(tidyr)

# Function to create horizontal bar plots for numeric variables
create_bar_plot <- function(data, var_names, title_prefix) {
  data_long <- data %>%
    select(all_of(var_names)) %>%
    pivot_longer(cols = everything(), names_to = "Question", values_to = "Score") %>%
    mutate(Question = factor(Question, levels = var_names))
  
  mean_scores <- data_long %>%
    group_by(Question) %>%
    summarise(Mean_Score = mean(Score, na.rm = TRUE), Count = n())
  
  plot <- ggplot(mean_scores, aes(x = Question, y = Mean_Score)) +
    geom_bar(stat = "identity") +
    geom_point(aes(y = Mean_Score), color = "red", size = 3) +
    geom_text(aes(y = Mean_Score, label = paste0(round(Mean_Score, 2), " (", Count, ")")), 
              hjust = -0.3, vjust = 0.5) +
    scale_x_discrete(labels = labels_list[var_names]) +
    scale_y_continuous(limits = c(0, 5)) +
    coord_flip() +
    ggtitle(paste("Mean Scores of", title_prefix)) +
    xlab("Questions") +
    ylab("Mean Score") +
    theme_bw() +
    theme(axis.text.x = element_text(angle = 90, hjust = 1), 
          plot.background = element_rect(fill = "white", color = NA),
          panel.background = element_rect(fill = "white", color = NA),
          plot.margin = unit(c(1, 2, 1, 1), "cm"))
  
  return(plot)
}

library(ggrepel)

# Function to create pie charts for labelled variables
# Extract variable labels
var_labels <- sapply(df, function(x) attr(x, "label"))

# Function to create pie charts for labelled variables
create_pie_chart <- function(data, var_name, var_label) {
  if (var_name %in% names(data)) {
    data_filtered <- data %>%
      filter(!is.na(.data[[var_name]]))
    
    if (nrow(data_filtered) > 0) {
      data_filtered <- data_filtered %>%
        count(.data[[var_name]]) %>%
        mutate(Percentage = n / sum(n) * 100)
      
      plot <- ggplot(data_filtered, aes(x = "", y = Percentage, fill = .data[[var_name]])) +
        geom_bar(stat = "identity", width = 1) +
        coord_polar("y") +
        geom_label_repel(aes(label = paste0(round(Percentage, 1), "% (", n, ")")), 
                         position = position_stack(vjust = 0.5),
                         show.legend = FALSE,
                         segment.size = 0.2) +
        ggtitle(var_label) +
        ylab(var_name) +
        theme_bw() +  # Set white background
        theme_void() +  # Remove the axis
        theme(legend.position = "right",
              plot.background = element_rect(fill = "white", color = NA),
              panel.background = element_rect(fill = "white", color = NA))
      return(plot)
    } else {
      message(paste("Skipping variable", var_name, "due to no valid data."))
      return(NULL)
    }
  } else {
    message(paste("Variable", var_name, "not found in the dataset."))
    return(NULL)
  }
}


# Function to create percentage bar plots for Q3 and Q5
create_percentage_bar_plot <- function(data, vars, title) {
  data_long <- data %>%
    select(all_of(vars)) %>%
    pivot_longer(cols = everything(), names_to = "Question", values_to = "Response") %>%
    filter(!is.na(Response)) %>%  # Keep NA values out of the percentage calculation
    group_by(Question, Response) %>%
    summarise(Count = n(), .groups = "drop") %>%
    group_by(Question) %>%
    mutate(Percentage = Count / sum(Count) * 100) %>%
    filter(Response != "0")  # Omit "0" responses
  
  data_long <- data_long %>%
    mutate(Response = factor(Response, levels = unique(data_long$Response[order(data_long$Percentage, decreasing = TRUE)])))
  
  plot <- ggplot(data_long, aes(x = Percentage, y = Question, fill = Response)) +
    geom_col(stat = "identity", position = "dodge") +
    geom_text(aes(label = paste0(sprintf("%.1f%%", Percentage), " (", Count, ")")), 
              position = position_dodge(width = 0.9), vjust = -0.5) +
    coord_flip() +
    ggtitle(title) +
    xlab("Percentage") +
    ylab("Question") +
    theme_bw() +
    theme(plot.background = element_rect(fill = "white", color = NA),
          panel.background = element_rect(fill = "white", color = NA))
  
  return(plot)
}


numeric_vars_q3 <- c("Q3.3_1", "Q3.3_2", "Q3.3_3",  "Q3.4_1", 
                     "Q3.4_2", "Q3.4_3", "Q3.4_4",  "Q3.6", "sum_q3.7") 
                     
                     
numeric_vars_q32 <- c("Q3.8_1", "Q3.8_2", "Q3.8_3", "Q3.8_4", "Q3.8_5")

numeric_vars_q33 <- c("Q3.8_6", 
                     "Q3.8_7", "Q3.8_8", "Q3.8_9", "Q3.8_10", "Q3.8_11", "Q3.8_12")

numeric_vars_q4 <- c("Q4.1_1", "Q4.1_2", "Q4.1_3", "Q4.1_4", "Q4.1_5", 
                     "Q4.1_6", "Q4.2_1", "Q4.2_2", "Q4.2_3")
                     
numeric_vars_q42 <-  c("Q4.2_4", "Q4.2_5", 
                     "Q4.2_6", "Q4.4_1", "Q4.4_2", "Q4.4_3", "Q4.4_4", "Q4.4_5", 
                     "Q4.4_6")

numeric_vars_q43 <-   c("Q4.5_1", "Q4.5_2", "Q4.5_3", "Q4.5_4", "Q4.5_5", 
                     "Q4.5_6")

numeric_vars_q44 <-  c("Q4.6_1", "Q4.6_2", "Q4.6_3", "Q4.6_4", "Q4.6_5", 
                     "Q4.6_6")

numeric_vars_q5 <- c("Q5.2_1", "Q5.2_2", "Q5.2_3",  "Q5.3_1", 
                     "Q5.3_2", "Q5.3_3", "Q5.3_4",  "Q5.5_1", "Q5.5_2", 
                     "Q5.5_3", "Q5.5_4", "Q5.5_5")


numeric_vars_q52 <- c( "Q5.10_1", "Q5.10_2", 
"Q5.10_3", "Q5.10_4", "Q5.10_5", "Q5.10_6", "Q5.10_7", 
"Q5.10_8")

numeric_vars_q53 <- c("Q5.10_9", "Q5.10_10", "Q5.10_11", "Q5.10_12", 
"Q5.10_13",  "Q5.12", "Q5.13")



# Export

# Example of saving plots with adjusted width
save_plot <- function(plot, filename) {
  ggsave(filename = filename, plot = plot, path = getwd(), width = 14, height = 8, units = "in")
}


# Loop to create pie charts for labelled variables and save them
for (var in labelled_columns) {
  var_label <- var_labels[var]
  plot <- create_pie_chart(df, var, var_label)
  if (!is.null(plot)) {
    print(plot)
    save_plot(plot, paste0(var, "_pie_chart.png"))
  }
}


# Create bar plots for numeric variables in sets
plot_q3 <- create_bar_plot(df, numeric_vars_q3, "Q3")
print(plot_q3)
save_plot(plot_q3, "Q3_bar_plot.png")

# Create bar plots for numeric variables in sets
plot_q32 <- create_bar_plot(df, numeric_vars_q32, "Q3")
print(plot_q32)
save_plot(plot_q32, "Q3_bar_plot2.png")

plot_q33 <- create_bar_plot(df,numeric_vars_q33, "Q3")
print(plot_q33)
save_plot(plot_q33, "Q3_bar_plot3.png")

plot_q4 <- create_bar_plot(df, numeric_vars_q4, "Q4")
print(plot_q4)
save_plot(plot_q4, "Q4_bar_plot.png")

plot_q42 <- create_bar_plot(df, numeric_vars_q42, "Q4")
print(plot_q42)
save_plot(plot_q42, "Q4_bar_plot2.png")

plot_q43 <- create_bar_plot(df, numeric_vars_q43, "Q4")
print(plot_q43)
save_plot(plot_q43, "Q4_bar_plot3.png")

plot_q44 <- create_bar_plot(df, numeric_vars_q44, "Q4")
print(plot_q44)
save_plot(plot_q44, "Q4_bar_plot4.png")

plot_q5 <- create_bar_plot(df, numeric_vars_q5, "Q5")
print(plot_q5)
save_plot(plot_q5, "Q5_bar_plot.png")

plot_q52 <- create_bar_plot(df, numeric_vars_q52, "Q5")
print(plot_q52)
save_plot(plot_q52, "Q5_bar_plot2.png")

plot_q53 <- create_bar_plot(df, numeric_vars_q53, "Q5")
print(plot_q53)
save_plot(plot_q53, "Q5_bar_plot3.png")

# Create percentage bar plots
plot_q3_percentage <- create_percentage_bar_plot(df, q3_vars, "Q3 - What types of Part-Time Jobs are you interested in?")
print(plot_q3_percentage)
ggsave(filename = "Q3_percentage_bar_plot.png", plot = plot_q3_percentage, path = getwd())

plot_q5_percentage <- create_percentage_bar_plot(df, q5_vars, "Q5 - What types of Part-Time Jobs are you interested in?")
print(plot_q5_percentage)
ggsave(filename = "Q5_percentage_bar_plot.png", plot = plot_q5_percentage, path = getwd())
  
