# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/alobireed/Rerun")

# Load the openxlsx library
library(openxlsx)
library(dplyr)

# Read data from a specific sheet of an Excel file
df_final <- read.xlsx("SpirometryData_Clean.xlsx")
df_HHS_data <- read.xlsx("HHS Completed_Cleaned.xlsx", sheet = "Clean")
df_airquality_data <- read.xlsx("AirQuality_Data_Clean.xlsx")

# Replace all negative values in df_airquality_data with NA
df_airquality_data[df_airquality_data < 0] <- NA

df_airquality_data <- df_airquality_data %>%
  mutate(Indicator = gsub("[^[:alnum:]]", "_", Indicator))  # Replace special chars with "_"


# Get rid of special characters
names(df_final) <- gsub(" ", "_", trimws(names(df_final)))
names(df_final) <- gsub("\\s+", "_", trimws(names(df_final), whitespace = "[\\h\\v\\s]+"))
names(df_final) <- gsub("\\(", "_", names(df_final))
names(df_final) <- gsub("\\)", "_", names(df_final))
names(df_final) <- gsub("\\-", "_", names(df_final))
names(df_final) <- gsub("/", "_", names(df_final))
names(df_final) <- gsub("\\\\", "_", names(df_final)) 
names(df_final) <- gsub("\\?", "", names(df_final))
names(df_final) <- gsub("\\'", "", names(df_final))
names(df_final) <- gsub("\\,", "_", names(df_final))
names(df_final) <- gsub("\\$", "", names(df_final))
names(df_final) <- gsub("\\+", "_", names(df_final))

str(df_final)
str(df_HHS_data)
str(df_airquality_data)




# Load required libraries
library(dplyr)
library(tidyr)
library(lubridate)

# Define the mapping of individuals to their date ranges
individual_dates <- tibble(
  Individual = c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10),
  StartDate = as.Date(c("2022-07-05", "2022-07-05", 
                        "2022-07-25", "2022-07-25", 
                        "2022-08-15", "2022-08-15", 
                        "2022-12-22", "2022-12-22", 
                        "2024-02-09", "2024-02-09")),
  EndDate = as.Date(c("2022-07-19", "2022-07-19", 
                      "2022-08-06", "2022-08-06", 
                      "2022-08-31", "2022-08-31", 
                      "2023-01-05", "2023-01-05", 
                      "2024-02-23", "2024-02-23"))
)

df_airquality_long <- df_airquality_data %>%
  pivot_longer(cols = `1`:`10`, names_to = "Individual", values_to = "Value") %>%
  mutate(Individual = as.numeric(Individual))

# Create a time sequence for 84 observations per individual (4-hour intervals)
time_sequence <- seq(from = 0, by = 4, length.out = 84)  # 4-hour gaps

# Convert air quality data from wide to long format (if not already done)
df_airquality_long <- df_airquality_long %>%
  left_join(individual_dates, by = "Individual") %>%  # Join to get StartDate and EndDate
  group_by(Individual, Indicator) %>%  # Restart date assignment for each Indicator
  mutate(
    ObservationIndex = row_number(),  # Assign row numbers within each Individual & Indicator
    DayIndex = (ObservationIndex - 1) %/% 6,  # Get day index (0-based)
    TimeIndex = (ObservationIndex - 1) %% 6,  # Get time of day index
    Date = StartDate + DayIndex,  # Restart date sequence per Indicator
    Time = hms::hms(hours = TimeIndex * 4),  # Assign 4-hour intervals
    DateTime = as.POSIXct(paste(Date, Time))
  ) %>%
  ungroup() %>%
  select(-ObservationIndex, -DayIndex, -TimeIndex, -StartDate, -EndDate)

# Convert spirometry data to include proper timestamps
df_final <- df_final %>%
  mutate(Date = as.Date(Day_time, origin = "1899-12-30"),  # Convert Excel serial date
         Time = ifelse(X3 == "AM", "08:00:00", "20:00:00"),  # Assign AM/PM timestamps
         DateTime = as.POSIXct(paste(Date, Time))) %>%
  select(-Day_time, -X3)

# Select only relevant columns from df_airquality_long to avoid duplication
df_airquality_long_clean <- df_airquality_long %>%
  select(-Date, -Time)

# Perform a LEFT JOIN to keep all air quality data and match available spirometry data
df_merged <- df_airquality_long_clean %>%
  left_join(df_final, by = c("Individual", "DateTime"))

# Remove rows where FVC is NA
df_merged <- df_merged %>%
  filter(!is.na(FVC))

# Ensure the HHS dataset has the correct identifier before merging
df_HHS_data <- df_HHS_data %>%
  rename(Individual = Respondent)  # Rename Respondent to Individual for proper merging

# Merge the HHS data with the already merged spirometry & air quality dataset
df_final <- df_merged %>%
  inner_join(df_HHS_data, by = "Individual")

# Define downwind individuals
downwind_ids <- c(2, 4, 6, 7, 9)

# Create a new column 'Location' to classify individuals
df_final <- df_final %>%
  mutate(Location = ifelse(Individual %in% downwind_ids, "Downwind", "Upwind"))

# Create a new column 'Location' to classify individuals
df_HHS_data <- df_HHS_data %>%
  mutate(Location = ifelse(Individual %in% downwind_ids, "Downwind", "Upwind"))

# Create a new column 'Location' to classify individuals
df_airquality_long_clean <- df_airquality_long_clean %>%
  mutate(Location = ifelse(Individual %in% downwind_ids, "Downwind", "Upwind"))

str(df_final)
colnames(df_final)
colnames(df_HHS_data)

create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each category variable
  for (category in categories) {
    # Ensure the category variable is a factor
    data[[category]] <- factor(data[[category]])
    
    # Calculate counts
    counts <- table(data[[category]])
    
    # Create a dataframe for this category
    freq_table <- data.frame(
      "Category" = rep(category, length(counts)),
      "Level" = names(counts),
      "Count" = as.integer(counts),
      stringsAsFactors = FALSE
    )
    
    # Calculate and add percentages
    freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
    
    # Add the result to the list
    all_freq_tables[[category]] <- freq_table
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}


vars_freq <- c("1_Do.you.regularly.have.any.of.the.following.complaints.concerning.the.indoor.air.you.breathe.inside.your.home?._Temperature"   ,                                          
 "1_Do.you.regularly.have.any.of.the.following.complaints.concerning.the.indoor.air.you.breathe.inside.your.home?._Dusty"                        ,                           
 "1_Do.you.regularly.have.any.of.the.following.complaints.concerning.the.indoor.air.you.breathe.inside.your.home?._Odors"                         ,                          
 "1_Do.you.regularly.have.any.of.the.following.complaints.concerning.the.indoor.air.you.breathe.inside.your.home?._Humidity"                       ,                         
 "1_Do.you.regularly.have.any.of.the.following.complaints.concerning.the.indoor.air.you.breathe.inside.your.home?._No_Complaints"                   ,                        
"1A_If.you.checked.any.of.the.above,.when.do.indoor.air.problems.seem.to.be.most.notable?_Morning"                                                   ,                      
 "1A_If.you.checked.any.of.the.above,.when.do.indoor.air.problems.seem.to.be.most.notable?_Afternoon"                                                 ,                      
 "1A_If.you.checked.any.of.the.above,.when.do.indoor.air.problems.seem.to.be.most.notable?_Night"                                                      ,                     
"1A_If.you.checked.any.of.the.above,.when.do.indoor.air.problems.seem.to.be.most.notable?_All_Day"                                                      ,                   
 "1A_If.you.checked.any.of.the.above,.when.do.indoor.air.problems.seem.to.be.most.notable?_All_Different_Times"                                          ,                   
"1A_If.you.checked.any.of.the.above,.when.do.indoor.air.problems.seem.to.be.most.notable?_Never"                                                          ,                 
"2_Do.you.or.others.in.your.home.regularly._Use_Strong_Cleaners"                                                                                           ,                
 "2_Do.you.or.others.in.your.home.regularly._Use_Air_Fresheners"                                                                                            ,                
 "2_Do.you.or.others.in.your.home.regularly._Use_Candles"                                                                                                    ,               
"2_Do.you.or.others.in.your.home.regularly._Vape"                                                                                                             ,             
"2_Do.you.or.others.in.your.home.regularly._Smoke_Tobacco"                                                                                                     ,            
"2_Do.you.or.others.in.your.home.regularly._Someone_Smokes_or_Vapes_Daily"                                                                                      ,           
"2_Do.you.or.others.in.your.home.regularly._Smoke_or_Vape_Inside_Home"                                                                                           ,          
 "2_Do.you.or.others.in.your.home.regularly._Exposed_to_Secondhand_Smoke"                                                                                         ,          
"2_Do.you.or.others.in.your.home.regularly._None_of_the_Above"                                                                                                     ,        
"3_During.the.last.30.days.while.at.home,.have.you.experienced.any.of.the.following.symptoms?_Frequent_Cough"                                                       ,       
"3_During.the.last.30.days.while.at.home,.have.you.experienced.any.of.the.following.symptoms?_Nasal_Congestion"                                                      ,      
 "3_During.the.last.30.days.while.at.home,.have.you.experienced.any.of.the.following.symptoms?_Wheezing_or_Chest_Tightness"                                           ,      
"3_During.the.last.30.days.while.at.home,.have.you.experienced.any.of.the.following.symptoms?_Multiple_Colds"                                                          ,    
 "3_During.the.last.30.days.while.at.home,.have.you.experienced.any.of.the.following.symptoms?_Runny_Nose_Watery_Eyes"                                                  ,    
 "3_During.the.last.30.days.while.at.home,.have.you.experienced.any.of.the.following.symptoms?_Shortness_of_Breath"                                                      ,   
 "4_Do.most.of.the.symptoms.checked.above.go.away.within.1.hour.after.leaving.home?"                                                                                      ,  
 "4A_If.no,.do.they.go.away.when.you.are.on.vacation.or.away.visiting.family/friends?"                                                                                     , 
 "5_What.percentage.of.your.day.do.you.typically.spend.inside.your.home?"                                                                                                   ,
 "6_What.percentage.of.your.day.do.you.typically.spend.outdoors?"                                                                                                           ,
 "7_What.percentage.of.your.day.do.you.typically.spend.at.work?"                                                                                                            ,
 "8_Do.you.work.outside.the.home?"                                                                                                                                          ,
"9_How.many.times.a.week.do.you.typically.vacuum?"                                                                                                                         ,
 "10_How.many.times.a.week.do.you.typically.dust?"                                                                                                                          ,
 "11_Please.rank.how.dusty.you.think.your.home.is?"                                                                                                                         ,
 "12_Have.there.been.any.renovation/demolition.related.activities.occurring.in.or.just.outside.your.home.in.the.past.30.days?.(i.e.,.new.carpet,.painting,.HVAC.work,.etc.)",
 "13_Has.there.been.evidence.of.water.leaks.or.visible.signs.of.more.moisture.than.usual.in.your.home?"                                                                     ,
 "14_Do.you.notice.pests/rodents.in.your.home?"                                                                                                                             ,
 "15_Do.you.have.a.gas.appliances.(please.circle.below)_Gas_Stove"                                                                                                          ,
 "15_Do.you.have.a.gas.appliances.(please.circle.below)_Gas_Dryer"                                                                                                          ,
 "15_Do.you.have.a.gas.appliances.(please.circle.below)_Gas_Water_Heater"                                                                                                   ,
 "15_Do.you.have.a.gas.appliances.(please.circle.below)_Other_Gas_Appliance"                                                                                                ,
 "16_Do.you.have.ventilation.(oven.range,.vent).in.your.kitchen.that.you.use.regularly?"                                                                                    ,
 "17_Do.you.have/use.a.woodstove?"                                                                                                                                          ,
 "18_Do.you.use.gas,.kerosene.or.propane.space.heaters?"                                                                                                                    ,
 "19_Do.you.have.and.use.a.fireplace?"                                                                                                                                      ,
 "20_Do.you.have.an.attached.garage?"                                                                                                                                       ,
 "21_Do.you.have.a.basement?"                                                                                                                                               ,
"21A_If.yes.to.question.18.above,.is.the.basement.moist/damp/wet?"                                                                                                         ,
 "21B_If.yes.to.moist/damp/wet.basement,.in.the.past.year.would.you.say.it.has.been.wetter.than.prior.years?"                                                               ,
 "22_Do.you.have.wall-to-wall.carpeting.in.more.than.half.of.the.rooms.in.your.home?"                                                                                       ,
 "23_Do.you.have.any.furry.pets.in.your.home?"                                                                                                                              ,
 "24_How.do.you.mostly.cool.your.home?Cool_Home_Air_Conditioner"                                                                                                            ,
 "24_How.do.you.mostly.cool.your.home?Cool_Home_Fans"                                                                                                                       ,
"24_How.do.you.mostly.cool.your.home?Cool_Home_Natural_Ventilation"                                                                                                        ,
"24_How.do.you.mostly.cool.your.home?Cool_Home_Other_Method"                                                                                                               ,
 "25_How.often.would.you.say.you.leave.your.windows.open.when.the.weather.is.warm?"                                                                                         ,
"26_Are.you.concerned.about.the.outdoor.air.in.your.community.worsening.your.asthma?"                                                                                      ,
"27_Do.you.want.to.learn.about.how.your.environment.can.worsen.asthma.symptoms?"                                                                                           ,
"28_Would.you.like.to.receive.information.and.have.a.discussion.about.your.in-home.air.monitoring.results.within.30.days.of.completing.the.study?")


df_freq <- create_frequency_tables(df_HHS_data, vars_freq)


library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    Median = numeric(),
    SEM = numeric(),  # Standard Error of the Mean
    SD = numeric(),   # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    median_val <- median(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      Median = median_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

colnames(df_final)

vars_desc <- c("ACQ_On.average,.during.the.past.week,.how.often.were.you.woken.by.your.asthma.during.the.night?",                                                                          
 "ACQ_On.average,.during.the.past.week,.how.bad.were.your.asthma.symptoms.when.you.woke.up.in.the.morning?"      ,                                                           
 "ACQ_In.general,.during.the.past.week,.how.limited.were.you.in.your.activities.because.of.your.asthma?"          ,                                                          
 "ACQ_In.general,.during.the.past.week,.how.much.shortness.of.breath.did.you.experience.because.of.your.asthma?"   ,                                                         
 "ACQ_In.general,.during.the.past.week,.how.much.of.the.time.did.you.wheeze?"                                       ,                                                        
 "ACQ_On.average,.during.the.past.week,.how.many.puffs.of.short-acting.bronchodilator.have.you.used.each.day?"       ,                                                       
 "ACQ_Total"                                                                                                          ,                                                      
 "ACQ_Weighted")

df_descriptive_stats <- calculate_descriptive_stats(df_HHS_data, vars_desc)

str(df_final)

# Load required libraries
library(lme4)
library(lmerTest)  # Provides p-values
library(broom.mixed)
library(dplyr)
library(ggplot2)
library(ggpubr)
library(performance)  # For model diagnostics
library(forecast)
library(mgcv)       # For GAMM (Generalized Additive Mixed Model)
library(nlme)       # For correlation structures in mixed models


# ✅ **Convert `DateTime` to numeric before fitting models**
df_airquality_long_clean <- df_airquality_long_clean %>%
  mutate(DateTime = as.numeric(DateTime))  # Convert POSIXct to numeric

# ✅ Function to fit GAMM for multiple pollution indicators
fit_gamm_and_format_log <- function(data, indicators) {
  library(mgcv)
  library(broom.mixed)
  library(dplyr)
  
  results_list <- list()
  
  for (indicator in indicators) {
    df_sub <- data %>% filter(Indicator == indicator)
    
    # Verificação mínima de estrutura
    if (nrow(df_sub) < 10 ||
        length(unique(df_sub$Location)) < 2 ||
        all(is.na(df_sub$LogValue)) ||
        length(unique(df_sub$DateTime)) < 2 ||
        length(unique(df_sub$Individual)) < 2) {
      
      cat("❌ Skipping:", indicator, "- insufficient variation or data\n")
      next
    }
    
    # Modelo GAMM
    model <- tryCatch({
      gamm(
        LogValue ~ Location + s(as.numeric(DateTime)),  # Convertido diretamente no modelo
        random = list(Individual = ~1),
        correlation = corAR1(form = ~ as.numeric(DateTime) | Individual),
        data = df_sub
      )
    }, error = function(e) {
      cat("❌ Model failed for:", indicator, "-", e$message, "\n")
      return(NULL)
    })
    
    if (is.null(model)) next
    
    # Tidy
    tidy_res <- tryCatch({
      out <- tidy(model$lme, effects = "fixed")
      if (nrow(out) > 0) {
        out$Indicator <- indicator
        return(out)
      } else {
        cat("⚠️ No fixed effects returned for:", indicator, "\n")
        return(NULL)
      }
    }, error = function(e) {
      cat("❌ Tidy failed for:", indicator, "-", e$message, "\n")
      return(NULL)
    })
    
    if (!is.null(tidy_res)) {
      results_list[[indicator]] <- tidy_res
    }
  }
  
  bind_rows(results_list)
}



# ✅ **Clean `Indicator` column (Remove spaces & special characters)**
df_airquality_long_clean <- df_airquality_long_clean %>%
  mutate(Indicator = gsub("µ", "u", Indicator),  # Replace µ with 'u'
         Indicator = gsub("/", "_", Indicator),  # Replace '/' with '_'
         Indicator = gsub(" ", "", Indicator))   # Remove spaces

# ✅ **Get unique cleaned indicator names**
response_vars <- unique(df_airquality_long_clean$Indicator)

# ✅ **Run the GAMM models for each air pollution indicator**
gamm_results <- fit_gamm_and_format_log(df_airquality_long_clean, response_vars)

df_gamm_results <- gamm_results$results

run_gamm_diagnostics <- function(models, output_dir = getwd()) {
  for (indicator in names(models)) {
    model <- models[[indicator]]
    
    cat("\n### Diagnostics for:", indicator, "###\n")
    
    # Model Performance Check
    cat("\n### Model Performance ###\n")
    tryCatch({
      print(performance::check_model(model))
    }, error = function(e) {
      cat("\nError running check_model:\n", e$message, "\n")
    })
    
    # Residual Diagnostics
    cat("\n### Residual Diagnostics ###\n")
    
    # Extract adjusted residuals (accounts for correlation structure)
    residuals_adj <- residuals(model, type = "normalized")
    
    if (!is.null(residuals_adj)) {
      print(summary(residuals_adj))
      
      # Set up output filenames
      hist_file <- file.path(output_dir, paste0("Residual_Hist_", indicator, ".png"))
      qqplot_file <- file.path(output_dir, paste0("QQ_Plot_", indicator, ".png"))
      acf_file <- file.path(output_dir, paste0("ACF_Plot_", indicator, ".png"))
      
      # Histogram of adjusted residuals with better binning
      png(hist_file, width = 800, height = 600)
      hist(residuals_adj, main = paste("Histogram of Adjusted Residuals for", indicator), 
           xlab = "Adjusted Residuals", breaks = 40)  # Increased bins for better granularity
      dev.off()
      
      # QQ Plot of Adjusted Residuals
      png(qqplot_file, width = 800, height = 600)
      qqnorm(residuals_adj, main = paste("QQ Plot of Adjusted Residuals for", indicator))
      qqline(residuals_adj)
      dev.off()
      
      # ACF Plot of Adjusted Residuals
      png(acf_file, width = 800, height = 600)
      acf(residuals_adj, main = paste("ACF of Adjusted Residuals for", indicator))
      dev.off()
      
      cat("\nSaved plots to:", output_dir, "\n")
      
    } else {
      cat("\nNo residuals found.\n")
    }
  }
}

# Run diagnostics and save plots
run_gamm_diagnostics(gamm_results$models)


# Examine data

library(ggplot2)
library(dplyr)
library(lubridate)

# Convert Unix timestamp to a readable date-time format
df_airquality_long_clean <- df_airquality_long_clean %>%
  mutate(DateTime = as.POSIXct(DateTime, origin="1970-01-01", tz="UTC"))

# Filter data for COppb to avoid any errors due to other indicators
df_COppb <- df_airquality_long_clean %>%
  filter(Indicator == "CO_ppb_") %>%
  mutate(Value = ifelse(Value < 0, NA, Value))  # Replace negative values with NA

# Histogram of COppb values to view distribution
ggplot(df_COppb, aes(x = Value)) +
  geom_histogram(bins = 30, fill = "steelblue", color = "black") +
  ggtitle("Histogram of COppb Values") +
  xlab("COppb Concentration (ppb)") +
  ylab("Frequency") +
  theme_minimal()

# Scatter plot of COppb values over DateTime
ggplot(df_COppb, aes(x = DateTime, y = Value, color = Location)) +
  geom_point(alpha = 0.6) +
  geom_smooth(method = "loess", se = FALSE) +
  ggtitle("COppb Values Over Time by Location") +
  xlab("Date and Time") +
  ylab("COppb Concentration (ppb)") +
  theme_minimal() +
  scale_color_manual(values = c("Upwind" = "blue", "Downwind" = "red"))

# Box plot to examine the distribution of COppb by Location
ggplot(df_COppb, aes(x = Location, y = Value, fill = Location)) +
  geom_boxplot() +
  ggtitle("Box Plot of COppb Values by Location") +
  xlab("Location") +
  ylab("COppb Concentration (ppb)") +
  theme_minimal() +
  scale_fill_manual(values = c("Upwind" = "lightblue", "Downwind" = "salmon"))

# Log transformation
# Load necessary libraries
library(ggplot2)
library(dplyr)

# Remove rows where Value is NA
df_airquality_long_clean <- df_airquality_long_clean %>% filter(!is.na(Value))

# Apply logarithmic transformation (add 0.01 to avoid log(0))
df_airquality_long_clean$LogValue <- log(df_airquality_long_clean$Value + 0.01)

# Define output directory (current working directory)
output_dir <- getwd()

# Create and save boxplots for each Indicator
for (indicator in unique(df_airquality_long_clean$Indicator)) {
  
  # Filter data for the current indicator
  df_subset <- df_airquality_long_clean %>% filter(Indicator == indicator)
  
  # Create boxplot
  p <- ggplot(df_subset, aes(x = Location, y = LogValue, fill = Location)) +
    geom_boxplot(alpha = 0.6, outlier.color = "red", outlier.size = 2) +
    theme_minimal() +
    ggtitle(paste("Box Plot of Log-Transformed", indicator, "Values by Location")) +
    xlab("Location") +
    ylab("Log(Concentration)") +
    scale_fill_manual(values = c("Upwind" = "blue", "Downwind" = "red"))
  
  # Save the plot
  ggsave(filename = paste0(output_dir, "/", indicator, "_BoxPlot_LogTransformed.png"),
         plot = p, width = 8, height = 5, dpi = 300)
  
  # Print message
  cat("Saved:", paste0(output_dir, "/", indicator, "_BoxPlot_LogTransformed.png"), "\n")
}

# Display the last plot generated
print(p)

str(df_airquality_long_clean)

# ✅ Function to fit GAMM for multiple pollution indicators
fit_gamm_and_format_log <- function(data, indicators) {
  library(mgcv)
  library(broom.mixed)
  library(dplyr)
  
  results_list <- list()
  
  for (indicator in indicators) {
    df_sub <- data %>% filter(Indicator == indicator)
    
    if (nrow(df_sub) < 10) {
      cat("❌ Skipping:", indicator, "- not enough data\n")
      next
    }
    
    model <- tryCatch({
      gamm(
        LogValue ~ Location + s(DateTime),
        random = list(Individual = ~1),
        correlation = corAR1(form = ~ DateTime | Individual),
        data = df_sub
      )
    }, error = function(e) {
      cat("❌ Model failed for:", indicator, "-", e$message, "\n")
      return(NULL)
    })
    
    if (is.null(model)) next
    
    tidy_res <- tryCatch({
      out <- tidy(model$lme, effects = "fixed")
      out$Indicator <- indicator
      out
    }, error = function(e) {
      cat("❌ Tidy failed for:", indicator, "-", e$message, "\n")
      return(NULL)
    })
    
    if (!is.null(tidy_res)) {
      results_list[[indicator]] <- tidy_res
    }
  }
  
  bind_rows(results_list)
}


response_vars <- unique(df_airquality_long_clean$Indicator)
df_gamm_results_log <- fit_gamm_and_format_log(df_airquality_long_clean, response_vars)


df_gamm_results_log <- gamm_results_log$results

run_gamm_diagnostics_log <- function(models, output_dir = getwd()) {
  for (indicator in names(models)) {
    model <- models[[indicator]]
    
    cat("\n### Diagnostics for:", indicator, "###\n")
    
    # Model Performance Check
    cat("\n### Model Performance ###\n")
    tryCatch({
      print(performance::check_model(model))
    }, error = function(e) {
      cat("\nError running check_model:\n", e$message, "\n")
    })
    
    # Residual Diagnostics
    cat("\n### Residual Diagnostics ###\n")
    
    # Extract adjusted residuals (accounts for correlation structure)
    residuals_adj <- residuals(model, type = "normalized")
    
    if (!is.null(residuals_adj)) {
      print(summary(residuals_adj))
      
      # Set up output filenames
      hist_file <- file.path(output_dir, paste0("Residual_Hist_", indicator, ".png"))
      qqplot_file <- file.path(output_dir, paste0("QQ_Plot_", indicator, ".png"))
      acf_file <- file.path(output_dir, paste0("ACF_Plot_", indicator, ".png"))
      
      # Histogram of adjusted residuals with better binning
      png(hist_file, width = 800, height = 600)
      hist(residuals_adj, main = paste("Histogram of Adjusted Residuals for", indicator), 
           xlab = "Adjusted Residuals", breaks = 40)  # Increased bins for better granularity
      dev.off()
      
      # QQ Plot of Adjusted Residuals
      png(qqplot_file, width = 800, height = 600)
      qqnorm(residuals_adj, main = paste("QQ Plot of Adjusted Residuals for", indicator))
      qqline(residuals_adj)
      dev.off()
      
      # ACF Plot of Adjusted Residuals
      png(acf_file, width = 800, height = 600)
      acf(residuals_adj, main = paste("ACF of Adjusted Residuals for", indicator))
      dev.off()
      
      cat("\nSaved plots to:", output_dir, "\n")
      
    } else {
      cat("\nNo residuals found.\n")
    }
  }
}

# Run diagnostics and save plots
run_gamm_diagnostics_log(gamm_results_log$models)


df_final <- df_final %>% filter(!is.na(Value))

# Second Hypothesis

# Load required libraries
library(dplyr)
library(lubridate)
library(lme4)
library(lmerTest)
library(broom.mixed)

names(df_final) <- gsub(" ", "_", trimws(names(df_final)))
names(df_final) <- gsub("\\s+", "_", trimws(names(df_final), whitespace = "[\\h\\v\\s]+"))
names(df_final) <- gsub("\\(", "_", names(df_final))
names(df_final) <- gsub("\\)", "_", names(df_final))
names(df_final) <- gsub("\\-", "_", names(df_final))
names(df_final) <- gsub("/", "_", names(df_final))
names(df_final) <- gsub("\\\\", "_", names(df_final)) 
names(df_final) <- gsub("\\?", "", names(df_final))
names(df_final) <- gsub("\\'", "", names(df_final))
names(df_final) <- gsub("\\,", "_", names(df_final))
names(df_final) <- gsub("\\$", "", names(df_final))
names(df_final) <- gsub("\\+", "_", names(df_final))
names(df_final) <- gsub("%", "", names(df_final))

# Ensure data is sorted by Individual & DateTime before computing differences
df_final <- df_final %>%
  arrange(Individual, DateTime)

df_final$FEV1 <- as.numeric(df_final$FEV1)

# Compute Δ (differences) for lung function and air quality indicators
df_final <- df_final %>%
  group_by(Individual, Indicator) %>%
  mutate(
    Delta_Value = Value - lag(Value),  # Change in air quality indicator
    Delta_FVC = FVC - lag(FVC),  # Change in FVC
    Delta_FEV1 = FEV1 - lag(FEV1),  # Change in FEV1
    Delta_FEV1_FVC = `FEV1_FVC.ratio.__` - lag(`FEV1_FVC.ratio.__`)  # Change in FEV1/FVC ratio
  ) %>%
  ungroup()

colnames(df_final)

# Remove first rows per individual (where lag() produced NA)
df_final <- df_final %>%
  filter(!is.na(Delta_Value) & !is.na(Delta_FVC) & !is.na(Delta_FEV1) & !is.na(Delta_FEV1_FVC))

# Load required libraries
library(lme4)
library(lmerTest)
library(broom.mixed)
library(ggplot2)
library(dplyr)

# Function to fit LMM, extract residuals, and save QQ-plots for each Indicator
fit_lmm_by_indicator <- function(dep_var, df) {
  
  # Extract unique air quality indicators
  indicators <- unique(df$Indicator)
  
  # Initialize list to store results
  all_results <- list()
  
  for (indicator in indicators) {
    
    # Subset data for the specific Indicator
    df_subset <- df %>% filter(Indicator == indicator)
    
    # Skip iteration if there's insufficient data for modeling
    if (nrow(df_subset) < 10) {  
      cat("Skipping Indicator:", indicator, "- Not enough data\n")
      next
    }
    
    # Fit the LMM model
    model <- lmer(as.formula(paste0(dep_var, " ~ Delta_Value + Location + (1 | Individual)")),
                  data = df_subset, REML = FALSE)
    
    # Extract fixed effects results with p-values
    results <- tidy(model, effects = "fixed") %>%
      mutate(Dependent_Variable = dep_var, Indicator = indicator)
    
    # Store results
    all_results[[indicator]] <- results
    
    # Extract residuals
    residuals_lmm <- residuals(model)
    
    # QQ-Plot of residuals
    qq_plot <- ggplot(data.frame(Residuals = residuals_lmm), aes(sample = Residuals)) +
      stat_qq() +
      stat_qq_line() +
      theme_minimal() +
      ggtitle(paste("QQ Plot of Residuals for", dep_var, "-", indicator))
    
    # Save the QQ plot
    plot_filename <- paste0("QQ_Plot_", dep_var, "_", indicator, ".png")
    ggsave(plot_filename, qq_plot, width = 7, height = 5)
    
    cat("Saved QQ Plot:", plot_filename, "\n")
  }
  
  # Combine all results into a single dataframe
  df_results <- bind_rows(all_results)
  return(df_results)
}

# Run LMMs for each lung function metric across all indicators
df_lmm_results_fvc <- fit_lmm_by_indicator("Delta_FVC", df_final)
df_lmm_results_fev1 <- fit_lmm_by_indicator("Delta_FEV1", df_final)
df_lmm_results_fev1_fvc <- fit_lmm_by_indicator("Delta_FEV1_FVC", df_final)

# Combine results into a single tidy dataframe
df_lmm_results <- bind_rows(df_lmm_results_fvc, df_lmm_results_fev1, df_lmm_results_fev1_fvc)

# Print final results
print(df_lmm_results)

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
  "AirQuality Data" = df_airquality_long, 
  "Merged Data" = df_final, 
  "Sample Freq" = df_freq,
  "Descriptives" = df_descriptive_stats,
  "GAMM results" = df_gamm_results_log,
  "LMM Results" = df_lmm_results
  
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")


# Visuals

# Load necessary libraries
library(ggplot2)
library(dplyr)
library(lubridate)
library(ggpubr)

# Convert DateTime to POSIXct for proper time series visualization
df_final <- df_final %>%
  mutate(DateTime = as.POSIXct(DateTime, origin="1970-01-01", tz="UTC"))

# Filter relevant air quality indicators
indicators_to_plot <- c("CO_ppb_", "NO2_ppb", "O3_ppb", "SO2_ppb", "PM2_5_ug_m_")

# Prepare data for visualization
df_plot <- df_final %>%
  filter(Indicator %in% indicators_to_plot) %>%
  select(DateTime, Indicator, Value, Delta_FVC, Delta_FEV1, Delta_FEV1_FVC, Location) %>%
  pivot_longer(cols = c(Value, Delta_FVC, Delta_FEV1, Delta_FEV1_FVC), 
               names_to = "Metric", values_to = "Measurement")

# Generate the time series plot
p1 <- ggplot(df_plot, aes(x = DateTime, y = Measurement, color = Location)) +
  geom_line(alpha = 0.7) +
  facet_wrap(~ Indicator, scales = "free_y") +
  theme_minimal() +
  ggtitle("Trends in Air Quality and Lung Function Changes Over Time") +
  xlab("Date and Time") +
  ylab("Measurement Value") +
  scale_color_manual(values = c("Upwind" = "blue", "Downwind" = "red"))

# Save the plot
ggsave("TimeSeries_AirQuality_LungFunction.png", p1, width = 12, height = 6, dpi = 300)

print(p1)



# Load necessary libraries
library(ggplot2)

# Prepare data for scatter plots
df_scatter <- df_final %>%
  select(Indicator, Delta_Value, Delta_FVC, Delta_FEV1, Delta_FEV1_FVC, Location) %>%
  pivot_longer(cols = c(Delta_FVC, Delta_FEV1, Delta_FEV1_FVC), 
               names_to = "LungFunctionMetric", values_to = "LungFunctionChange")

# Generate scatter plots with regression lines
p2 <- ggplot(df_scatter, aes(x = Delta_Value, y = LungFunctionChange, color = Location)) +
  geom_point(alpha = 0.6) +
  geom_smooth(method = "lm", se = TRUE, linetype = "dashed") +
  facet_wrap(~ LungFunctionMetric + Indicator, scales = "free") +
  theme_minimal() +
  ggtitle("Scatter Plots of Lung Function Changes vs Air Pollution Changes") +
  xlab("Change in Air Pollutant Levels (Delta_Value)") +
  ylab("Change in Lung Function (Delta)") +
  scale_color_manual(values = c("Upwind" = "blue", "Downwind" = "red"))

# Save the plot
ggsave("Scatter_LungFunction_vs_AirPollution.png", p2, width = 12, height = 6, dpi = 300)

print(p2)


# Load necessary libraries
library(ggplot2)
library(interactions)

# Fit an interaction model
interaction_model <- lmer(Delta_FEV1 ~ Delta_Value * Location + (1 | Individual), data = df_final, REML = FALSE)

# Interaction plot
p3 <- interact_plot(interaction_model, pred = Delta_Value, modx = Location, 
                    interval = TRUE, plot.points = TRUE) +
  ggtitle("Interaction Effect of Location on Air Pollution and Lung Function Changes") +
  xlab("Change in Air Pollutant Levels (Delta_Value)") +
  ylab("Change in FEV1") +
  theme_minimal()

# Save the plot
ggsave("InteractionPlot_Location_AirPollution.png", p3, width = 8, height = 6, dpi = 300)

print(p3)

colnames(df_final)




# Load necessary libraries
library(lme4)
library(lmerTest)
library(dplyr)

# Create a binary variable for ACQ threshold
df_final <- df_final %>%
  mutate(ACQ_High = ifelse(ACQ_Weighted >= 1.5, "High", "Low"))

colnames(df_final) <- gsub("%", "_", colnames(df_final))
colnames(df_final)

# Fit LMM: Does ACQ ≥1.5 predict FEV1/FVC ratio?
model_fev1_fvc <- lmer(FEV1_FVC.ratio.___ ~ ACQ_High + (1 | Individual), data = df_final, REML = FALSE)

# Summary of results
summary(model_fev1_fvc)

# Extract fixed effects with p-values
library(broom.mixed)
fev1_fvc_results <- tidy(model_fev1_fvc, effects = "fixed")



# Load necessary libraries
library(lme4)
library(lmerTest)
library(dplyr)
library(broom.mixed)

# Ensure column names are clean
colnames(df_final) <- gsub("%", "_", colnames(df_final))

colnames(df_final)
head(df_merged)


# Ensure ACQ_Weighted is numeric
df_final <- df_final %>%
  mutate(ACQ_Weighted = as.numeric(ACQ_Weighted))

# Select only Individual and ACQ_Weighted, ensuring one row per Individual
df_acq_weighted <- df_final %>%
  select(Individual, ACQ_Weighted) %>%
  distinct()  # Keep only one ACQ_Weighted value per Individual

# Merge ACQ_Weighted into df_merged using Individual as key
df_merged <- df_merged %>%
  left_join(df_acq_weighted, by = "Individual")

# Check if the merge was successful
head(df_merged)




df_airquality_summary <- df_merged %>%
  group_by(Individual, Indicator) %>%
  summarise(AirQuality_Mean = mean(Value, na.rm = TRUE)) %>%
  ungroup()

df_acq_analysis <- df_merged %>%
  select(Individual, ACQ_Weighted) %>%
  distinct() %>%
  left_join(df_airquality_summary, by = "Individual")

# Fit cross-sectional LMM (random effect on Individual not needed anymore)
lm_acq <- lm(ACQ_Weighted ~ AirQuality_Mean, data = df_acq_analysis)

summary(lm_acq)





df_acq_unique <- df_final %>%
  select(Individual, ACQ_High) %>%
  distinct()  # Keep only one row per Individual

df_merged <- df_merged %>%
  left_join(df_acq_unique, by = "Individual")

table(df_merged$ACQ_High)  # Check distribution of High/Low


library(lme4)
library(broom.mixed)
library(dplyr)

# Convert ACQ_High to a binary factor
df_merged <- df_merged %>%
  mutate(ACQ_High = factor(ACQ_High, levels = c("Low", "High")))

# List of air quality indicators
indicators <- unique(df_merged$Indicator)

glmm_results <- list()  # Store models

for (indicator in indicators) {
  # Subset data for the current indicator
  df_subset <- df_merged %>% filter(Indicator == indicator)
  
  # Fit GLMM with Indicator's Value as IV
  model <- glmer(ACQ_High ~ Value + (1 | Individual), 
                 data = df_subset, 
                 family = binomial, 
                 control = glmerControl(optimizer = "bobyqa"))  # Use optimizer for better convergence
  
  # Store results
  glmm_results[[indicator]] <- tidy(model) %>%
    mutate(Indicator = indicator)  # Add indicator name
}

# Combine results into a dataframe
df_glmm_results <- bind_rows(glmm_results)

# Print model summary
print(df_glmm_results)







# Load required libraries
library(openxlsx)
library(dplyr)
library(lme4)
library(lmerTest)
library(broom.mixed)

# Create a new workbook
wb <- createWorkbook()

# Add LMM results sheet
addWorksheet(wb, "LMM_Results")
writeData(wb, "LMM_Results", df_lmm_results)

# Add Linear Model results sheet
addWorksheet(wb, "LM_ACQ_Results")
lm_acq_results <- data.frame(
  Variable = c("Intercept", "AirQuality_Mean"),
  Estimate = coef(summary(lm_acq))[, "Estimate"],
  `p-value` = coef(summary(lm_acq))[, "Pr(>|t|)"]
)
writeData(wb, "LM_ACQ_Results", lm_acq_results)

# Add GLMM results sheet
addWorksheet(wb, "GLMM_Results")
writeData(wb, "GLMM_Results", df_glmm_results)

# Add FEV1/FVC model results sheet
addWorksheet(wb, "FEV1_FVC_Model")
writeData(wb, "FEV1_FVC_Model", fev1_fvc_results)

# Define file path and save the Excel file
file_path <- "AirQuality_ACQ_Models2.xlsx"
saveWorkbook(wb, file_path, overwrite = TRUE)

# Print file path for reference
print(paste("File saved at:", file_path))











# Load necessary libraries
library(dplyr)
library(broom)

# Aggregate air quality data per individual and indicator
df_airquality_summary <- df_merged %>%
  group_by(Individual, Indicator) %>%
  summarise(AirQuality_Mean = mean(Value, na.rm = TRUE), .groups = "drop")

# Extract unique ACQ scores per individual
df_acq_analysis <- df_merged %>%
  select(Individual, ACQ_Weighted) %>%
  distinct() %>%
  left_join(df_airquality_summary, by = "Individual")

# Initialize list to store results
lm_results_list <- list()

# Get unique air quality indicators
indicators <- unique(df_acq_analysis$Indicator)

# Loop through each air quality indicator and fit a separate model
for (indicator in indicators) {
  
  # Subset data for the current indicator
  df_subset <- df_acq_analysis %>%
    filter(Indicator == indicator)
  
  # Skip if there's insufficient data
  if (nrow(df_subset) < 10) {  
    cat("Skipping Indicator:", indicator, "- Not enough data\n")
    next
  }
  
  # Fit linear model
  lm_model <- lm(ACQ_Weighted ~ AirQuality_Mean, data = df_subset)
  
  # Extract coefficients and store results
  results <- tidy(lm_model) %>%
    mutate(Indicator = indicator)
  
  lm_results_list[[indicator]] <- results
  
  # Print model summary
  cat("\n### Model Summary for:", indicator, "###\n")
  print(summary(lm_model))
}

# Combine all results into a single dataframe
df_lm_results <- bind_rows(lm_results_list)

# Print final results
print(df_lm_results)















# RERUN

# ✅ Updated GAMM function including Environment
fit_gamm_and_format_log <- function(data, indicators) {
  library(mgcv)
  library(broom.mixed)
  library(dplyr)
  
  results_list <- list()
  models_list <- list()
  
  for (indicator in indicators) {
    df_sub <- data %>% filter(Indicator == indicator)
    
    if (nrow(df_sub) < 10) {
      cat("\u274c Skipping:", indicator, "- not enough data\n")
      next
    }
    
    model <- tryCatch({
      gamm(
        LogValue ~ Location + Environment + s(DateTime),
        random = list(Individual = ~1),
        correlation = corAR1(form = ~ DateTime | Individual),
        data = df_sub
      )
    }, error = function(e) {
      cat("\u274c Model failed for:", indicator, "-", e$message, "\n")
      return(NULL)
    })
    
    if (is.null(model)) next
    
    tidy_res <- tryCatch({
      out <- tidy(model$lme, effects = "fixed")
      out$Indicator <- indicator
      out
    }, error = function(e) {
      cat("\u274c Tidy failed for:", indicator, "-", e$message, "\n")
      return(NULL)
    })
    
    if (!is.null(tidy_res)) {
      results_list[[indicator]] <- tidy_res
      models_list[[indicator]] <- model$lme
    }
  }
  
  list(results = bind_rows(results_list), models = models_list)
}

# ✅ Updated GLMM model (ACQ_High ~ Value + Environment)
glmm_results <- list()

for (indicator in unique(df_merged$Indicator)) {
  df_subset <- df_merged %>% filter(Indicator == indicator)
  
  if (nrow(df_subset) < 10 || any(is.na(df_subset$Environment))) {
    next
  }
  
  model <- glmer(ACQ_High ~ Value + Environment + (1 | Individual),
                 data = df_subset,
                 family = binomial,
                 control = glmerControl(optimizer = "bobyqa"))
  
  glmm_results[[indicator]] <- tidy(model) %>%
    mutate(Indicator = indicator)
}

df_glmm_results <- bind_rows(glmm_results)

# ✅ Updated LMM model (ACQ_Weighted ~ AirQuality_Mean + Environment)
df_airquality_summary <- df_merged %>%
  group_by(Individual, Indicator, Environment) %>%
  summarise(AirQuality_Mean = mean(Value, na.rm = TRUE), .groups = "drop")

df_acq_analysis <- df_merged %>%
  select(Individual, ACQ_Weighted) %>%
  distinct() %>%
  left_join(df_airquality_summary, by = "Individual")

lm_results_list <- list()

for (indicator in unique(df_acq_analysis$Indicator)) {
  df_subset <- df_acq_analysis %>% filter(Indicator == indicator)
  
  if (nrow(df_subset) < 10 || any(is.na(df_subset$Environment))) {
    next
  }
  
  lm_model <- lm(ACQ_Weighted ~ AirQuality_Mean + Environment, data = df_subset)
  
  results <- tidy(lm_model) %>% mutate(Indicator = indicator)
  lm_results_list[[indicator]] <- results
}

df_lm_results <- bind_rows(lm_results_list)

# ✅ Updated FEV1/FVC model
model_fev1_fvc <- lmer(FEV1_FVC.ratio.___ ~ ACQ_High + Environment + (1 | Individual),
                       data = df_final, REML = FALSE)
fev1_fvc_results <- tidy(model_fev1_fvc, effects = "fixed")


