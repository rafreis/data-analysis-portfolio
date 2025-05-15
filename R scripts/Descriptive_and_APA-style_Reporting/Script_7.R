# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/alobireed")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df_spirometrydata <- read.xlsx("SpirometryData_Clean.xlsx")
df_HHS_data <- read.xlsx("HHS Completed_Cleaned.xlsx", sheet = "Clean")
df_airquality_data <- read.xlsx("AirQuality_Data_Clean.xlsx")

# Replace all negative values in df_airquality_data with NA
df_airquality_data[df_airquality_data < 0] <- NA

df_airquality_data <- df_airquality_data %>%
  mutate(Indicator = gsub("[^[:alnum:]]", "_", Indicator))  # Replace special chars with "_"


# Get rid of special characters
names(df_spirometrydata) <- gsub(" ", "_", trimws(names(df_spirometrydata)))
names(df_spirometrydata) <- gsub("\\s+", "_", trimws(names(df_spirometrydata), whitespace = "[\\h\\v\\s]+"))
names(df_spirometrydata) <- gsub("\\(", "_", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("\\)", "_", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("\\-", "_", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("/", "_", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("\\\\", "_", names(df_spirometrydata)) 
names(df_spirometrydata) <- gsub("\\?", "", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("\\'", "", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("\\,", "_", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("\\$", "", names(df_spirometrydata))
names(df_spirometrydata) <- gsub("\\+", "_", names(df_spirometrydata))

str(df_spirometrydata)
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
df_spirometrydata <- df_spirometrydata %>%
  mutate(Date = as.Date(Day_time, origin = "1899-12-30"),  # Convert Excel serial date
         Time = ifelse(X3 == "AM", "08:00:00", "20:00:00"),  # Assign AM/PM timestamps
         DateTime = as.POSIXct(paste(Date, Time))) %>%
  select(-Day_time, -X3)

# Select only relevant columns from df_airquality_long to avoid duplication
df_airquality_long_clean <- df_airquality_long %>%
  select(-Date, -Time)

# Perform a LEFT JOIN to keep all air quality data and match available spirometry data
df_merged <- df_airquality_long_clean %>%
  left_join(df_spirometrydata, by = c("Individual", "DateTime"))

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

# ✅ Ensure `see` package is installed for visualization in performance package
if (!requireNamespace("see", quietly = TRUE)) {
  install.packages("see")
}

# ✅ **Convert `DateTime` to numeric before fitting models**
df_airquality_long_clean <- df_airquality_long_clean %>%
  mutate(DateTime = as.numeric(DateTime))  # Convert POSIXct to numeric

# ✅ **Function to fit LMM for multiple pollution indicators**
fit_lmm_and_format <- function(data, indicators, save_plots = FALSE) {
  lmm_results_list <- list()
  lmm_models_list <- list()  # Store models for diagnostics
  
  for (indicator in indicators) {
    # Filter data for the current Indicator
    data_subset <- data %>% filter(Indicator == indicator)
    
    # Skip iteration if there's insufficient data for modeling
    if (nrow(data_subset) < 5) {
      cat("Skipping Indicator:", indicator, "- Not enough data\n")
      next
    }
    
    # Construct the model formula
    formula <- as.formula("Value ~ Location + (1 | Individual) + (1 | DateTime)")
    
    # Fit the linear mixed-effects model
    lmer_model <- lmer(formula, data = data_subset)
    
    # Store the model in the list
    lmm_models_list[[indicator]] <- lmer_model
    
    # Extract the tidy output for fixed effects (includes p-values)
    fixed_effects_results <- tidy(lmer_model, effects = "fixed") %>%
      mutate(Indicator = indicator)  # Add indicator name
    
    # Extract random effects variance
    random_effects_results <- as.data.frame(VarCorr(lmer_model))
    
    # Store results in a list
    lmm_results_list[[indicator]] <- list(
      fixed = fixed_effects_results,
      random = random_effects_results
    )
    
    # Print model summary for debugging
    print(summary(lmer_model))
    
    # Generate and save residual plots if requested
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", indicator)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lmer_model) ~ fitted(lmer_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lmer_model), main = "Q-Q Plot")
      qqline(residuals(lmer_model))
      dev.off()
    }
  }
  
  # Combine fixed effects results into a tidy dataframe
  df_lmm_results_fixed <- bind_rows(lapply(lmm_results_list, `[[`, "fixed"))
  
  return(list(results = df_lmm_results_fixed, models = lmm_models_list))  # Return results + models
}

# ✅ **Clean `Indicator` column (Remove spaces & special characters)**
df_airquality_long_clean <- df_airquality_long_clean %>%
  mutate(Indicator = gsub("µ", "u", Indicator),  # Replace µ with 'u'
         Indicator = gsub("/", "_", Indicator),  # Replace '/' with '_'
         Indicator = gsub(" ", "", Indicator))   # Remove spaces

# ✅ **Get unique cleaned indicator names**
response_vars <- unique(df_airquality_long_clean$Indicator)

# ✅ **Run the LMM models for each air pollution indicator**
lmm_results <- fit_lmm_and_format(df_airquality_long_clean, response_vars, save_plots = TRUE)

df_lmm_results <- lmm_results$results

# ✅ **Function to run diagnostics on fitted LMM models**
run_lmm_diagnostics <- function(models) {
  for (indicator in names(models)) {
    model <- models[[indicator]]
    
    cat("\n### Diagnostics for:", indicator, "###\n")
    
    # Get residuals
    residuals <- residuals(model)
    fitted_values <- fitted(model)
    
    # Create diagnostic plots
    qq_plot <- ggqqplot(residuals) + 
      ggtitle(paste("Q-Q Plot:", indicator))
    
    residuals_vs_fitted_plot <- ggplot(data.frame(Fitted = fitted_values, Residuals = residuals), aes(x = Fitted, y = Residuals)) +
      geom_point() +
      geom_smooth(method = "loess", se = FALSE, color = "red") +
      ggtitle(paste("Residuals vs Fitted:", indicator)) +
      theme_minimal()
    
    # Check if forecast package is available for ACF plot
    if ("forecast" %in% installed.packages()) {
      acf_plot <- ggAcf(residuals) + ggtitle(paste("Autocorrelation (ACF):", indicator))  # FIXED: Removed autoplot()
    } else {
      acf_values <- acf(residuals, plot = FALSE)$acf[-1]
      acf_plot <- ggplot(data.frame(ACF = acf_values, Lag = 1:length(acf_values)), aes(x = Lag, y = ACF)) +
        geom_bar(stat = "identity") +
        ggtitle(paste("Autocorrelation (ACF):", indicator)) +
        theme_minimal()
    }
    
    # Print plots
    print(qq_plot)
    print(residuals_vs_fitted_plot)
    print(acf_plot)
    
    # Additional model performance check
    cat("\n### Model Performance ###\n")
    print(check_model(model))
  }
}

# ✅ **Run diagnostics on all fitted models**
run_lmm_diagnostics(lmm_results$models)
