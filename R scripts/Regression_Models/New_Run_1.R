# === 1. SETUP ===
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/alobireed/Rerun")

library(openxlsx)
library(dplyr)
library(tidyr)
library(lubridate)
library(hms)
library(mgcv)
library(nlme)
library(broom.mixed)
library(lme4)
library(lmerTest)
library(ggplot2)
library(ggpubr)
library(performance)
library(interactions)

# === 2. LOAD DATA ===
df_final <- read.xlsx("SpirometryData_Clean.xlsx")
df_HHS_data <- read.xlsx("HHS Completed_Cleaned.xlsx", sheet = "Clean")
df_airquality_data <- read.xlsx("AirQuality_Data_Clean.xlsx")

# Clean column names
df_airquality_data[df_airquality_data < 0] <- NA
df_airquality_data <- df_airquality_data %>% mutate(Indicator = gsub("[^[:alnum:]]", "_", Indicator))
names(df_final) <- gsub("[[:space:]()\\/?,'$%+\\-]", "_", trimws(names(df_final)))

names(df_HHS_data) <- gsub("[[:space:]()\\/?,'$%+\\-]", "_", trimws(names(df_HHS_data)))


# === 3. BUILD TIMESTAMP FOR AIR QUALITY ===
individual_dates <- tibble(
  Individual = 1:10,
  StartDate = as.Date(c("2022-07-05", "2022-07-05", "2022-07-25", "2022-07-25", "2022-08-15", "2022-08-15", "2022-12-22", "2022-12-22", "2024-02-09", "2024-02-09")),
  EndDate = as.Date(c("2022-07-19", "2022-07-19", "2022-08-06", "2022-08-06", "2022-08-31", "2022-08-31", "2023-01-05", "2023-01-05", "2024-02-23", "2024-02-23"))
)

df_airquality_long <- df_airquality_data %>%
  pivot_longer(cols = `1`:`10`, names_to = "Individual", values_to = "Value") %>%
  mutate(Individual = as.numeric(Individual)) %>%
  left_join(individual_dates, by = "Individual") %>%
  group_by(Individual, Indicator) %>%
  mutate(
    ObservationIndex = row_number(),
    DayIndex = (ObservationIndex - 1) %/% 6,
    TimeIndex = (ObservationIndex - 1) %% 6,
    Date = StartDate + DayIndex,
    Time = hms(hours = TimeIndex * 4),
    DateTime = as.POSIXct(paste(Date, Time))
  ) %>%
  ungroup() %>%
  select(-ObservationIndex, -DayIndex, -TimeIndex, -StartDate, -EndDate)

# Save timestamped air quality
df_airquality_long_clean <- df_airquality_long %>% select(-Date, -Time)
df_airquality_long_clean$LogValue <- log(df_airquality_long_clean$Value + 0.01)
downwind_ids <- c(2, 4, 6, 7, 9)
df_airquality_long_clean$Location <- ifelse(df_airquality_long_clean$Individual %in% downwind_ids, "Downwind", "Upwind")

# === 4. BUILD TIMESTAMP FOR SPIROMETRY ===
df_final <- df_final %>%
  mutate(Date = as.Date(Day_time, origin = "1899-12-30"),
         Time = ifelse(X3 == "AM", "08:00:00", "20:00:00"),
         DateTime = as.POSIXct(paste(Date, Time))) %>%
  select(-Day_time, -X3)

df_HHS_data <- df_HHS_data %>% rename(Individual = Respondent)
df_HHS_data$Location <- ifelse(df_HHS_data$Individual %in% downwind_ids, "Downwind", "Upwind")

# === 5. MERGE DATA ===
df_merged <- df_airquality_long_clean %>% left_join(df_final, by = c("Individual", "DateTime")) %>% filter(!is.na(FVC))
df_final <- df_merged %>% inner_join(df_HHS_data, by = "Individual")
df_final <- df_final %>% mutate(Location = ifelse(Individual %in% downwind_ids, "Downwind", "Upwind"))

# Save preprocessed merged data
saveRDS(df_final, "df_final_preprocessed.rds")

# === 6. MODEL 1: GAMM ===
response_vars <- unique(df_airquality_long_clean$Indicator)

gamm_results <- list()

for (indicator in response_vars) {
  df_sub <- df_airquality_long_clean %>% filter(Indicator == indicator)
  if (nrow(df_sub) < 10) next
  model <- tryCatch({
    gamm(LogValue ~ Location + Environment + s(as.numeric(DateTime)),
         random = list(Individual = ~1),
         correlation = corAR1(form = ~ as.numeric(DateTime) | Individual),
         data = df_sub)
  }, error = function(e) NULL)
  if (!is.null(model)) {
    res <- tidy(model$lme, effects = "fixed")
    res$Indicator <- indicator
    gamm_results[[indicator]] <- res
  }
}
df_gamm_results <- bind_rows(gamm_results)
write.xlsx(df_gamm_results, "GAMM_with_Environment.xlsx")

# === 7. MODEL 2: ACQ_Weighted ~ AirQuality_Mean + Environment ===
df_airquality_summary <- df_merged %>% group_by(Individual, Indicator, Environment) %>% summarise(AirQuality_Mean = mean(Value, na.rm = TRUE), .groups = "drop")
df_acq_weighted <- df_final %>% select(Individual, ACQ_Weighted) %>% distinct()
df_acq_analysis <- left_join(df_airquality_summary, df_acq_weighted, by = "Individual")

lm_results <- list()
for (indicator in unique(df_acq_analysis$Indicator)) {
  df_sub <- df_acq_analysis %>% filter(Indicator == indicator)
  if (nrow(df_sub) < 10) next
  model <- lm(ACQ_Weighted ~ AirQuality_Mean + Environment, data = df_sub)
  lm_results[[indicator]] <- tidy(model) %>% mutate(Indicator = indicator)
}
df_lm_results <- bind_rows(lm_results)
write.xlsx(df_lm_results, "LM_ACQ_with_Environment.xlsx")

# === 8. MODEL 3: ACQ_High ~ Value + Environment + (1|Individual) ===
df_final <- df_final %>% mutate(ACQ_High = factor(ifelse(ACQ_Weighted >= 1.5, "High", "Low")))
df_merged <- left_join(df_airquality_long_clean, df_final %>% select(Individual, DateTime, ACQ_High), by = c("Individual", "DateTime"))

glmm_results <- list()
for (indicator in unique(df_merged$Indicator)) {
  df_sub <- df_merged %>% filter(Indicator == indicator)
  if (nrow(df_sub) < 10) next
  model <- tryCatch({
    glmer(ACQ_High ~ Value + Environment + (1 | Individual),
          data = df_sub,
          family = binomial,
          control = glmerControl(optimizer = "bobyqa"))
  }, error = function(e) NULL)
  if (!is.null(model)) {
    glmm_results[[indicator]] <- tidy(model) %>% mutate(Indicator = indicator)
  }
}
df_glmm_results <- bind_rows(glmm_results)
write.xlsx(df_glmm_results, "GLMM_ACQ_High_with_Environment.xlsx")

# === 9. MODEL 4: FEV1_FVC ~ ACQ_High + Environment + (1 | Individual) ===
df_final$FEV1 <- as.numeric(df_final$FEV1)
model_fev1_fvc <- lmer(FEV1_FVC.ratio.___ ~ ACQ_High + Environment + (1 | Individual), data = df_final, REML = FALSE)
fev1_fvc_results <- tidy(model_fev1_fvc, effects = "fixed")
write.xlsx(fev1_fvc_results, "FEV1_FVC_Model_with_Environment.xlsx")
