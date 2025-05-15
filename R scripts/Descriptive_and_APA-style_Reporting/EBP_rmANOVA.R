library(readr)
df <- read_csv("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/elenakoning/EBP_Survey_Data_2.csv")
View(df)

library(tidyverse)
df <- df %>%
  mutate(
    rNEDQScore = case_when(
      NEDQScore == 1 ~ 1,
      NEDQScore == 2 ~ 3,
      NEDQScore == 3 ~ 4,
      NEDQScore == 4 ~ 2
    )
  )

# Reverse scoring necessary columns
columns_to_reverse2 <- c("Foodresponsiveness", "Enjoymentoffood", "Emotionalovereating")
for (col in columns_to_reverse2) {
  df[, col] <- 6 - df[, col]
}

# Standardizing the data
df$zBESScore <- scale(df$BESScore)
df$zDQSScore <- scale(df$DQSScore)
df$zSlownessineating <- scale(df$Slownessineating)
df$zrNEDQScore <- scale(df$rNEDQScore)
df$zYFASScore <- scale(df$YFASScore)
df$zHunger <- scale(df$Hunger)
df$zFoodfussiness <- scale(df$Foodfussiness)
df$zFoodresponsiveness <- scale(df$Foodresponsiveness)
df$zEnjoymentoffood <- scale(df$Enjoymentoffood)
df$zEmotionalovereating <- scale(df$Emotionalovereating)
df$zEmotionalundereating <- scale(df$Emotionalundereating)
df$zSatietyresponsiveness <- scale(df$Satietyresponsiveness)

# Creating composite index
df$Comp1 <- rowMeans(subset(df, select = c(zFoodfussiness, zFoodresponsiveness, zEnjoymentoffood, zEmotionalovereating, 
                       zEmotionalundereating, zSatietyresponsiveness)))
df$Comp2 <- rowMeans(subset(df, select = c(zBESScore, zrNEDQScore, zYFASScore, zHunger)))
df$Comp3 <- rowMeans(subset(df, select = c(zSlownessineating, zDQSScore)))

# Pivoting data
dflong <- pivot_longer(df, cols = Comp1:Comp3, names_to = "CompNum", values_to = "DV")

# Coding data for MLM
dflong <- dflong %>%
  mutate(
    Men = case_when(
      Gender == 1 ~ 1,
      Gender == 2 ~ 0,
      Gender == 3 ~ -1),
    Women = case_when(
      Gender == 1 ~ 0,
      Gender == 2 ~ 1,
      Gender == 3 ~ -1),
    AcuteEat = case_when(
      CompNum == "Comp1" ~ 1,
      CompNum == "Comp2" ~ 0,
      CompNum == "Comp3" ~ -1),
    MealImpul = case_when(
      CompNum == "Comp1" ~ 0,
      CompNum == "Comp2" ~ 1,
      CompNum == "Comp3" ~ -1)
    )

dflong$cAge <- scale(dflong$Age, center = TRUE, scale = FALSE)

# Conducting MLM
library(lme4)
library(lmerTest)
mlm1 <- lmer(DV ~ cAge*AcuteEat + cAge*MealImpul + (1|ParticipantID), data = dflong)
summary(mlm1)

mlm1b <- lmer(DV ~ cAge*CompNum + (1|ParticipantID), data = dflong)
summary(mlm1b)

library(interactions)
sim_slopes(mlm1b, pred = "cAge", modx = "CompNum", data = dflong)

# Conducting repeated measures ANOVA
library(afex)

aov_ez("ParticipantID", "DV", dflong, between = "Diagnosis", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "Age", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "Gender", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "Medication", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "CannabisRoutineUse", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "Exercise", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "Alcoholdrinksperweek", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "Smoking", within = "CompNum")
aov_ez("ParticipantID", "DV", dflong, between = "BMI", within = "CompNum")

# Create a boxplot for the repeated measures ANOVA results
# Create a new data frame to store the ANOVA results
anova_results <- data.frame(
  Factor = c("Diagnosis", "Age", "Gender", "Medication", "CannabisRoutineUse", "Alcoholdrinksperweek", "Smoking", "Exercise", "BMI"),
  p_value = c(0.116, 0.007, 0.574, 0.035, 0.683, 0.232, 0.418, 0.241, 0.028) 
)


