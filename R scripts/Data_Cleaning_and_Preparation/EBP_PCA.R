library(psych)
library(tidyverse)

# Re-categorizing NEDQ in order
Data2 <- Data2 %>%
  mutate(
    rNEDQScore = case_when(
      NEDQScore == 1 ~ 1,
      NEDQScore == 2 ~ 3,
      NEDQScore == 3 ~ 4,
      NEDQScore == 4 ~ 2
    )
  )
  
EatingBeh2 <- Data2 %>% select(BESScore, DQSScore, rNEDQScore, YFASScore, Foodfussiness, Foodresponsiveness, Enjoymentoffood, Emotionalovereating, Emotionalundereating, Slownessineating, Hunger, Satietyresponsiveness)

# Reverse scoring necessary columns
  columns_to_reverse2 <- c("Foodresponsiveness", "Enjoymentoffood", "Emotionalovereating")
  
for (col in columns_to_reverse2) {
  EatingBeh2[, col] <- 6 - EatingBeh2[, col]
}

# Conducting PCA
fa.parallel(EatingBeh2, fa = "pc")

pca3 <- principal(EatingBeh2, nfactors = 3)
pca3

#Creating the composite index
EatingBeh2$Comp1 <- rowMeans(stEatingBeh2$Foodfussiness, stEatingBeh2$Foodresponsiveness, EatingBeh2$Enjoymentoffood, EatingBeh2$Emotionalovereating, EatingBeh2$Emotionalundereating, EatingBeh2$Satietyresponsiveness)
EatingBeh2$Comp2 <- rowMeans(EatingBeh2$BESScore, EatingBeh2$rNEDQScore, EatingBeh2$YFASScore, EatingBeh2$Hunger)
EatingBeh2$Comp3 <- rowMeans(EatingBeh2$Slownessineating, EatingBeh2$DQSScore)

# Standardizing the data
stEatingBeh2 <- scale(EatingBeh2)


