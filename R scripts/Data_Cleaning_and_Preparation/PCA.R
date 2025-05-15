library(psych)
library(tidyverse)

Data2 <- df

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
EatingBeh2$Comp1 <- rowMeans(EatingBeh2[, c("Foodfussiness", "Foodresponsiveness", "Enjoymentoffood", "Emotionalovereating", "Emotionalundereating", "Satietyresponsiveness")], na.rm = TRUE)
EatingBeh2$Comp2 <- rowMeans(EatingBeh2[, c("BESScore", "rNEDQScore", "YFASScore", "Hunger")], na.rm = TRUE)
EatingBeh2$Comp3 <- rowMeans(EatingBeh2[, c("Slownessineating", "DQSScore")], na.rm = TRUE)


# Standardizing the data
stEatingBeh2 <- scale(EatingBeh2)


# Laden der benötigten Bibliotheken
library(FactoMineR)
library(factoextra)

# Durchführen der PCA
pca_result <- PCA(stEatingBeh2, graph = FALSE)

# Eigenwerte und Varianzerklärung anzeigen
eigenvalues <- get_eigenvalue(pca_result)
eigenvalues

# Scree Plot der Eigenwerte erstellen
fviz_eig(pca_result, addlabels = TRUE, ylim = c(0, 50))

# Biplot von Individuen und Variablen erstellen
biplot <- fviz_pca_biplot(pca_result, 
                          col.var = "contrib", # Färben nach Beitrag zur Variabilität
                          gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"),
                          repel = TRUE, # Verhindert Textüberlappung
                          axes = c(1, 2))  # Welche Hauptkomponenten sollen geplottet werden

# Biplot anzeigen
print(biplot)

##########################
# Biplot von Variablen erstellen (ohne Individuen)
biplot_var_only <- fviz_pca_var(pca_result, col.var = "contrib",
                                gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"),
                                repel = TRUE, axes = c(1, 2))

# Biplot anzeigen
print(biplot_var_only)

