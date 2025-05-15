setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2023/elenakoning/")

## BOXPLOTS

library(ggplot2)
library(dplyr)
library(plotly)
library(scatterplot3d)
library(htmlwidgets)

# List of factors used in RM ANOVA
# Ensure CompNum is a factor
dflong$CompNum <- factor(dflong$CompNum)

# Define colors for CompNum
comp_colors <- c("Comp1" = "blue", "Comp2" = "red", "Comp3" = "green")

# List of factors used in RM ANOVA
factors <- c("Diagnosis", "Age", "Gender", "Medication", "CannabisRoutineUse", 
             "Exercise", "Alcoholdrinksperweek", "Smoking")

dflong$BMICategory <- cut(dflong$BMI,
                      breaks = c(-Inf, 18.5, 24.9, 29.9, Inf),
                      labels = c("<18.5", "18.5-24.9", "25-29.9", "≥30"),
                      right = FALSE)

# Ensure the new BMI category is a factor
dflong$BMICategory <- factor(dflong$BMICategory)

# Add BMI category to the factors list
factors <- c(factors, "BMICategory")

# Relabeling the factors
dflong$Diagnosis <- factor(dflong$Diagnosis, levels = c(1, 3, 4), labels = c("MDD", "BDI", "BDII"))
dflong$Age <- factor(dflong$Age, levels = c(1, 2, 3, 4, 5), labels = c("Less than 20", "20 - 29", "30 - 39", "40 - 49", "50 - 59"))
dflong$Gender <- factor(dflong$Gender, levels = c(1, 2, 3), labels = c("Male", "Female", "Non-binary"))
dflong$Medication <- factor(dflong$Medication, levels = c(1, 2, 3, 4, 5), labels = c("Antidepressant", "Mood Stabilizer", "Anticonvulsant", "Antipsychotic", "None"))
dflong$CannabisRoutineUse <- factor(dflong$CannabisRoutineUse, levels = c(1, 2), labels = c("Yes", "No"))
dflong$Exercise <- factor(dflong$Exercise, levels = c(1, 2), labels = c("< 150 minutes last week", "> 150 minutes last week"))
dflong$Alcoholdrinksperweek <- factor(dflong$Alcoholdrinksperweek, levels = c(1, 2, 3, 4), labels = c("I rarely/never drink alcohol", "Less than 14 units", "Between 14 & 21 units", "More than 21 units"))
dflong$Smoking <- factor(dflong$Smoking, levels = c(1, 2, 3), labels = c("I have never smoked more than 100 cigarettes", "Current smoker", "Ex-smoker"))

# Loop through each factor and create boxplots
for(factor in factors) {
  plot <- dflong %>%
    ggplot(aes(x = CompNum, y = DV, fill = CompNum)) +  # Fill color based on CompNum
    geom_boxplot() +
    scale_fill_manual(values = comp_colors) +  # Apply defined color palette
    facet_wrap(~get(factor), scales = "free") +
    theme_bw() +  # White background theme
    ggtitle(paste("Boxplots of Participants' Component Scores by Component Number and", factor)) +
    xlab("Component Number") +
    ylab("Participants' Component Scores") +
    theme(legend.position = "none")  # Remove legend
  
  # Print and save the plot
  print(plot)
  ggsave(paste0("boxplot_", factor, ".png"), plot, width = 10, height = 6)
}


## SCATTERPLOT
df <- df %>%
  rename(
    `Component_1` = 'Comp1',
    `Component_2` = 'Comp2',
    `Component_3` = 'Comp3'
  )

df <- df %>%
  mutate(HighestComponent = factor(
    apply(.[, c("Component_1", "Component_2", "Component_3")], 1, function(x) which.max(abs(x)))
  ))

fig <- plot_ly(data = df, x = ~Component_1, y = ~Component_2, z = ~Component_3, color = ~HighestComponent, colors = c("blue", "red", "green")) %>%
  add_markers() %>%
  layout(scene = list(
    xaxis = list(title = 'Component 1'),
    yaxis = list(title = 'Component 2'),
    zaxis = list(title = 'Component 3')
  ))

fig
# Save the figure as an HTML file
saveWidget(fig, file = "3D_Scatter_Plot.html")


library(ggplot2)

# Map numeric values to component names
df$HighestComponent <- factor(
  case_when(
    df$HighestComponent == 1 ~ "Component_1",
    df$HighestComponent == 2 ~ "Component_2",
    df$HighestComponent == 3 ~ "Component_3"
  )
)

# Get all pairs of components
component_pairs <- combn(c("Component_1", "Component_2", "Component_3"), 2, simplify = FALSE)
# Define colors for CompNum
comp_colors <- c("Component 1" = "blue", "Component 2" = "red", "Component 3" = "green")


# Update HighestComponent levels to have space instead of underscore
df$HighestComponent <- factor(
  df$HighestComponent,
  levels = c("Component_1", "Component_2", "Component_3"),
  labels = c("Component 1", "Component 2", "Component 3")
)

# Iterate over each pair and create a plot
for(pair in component_pairs) {
  x_col <- pair[1]
  y_col <- pair[2]
  
  # Replace underscores with spaces for axis labels and titles
  x_col_label <- gsub("_", " ", x_col)
  y_col_label <- gsub("_", " ", y_col)
  
  plot <- ggplot(df, aes(x = .data[[x_col]], y = .data[[y_col]], color = HighestComponent)) +
    geom_point() +
    stat_ellipse(type = "t", level = 0.95) +
    scale_color_manual(values = comp_colors, name = "Highest Component") +
    theme_bw() +
    labs(title = paste("Scatter Plot of", x_col_label, "vs", y_col_label),
         x = x_col_label,
         y = y_col_label)
  
  # Print and save the plot
  print(plot)
  ggsave(paste0("Scatter_Plot_", gsub("_", " ", x_col), "_vs_", gsub("_", " ", y_col), ".png"), plot, width = 7, height = 5)
}