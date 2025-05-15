# Convert 'Group.Name' to a factor with a specified order
group_order <- c("Placebo", "1g Primacolll", "2.5g Primacoll", "5g Primacoll", "10g Primacoll", "Verisol", "Vital Protein")
group_labels <- c("Placebo", "1g Primacolll", "2.5g Primacoll", "5g Primacoll", "10g Primacoll", "Porcine Collagen", "Bovine Collagen")

df$Group.Name <- factor(df$Group.Name, levels = group_order)

# Generate separate boxplots for each dependent variable
for (dv in dvs) {
  plot <- ggplot(df, aes(x = Timepoint, y = !!sym(dv), fill = Group.Name)) +
    geom_boxplot() +
    scale_fill_manual(values = scales::hue_pal()(length(group_order)), 
                      breaks = group_order, 
                      labels = group_labels) +
    labs(
      title = paste("Boxplot of", dv),
      x = "Timepoint",
      y = "Measurement Value"
    ) +
    theme(
      panel.background = element_rect(fill = "white"),
      panel.grid.major = element_line(colour = "grey92"),
      panel.grid.minor = element_line(colour = "grey92")
    )
  
  # Show plot
  print(plot)
  
  # Save plot to a file if needed
  ggsave(filename = paste0("boxplot_", dv, ".png"), plot = plot)
}



# Filtered groups
filtered_groups <- c("Placebo", "5g Primacoll", "10g Primacoll", "Vital Protein")  # Note: Using original name "Vital Protein" for "Bovine Collagen"

# Generate separate boxplots for specified dependent variables and groups
for (dv in dvs) {
  filtered_df <- df %>% filter(Group.Name %in% filtered_groups)
  
  plot <- ggplot(filtered_df, aes(x = Timepoint, y = !!sym(dv), fill = Group.Name)) +
    geom_boxplot() +
    scale_fill_manual(values = scales::hue_pal()(length(filtered_groups)), 
                      breaks = filtered_groups, 
                      labels = c("Placebo", "5g Primacoll", "10g Primacoll", "Bovine Collagen")) +  # Updating labels here
    labs(
      title = paste("Filtered Boxplot of", dv),
      x = "Timepoint",
      y = "Measurement Value"
    ) +
    theme(
      panel.background = element_rect(fill = "white"),
      panel.grid.major = element_line(colour = "grey92"),
      panel.grid.minor = element_line(colour = "grey92")
    )
  
  # Show plot
  print(plot)
  
  # Save plot to a file with a different name
  ggsave(filename = paste0("filtered_boxplot_", dv, ".png"), plot = plot)
}

