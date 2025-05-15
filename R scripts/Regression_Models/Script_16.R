# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/chris0058")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df_FDI <- read.xlsx("Cleaned_Data.xlsx", sheet = "FDI")
df_GII <- read.xlsx("Cleaned_Data.xlsx", sheet = "GII")

str(df_FDI)
str(df_GII)

library(dplyr)
library(tidyr)

# Convert FDI year columns to numeric
df_FDI[,-1] <- lapply(df_FDI[,-1], function(x) as.numeric(as.character(x)))

# Convert GII year columns to numeric (excluding non-year columns)
df_GII[,-c(1:4)] <- lapply(df_GII[,-c(1:4)], function(x) as.numeric(as.character(x)))

# Reshape FDI Data (Long Format)
df_FDI_long <- df_FDI %>%
  pivot_longer(cols = -Reference.area, names_to = "Year", values_to = "FDI") %>%
  mutate(Year = as.integer(Year))

# Reshape GII Data (Long Format) - Ensure only year columns are pivoted
df_GII_long <- df_GII %>%
  pivot_longer(cols = matches("^\\d{4}$"), names_to = "Year", values_to = "GII") %>%
  mutate(Year = as.integer(Year))

# Merge FDI and GII Data
df_final <- df_FDI_long %>%
  inner_join(df_GII_long, by = c("Reference.area" = "Economy.Name", "Year"))

# Define the list of OECD countries
oecd_countries <- c("Australia", "Austria", "Belgium", "Canada", "Chile", "Colombia", 
                    "Costa Rica", "Czechia", "Denmark", "Estonia", "Finland", "France", 
                    "Germany", "Greece", "Hungary", "Iceland", "Ireland", "Israel", 
                    "Italy", "Japan", "Latvia", "Lithuania", "Luxembourg", "Mexico", 
                    "Netherlands", "New Zealand", "Norway", "Poland", "Portugal", "Slovak Republic", 
                    "Slovenia", "Spain", "Sweden", "Switzerland", "United Kingdom", "United States")

# Filter the dataset to include only OECD countries
df_oecd <- df_final %>%
  filter(Reference.area %in% oecd_countries)

# Compute the OECD average only for these countries
df_oecd_avg <- df_oecd %>%
  group_by(Year, Indicator) %>%
  summarise(FDI = mean(FDI, na.rm = TRUE), GII = mean(GII, na.rm = TRUE), .groups = "drop") %>%
  mutate(Reference.area = "OECD Average")

# Append the new OECD average data to the main dataset
df_final <- bind_rows(df_final, df_oecd_avg)


unique(df_final$Indicator)


# Rename 'Indicator' column values
df_final <- df_final %>%
  mutate(Indicator = gsub("Global Innovation Index: ", "", Indicator))

# Check the updated unique values
unique(df_final$Indicator)

library(dplyr)
library(ggplot2)
library(broom)
library(ggpubr)
library(patchwork)
library(gridExtra)
library(readr)  # For saving tables


# Define indicators
indicators <- c(
  "Global Innovation Index",
  "Institutions index",
  "Human capital and research index",
  "Infrastructure index",
  "Market sophistication index",
  "Business sophistication index",
  "Knowledge and technology outputs index",
  "Creative outputs index"
)

# Initialize lists to store results
all_regression_results <- list()
interaction_results <- list()
z_test_results <- list()
all_plots <- list()

# Define font size adjustments
text_size <- 10
plot_theme <- theme_minimal() +
  theme(
    text = element_text(size = text_size),  # Global text size
    plot.title = element_text(size = text_size + 2, face = "bold"),  # Title size
    axis.title = element_text(size = text_size),  # Axis label size
    axis.text = element_text(size = text_size - 2),  # Tick labels
    legend.text = element_text(size = text_size - 2),
    legend.title = element_text(size = text_size)
  )

# Loop through each GII indicator
for (indicator in indicators) {
  
  # Filter data for Ireland and OECD Average
  df_reg <- df_final %>%
    filter(Indicator == indicator & Reference.area %in% c("Ireland", "OECD Average"))
  
  # Separate datasets for Ireland and OECD
  df_ireland <- df_reg %>% filter(Reference.area == "Ireland")
  df_oecd <- df_reg %>% filter(Reference.area == "OECD Average")
  
  # Ensure enough data for regression
  if (nrow(df_ireland) > 1 & nrow(df_oecd) > 1) {
    
    ### --- MODEL 1: SEPARATE REGRESSIONS FOR IRELAND AND OECD --- ###
    model_ireland <- lm(GII ~ FDI + Year, data = df_ireland)
    model_oecd <- lm(GII ~ FDI + Year, data = df_oecd)
    
    # Store Regression Results
    regression_results <- bind_rows(
      tidy(model_ireland) %>% mutate(Country = "Ireland"),
      tidy(model_oecd) %>% mutate(Country = "OECD Average")
    ) %>%
      mutate(Indicator = indicator) %>%
      select(Indicator, Country, term, estimate, std.error, statistic, p.value)
    
    all_regression_results[[indicator]] <- regression_results
    
    ### --- MODEL 2: INTERACTION MODEL (IRELAND VS. OECD COMPARISON) --- ###
    df_reg <- df_reg %>% mutate(Country = ifelse(Reference.area == "Ireland", 1, 0))
    model_interaction <- lm(GII ~ FDI * Country + Year, data = df_reg)
    
    # Store Interaction Model Results
    interaction_results[[indicator]] <- tidy(model_interaction) %>%
      mutate(Indicator = indicator)
    
    # Compute Z-Test for Difference in Coefficients
    beta_ireland <- coef(model_ireland)["FDI"]
    beta_oecd <- coef(model_oecd)["FDI"]
    se_ireland <- summary(model_ireland)$coefficients["FDI", 2]
    se_oecd <- summary(model_oecd)$coefficients["FDI", 2]
    
    z_value <- (beta_ireland - beta_oecd) / sqrt(se_ireland^2 + se_oecd^2)
    p_value <- 2 * (1 - pnorm(abs(z_value)))  # Two-tailed test
    
    z_test_results[[indicator]] <- data.frame(
      Indicator = indicator,
      Beta_Ireland = beta_ireland,
      Beta_OECD = beta_oecd,
      Z_Value = z_value,
      P_Value = p_value
    )
    
    ### --- DIAGNOSTIC PLOTS --- ###
    p1 <- ggplot(df_ireland, aes(x = FDI, y = GII)) +
      geom_point(size = 3) +
      geom_smooth(method = "lm", color = "red", se = FALSE) +
      labs(title = paste("FDI vs", indicator, "(Ireland)"),
           x = "FDI (% of GDP)", y = indicator) +
      plot_theme
    
    p2 <- ggplot(df_oecd, aes(x = FDI, y = GII)) +
      geom_point(size = 3) +
      geom_smooth(method = "lm", color = "red", se = FALSE) +
      labs(title = paste("FDI vs", indicator, "(OECD Average)"),
           x = "FDI (% of GDP)", y = indicator) +
      plot_theme
    
    qq_ireland <- ggplot(data.frame(resid = resid(model_ireland)), aes(sample = resid)) +
      stat_qq() + stat_qq_line() +
      labs(title = paste("QQ Plot (Ireland) -", indicator),
           x = "Theoretical Quantiles", y = "Residuals") +
      plot_theme
    
    resid_ireland <- ggplot(data.frame(fitted = fitted(model_ireland), resid = resid(model_ireland)),
                            aes(x = fitted, y = resid)) +
      geom_point() + geom_smooth(method = "lm", se = FALSE, color = "red") +
      labs(title = paste("Residual vs. Fitted (Ireland) -", indicator),
           x = "Fitted Values", y = "Residuals") +
      plot_theme
    
    qq_oecd <- ggplot(data.frame(resid = resid(model_oecd)), aes(sample = resid)) +
      stat_qq() + stat_qq_line() +
      labs(title = paste("QQ Plot (OECD) -", indicator),
           x = "Theoretical Quantiles", y = "Residuals") +
      plot_theme
    
    resid_oecd <- ggplot(data.frame(fitted = fitted(model_oecd), resid = resid(model_oecd)),
                         aes(x = fitted, y = resid)) +
      geom_point() + geom_smooth(method = "lm", se = FALSE, color = "red") +
      labs(title = paste("Residual vs. Fitted (OECD) -", indicator),
           x = "Fitted Values", y = "Residuals") +
      plot_theme
    
    final_plot <- (p1 + p2) / (qq_ireland + resid_ireland) / (qq_oecd + resid_oecd)
    
    # Save plots correctly
    filename <- paste0("Diagnostics_", gsub(" ", "_", indicator), ".png")
    print(final_plot)
    ggplot2::ggsave(filename, plot = final_plot, width = 10, height = 6, dpi = 300)
  }
}

# Combine all results into final tables
final_regression_table <- bind_rows(all_regression_results)
interaction_results_df <- bind_rows(interaction_results)
z_test_results_df <- bind_rows(z_test_results)
 

# Print tables in console
print("Regression Results (Separate Models):")
print(final_regression_table)

print("Interaction Model Results:")
print(interaction_results_df)

print("Z-Test Results for FDI Coefficients (Ireland vs. OECD):")
print(z_test_results_df)



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
  "Data" = df_final, 
  "Regression REsults" = final_regression_table,
  "Interaction Results" = interaction_results_df, 
  "Z-test Results" = z_test_results_df
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables.xlsx")
