# Load the required library
library(WDI)

# Define the country codes for the selected Latin American countries
countries <- c("BRA", "CHL", "MEX", "COL", "ARG", "PER", "URY", "CRI", "ECU", "PAN")

# Define the indicators to retrieve
indicators <- c(FDI = "BX.KLT.DINV.WD.GD.ZS",     # Foreign direct investment, net inflows (% of GDP)
                UP = "SP.URB.GROW",                # Urban population growth (annual %)
                TO = "NE.TRD.GNFS.ZS",             # Trade openness (exports plus imports as % of GDP)
                CO2 = "EN.GHG.CO2.PC.CE.AR5",      # CO2 emissions per capita (metric tons per capita)
                GDP = "NY.GDP.MKTP.KD.ZG",         # GDP growth (annual %)
                RE = "EG.FEC.RNEW.ZS")             # Renewable energy consumption (% of total final energy consumption)

# Retrieve the data from the World Bank's API for the years 2000 to 2022
df <- WDI(country = countries, indicator = indicators, start = 2000, end = 2022, extra = TRUE, cache = NULL)


str(df)

library(dplyr)

# Ensure the data is sorted by country and year for correct lagging
df <- df %>%
  arrange(country, year)

# Create lagged variables for FDI (1-year lag)
df <- df %>%
  group_by(country) %>%
  mutate(
    FDI_lag1 = lag(FDI, 1)  # Creates a lag of 1 year
  )

# Optional: Create more lagged variables depending on the analysis requirement
# For example, a 2-year lag for FDI if needed
df <- df %>%
  mutate(
    FDI_lag2 = lag(FDI, 2)  # Creates a lag of 2 years
  )

# Check for and handle any NA values in the newly created lagged columns
# You can choose to fill NA with 0 or exclude those rows
df <- df %>%
  filter(!is.na(FDI_lag1), !is.na(FDI_lag2))

# Drop non-essential columns if they are not needed for the analysis
df <- df %>%
  select(-iso2c, -iso3c, -status, -lastupdated, -region, -capital, -longitude, -latitude, -income, -lending)

# View the structure of the updated dataframe
str(df)

# Creating dummy variables for 'country'
country_dummies <- model.matrix(~ country - 1, data = df)  # '- 1' omits the intercept
df <- cbind(df, country_dummies)  # Bind the dummy variables to the original data frame



# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/sarahjenner")


# Optionally, write the cleaned data back to CSV for import into SPSS
write.csv(df, "cleaned_WDI_data_for_SPSS.csv", row.names = FALSE)
