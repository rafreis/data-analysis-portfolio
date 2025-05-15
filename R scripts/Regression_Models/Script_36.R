# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/lauren_mahoney")

# Load the openxlsx library
library(openxlsx)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("AbundanceMeta_Mahoney.xlsx", sheet = "TotalAbundance")

# Get rid of special characters

names(df) <- gsub(" ", "_", trimws(names(df)))
names(df) <- gsub("\\s+", "_", trimws(names(df), whitespace = "[\\h\\v\\s]+"))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))
names(df) <- gsub("\\$", "", names(df))
names(df) <- gsub("\\+", "_", names(df))

# Trim all values

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))



# Trim all values and convert to lowercase

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
    # Convert character strings to lowercase
    x <- tolower(x)
  }
  return(x)
}))



str(df)

# Load necessary libraries
library(ggplot2)
library(MASS)
library(mixtools)
library(reshape2)

# Assuming abundance_raw is your calculated total species abundance
abundance_raw <- colSums(df[, -c(1, 2)])  # Sum of species abundances, excluding non-numeric columns

# Define breaks and check maximum value to ensure coverage
breaks <- c(0, 2, 4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, Inf)
bin_labels <- c("1", "2-3", "4-7", "8-15", "16-31", "32-63", "64-127", "128-255",
                "256-511", "512-1023", "1024-2047", "2048-4095", "4096+")

# Function to assign each abundance value to a bin
assign_bins <- function(x) {
  findInterval(x, vec = breaks, rightmost.closed = TRUE)
}

# Apply binning function
bin_data <- assign_bins(abundance_raw + 1)  # Add 1 to avoid the 0 issue

# Convert numeric bins to factor with custom labels for plotting
bin_data_factor <- factor(bin_data, labels = bin_labels)

# Prepare data frame for ggplot
plot_data <- data.frame(Bin = bin_data_factor)

# Plotting
ggplot(plot_data, aes(x = Bin)) +
  geom_bar(fill = "gray", color = "black") +  # Use geom_bar for pre-binned data
  labs(title = "Species Abundance Distribution (Log2 Bins)",
       x = "Species Abundance (in Log2 classes)",
       y = "Number of Species") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))  # Rotate x labels for better readability

# Save the plot if needed
ggsave("SAD_Log2_Bins.pdf", width=10, height=6)


library(MASS)         # For fitting log-normal distributions
library(mixtools)     # For Gaussian Mixture Models
library(VGAM)         # For fitting log-series using pospoisson
library(fitdistrplus) # Enhanced distribution fitting tools

# Load your data
abundance_raw <- colSums(df[, -c(1, 2)])  

# Adding a small constant to all counts to ensure no zero counts for log transformations
abundance_raw_adjusted <- abundance_raw + 1.01

# 1. Fit Log-normal Distribution
log_normal_fit <- fitdistr(abundance_raw_adjusted, "lognormal")
print(log_normal_fit)

# Extract parameters from the log-normal fit
plot_lognormal_model <- function(data, fit) {
  log_normal_meanlog <- fit$estimate["meanlog"]
  log_normal_sdlog <- fit$estimate["sdlog"]
  
  # Store the plot within the function
  plot_obj <- function() {
    plot(density(data), col = "blue", lwd = 2, 
         main = "Log-normal Model vs Data Density", 
         xlab = "Log-transformed Abundance", ylab = "Density")
    curve(dlnorm(x, meanlog = log_normal_meanlog, sdlog = log_normal_sdlog), 
          add = TRUE, col = "red", lwd = 2)
    legend("topright", legend = c("Data Density", "Log-normal Model Density"), 
           col = c("blue", "red"), lwd = 2)
  }
  
  return(plot_obj)
}
# Call the function and store the plot object
plot_obj <- plot_lognormal_model(abundance_raw, log_normal_fit)

# Print the plot
plot_obj()

abundance_raw_adjusted_log_series  <- abundance_raw_adjusted - 0.01

abundance_raw_adjusted_log_series

# 2. Fit Log-series Distribution
#log_series_fit <- vglm(as.integer(as.integer(abundance_raw_adjusted_log_series)) ~ 1, pospoisson(), trace = TRUE)

# Trying a simple GLM with Poisson distribution
glm_fit <- glm(abundance_raw_adjusted_log_series ~ 1, family = poisson(link = "log"))
summary(glm_fit)


# Function to create the plot object for GLM with Poisson distribution
plot_glm_poisson_model <- function(data, glm_fit) {
  # Extract the model parameter (intercept)
  lambda_hat <- exp(coef(glm_fit)[1])  # The rate parameter (mean) for Poisson
  
  # Create a plotting function
  plot_obj <- function() {
    plot(density(data), col = "blue", lwd = 2, 
         main = "GLM Poisson Model vs Data Density", 
         xlab = "Log-transformed Abundance", ylab = "Density")
    
    # Overlay the Poisson density (use exp(data) since the data is log-transformed)
    curve(dpois(round(exp(x)), lambda = lambda_hat), add = TRUE, col = "red", lwd = 2)
    
    legend("topright", legend = c("Data Density", "GLM Poisson Model Density"), 
           col = c("blue", "red"), lwd = 2)
  }
  
  return(plot_obj)
}

# Call the function and store the plot object
plot_obj <- plot_glm_poisson_model(abundance_raw, glm_fit)

# Print the plot
plot_obj()


# Transform the adjusted data for fitting distributions that require non-zero data
abundance_raw_logtransformed <- log(abundance_raw_adjusted)

abundance_raw_logtransformed

library(mclust)

# Fit models with different numbers of components (e.g., 1 to 4)
gaussian_mix_models <- list()
for (k in 1:4) {
  gaussian_mix_models[[k]] <- Mclust(abundance_raw_logtransformed, G = k)
}

# Extract parameters for each model
all_parameters <- lapply(gaussian_mix_models, function(model) {
  summary(model, parameters = TRUE)
})

# Print parameters for each model
for (i in seq_along(all_parameters)) {
  cat("Model with", i, "component(s):\n")
  print(all_parameters[[i]])
  cat("\n")
}

# Plot density for each model
par(mfrow = c(2, 2))  # Arrange plots in a 2x2 grid

for (k in 1:4) {
  plot(gaussian_mix_models[[k]], what = "density", 
       main = paste("Gaussian Mixture Model with", k, "Component(s)"), 
       xlab = "Log-transformed Abundance", ylab = "Density")
  lines(density(abundance_raw_logtransformed), col = "blue", lwd = 2)  # Overlay data density
  legend("topright", legend = c("Model Density", "Data Density"), 
         col = c("black", "blue"), lwd = 2)
}


# POisson Lognormal Model
library(poilog)

# Fit a Poisson lognormal model using poilogMLE, which directly fits the Poisson log-normal model to count data
# Make sure the data is appropriate for the function, generally requiring non-zero integer counts
abundance_raw_adjusted_log_series <- as.integer(abundance_raw_adjusted_log_series)
poisson_lognormal_model <- poilogMLE(abundance_raw_adjusted_log_series)

# Print the model fit summary
print(poisson_lognormal_model)
str(poisson_lognormal_model)

# Function to plot the Poisson-lognormal model density and data density
plot_poisson_lognormal_model <- function(data, model_fit) {
  # Extract parameters from the model
  mu_hat <- model_fit$par[1]     # Estimated mean (mu)
  sigma_hat <- model_fit$par[2]  # Estimated standard deviation (sigma)
  
  # Create the plot function
  plot_obj <- function() {
    plot(density(log(data)), col = "blue", lwd = 2, 
         main = "Poisson-lognormal Model vs Data Density", 
         xlab = "Log-transformed Abundance", ylab = "Density")
    
    # Overlay the Poisson-lognormal density (use dpoilog and log-transformed values)
    curve(dpoilog(exp(x), mu = mu_hat, sig = sigma_hat), add = TRUE, col = "red", lwd = 2)
    
    legend("topright", legend = c("Data Density", "Poisson-lognormal Model Density"), 
           col = c("blue", "red"), lwd = 2)
  }
  
  return(plot_obj)
}

# Call the function and store the plot object
plot_obj <- plot_poisson_lognormal_model(abundance_raw, poisson_lognormal_model)

# Print the plot
plot_obj()


# AIC AND BIC COMPARISON

# Log-normal model fit details
logL <- log_normal_fit$loglik
n <- log_normal_fit$n
k <- length(log_normal_fit$estimate)
AIC_log_normal <- 2 * k - 2 * logL
BIC_log_normal <- log(n) * k - 2 * logL
print(paste("Calculated AIC LogNormal:", AIC_log_normal))
print(paste("Calculated BIC LogNormal:", BIC_log_normal))

# GLM Log-Series model fit coefficients
AIC_GLMLogseries <- glm_fit$aic
# Automatically reported AIC
print(paste("AIC LogSeries:", glm_fit$aic))

# Calculate BIC manually for GLM
n <- length(glm_fit$fitted.values)  # number of observations
k <- length(glm_fit$coefficients)   # number of parameters
logL <- glm_fit$deviance * -0.5     # rough estimate of the log-likelihood
BIC_GLMLogseries <- log(n) * k - 2 * logL
print(paste("BIC LogSeries:", BIC_GLMLogseries))


# Gaussian Mixture model fit coefficients
# Extract and print GMM model details
lapply(gaussian_mix_models, function(model) {
  cat("Gaussian Mixture Model for", model$G, "components:\n")
  cat("Gaussian Mixture BIC:", model$bic, "\n")
  print(summary(model)$parameters)
  cat("\n")
})


# Poisson Log-normal model details from sads package
cat("Poisson Log-normal model details:\n")
str(poisson_lognormal_model)

# Extract log-likelihood value
logLval <- poisson_lognormal_model$logLval

# Number of parameters ('mu' and 'sig')
num_parameters <- length(poisson_lognormal_model$par)

# Check the length of the series
n <- length(abundance_raw_adjusted_log_series)

# Calculate AIC and BIC using the extracted log likelihood
AIC_Poisson_LogNormal <- -2 * logLval + 2 * num_parameters
BIC_Poisson_LogNormal <- -2 * logLval + num_parameters * log(n)

# Print results
print(paste("Calculated AIC for Poisson LogNormal:", AIC_Poisson_LogNormal))
print(paste("Calculated BIC for Poisson LogNormal:", BIC_Poisson_LogNormal))


# USING SADS PACKAGE

# Load the package
library(sads)

# Define the abundance data (make sure data is properly prepared)
abundance_data <- abundance_raw_adjusted_log_series  # Replace with your actual data

# List of models to fit
models_to_fit <- c("bs", "exp", "gamma", "geom", "lnorm", "ls", "mzsm", "nbinom", 
                   "pareto", "poilog", "power", "powbend", "volkov", "weibull")

# Function to fit a single model and return AIC
fit_sad_model <- function(model_name, data) {
  tryCatch({
    fit <- fitsad(data, sad = model_name)
    aic_value <- AIC(fit)  # Extract AIC value
    return(list(model = model_name, fit = fit, AIC = aic_value))
  }, error = function(e) {
    return(list(model = model_name, fit = NULL, AIC = NA))  # Handle errors if fitting fails
  })
}

# Fit all models and collect results
fit_results <- lapply(models_to_fit, fit_sad_model, data = abundance_data)
str(fit_results)
# Create a summary data frame with model names and AIC values
fit_summary <- data.frame(
  Model = sapply(fit_results, function(x) x$model),
  AIC = sapply(fit_results, function(x) x$AIC)
)

# Sort models by AIC value (smaller is better)
fit_summary <- fit_summary[order(fit_summary$AIC), ]

# Display the results
print(fit_summary)

# Function to plot a fitted model and data density
plot_model_comparison <- function(data, fit_result, model_name) {
  plot(density(data), col = "blue", lwd = 2, 
       main = paste("Model:", model_name), 
       xlab = "Abundance", ylab = "Density")
  
  # Generate the curve using the fitted parameters for each model
  if (model_name == "mzsm") {
    curve(dmzsm(round(x), theta = fit_result@fullcoef["theta"], 
                J = fit_result@fullcoef["J"]), add = TRUE, col = "red", lwd = 2)
  } else if (model_name == "ls") {
    curve(dls(round(x), N = fit_result@fullcoef["N"], 
              alpha = fit_result@fullcoef["alpha"]), add = TRUE, col = "red", lwd = 2)
  } else if (model_name == "volkov") {
    curve(dvolkov(round(x), theta = fit_result@fullcoef["theta"], 
                  m = fit_result@fullcoef["m"], J = fit_result@fullcoef["J"]), 
          add = TRUE, col = "red", lwd = 2)
  } else if (model_name == "poilog") {
    curve(dpoilog(round(x), mu = fit_result@fullcoef["mu"], sig = fit_result@fullcoef["sig"]),
          from = min(data), to = max(data), add = TRUE, col = "red", lwd = 2)
    
  }
  
  # Add legend
  legend("topright", legend = c("Data Density", "Model Density"), 
         col = c("blue", "red"), lwd = 2)
}

# Top 4 models to plot
top_models <- c("mzsm", "ls", "volkov", "poilog")

# Filter the top models and their results
top_results <- fit_results[which(sapply(fit_results, function(x) x$model) %in% top_models)]

# Plot each model in a 2x2 grid
par(mfrow = c(2, 2))  # Arrange plots in a 2x2 layout
for (result in top_results) {
  if (!is.null(result$fit)) {
    plot_model_comparison(abundance_raw_adjusted, result$fit, result$model)
  }
}




