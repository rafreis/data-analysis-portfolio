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
  labs(title = "Species Abundance Distribution",
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


abundance_raw_adjusted_log_series  <- abundance_raw_adjusted - 0.01

abundance_raw_adjusted_log_series


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


# Group species into octaves
octaves <- octav(abundance_data)

# Plot the octave distribution
plot(octaves, main = "Species Abundance Octave Plot", 
     xlab = "Octave (log2 abundance)", ylab = "Frequency")



# Best Model

# Assume you have already fit the models
best_model_name <- "mzsm"  # Declare the best model
best_model_result <- fit_results[[which(sapply(fit_results, function(x) x$model) == best_model_name)]]$fit

# Generate density plot comparing the model and data
plot(density(abundance_data), col = "blue", lwd = 2, 
     main = paste("Density Comparison - Model:", best_model_name), 
     xlab = "Abundance", ylab = "Density")

# Overlay the fitted model density
curve(dmzsm(round(x), theta = best_model_result@fullcoef["theta"], 
            J = best_model_result@fullcoef["J"]), 
      add = TRUE, col = "red", lwd = 2)

# Add legend
legend("topright", legend = c("Data Density", "Model Density"), 
       col = c("blue", "red"), lwd = 2)


# Generate QQ plot
qqsad(best_model_result, main = "QQ Plot - mzsm Model")
# Generate PP plot
ppsad(best_model_result, main = "PP Plot - mzsm Model")


# Print model summary
print(summary(best_model_result))

# Extract AIC and log-likelihood
aic_value <- AIC(best_model_result)
log_likelihood <- logLik(best_model_result)

cat("AIC:", aic_value, "\n")
cat("Log-Likelihood:", log_likelihood, "\n")


library(diptest)

# Assuming abundance_data contains your species abundance data
dip_result <- dip.test(abundance_data)
print(dip_result)

# Assuming abundance_data contains your species abundance data
dip_result <- dip.test(abundance_raw_logtransformed)
print(dip_result)


# DORNELAS SCRIPTS

abundance_transformed <- log1p(abundance_raw_adjusted_log_series)
abundance_transformed <- round(abundance_transformed)
abundance_transformed <- unname(abundance_transformed)

library(stats) # For optimization functions

max.abund <- max(abundance_transformed) # Maximum observed abundance

# Function to calculate probabilities using Pielou's formula with vector handling
mkpielou <- function(params) {
  m <- params[1]
  sig <- params[2]
  max_abund <- max(abundance_transformed, na.rm = TRUE)
  
  probs <- rep(NA, max_abund)
  
  integrand <- function(lam, r) {
    valid_lam <- lam > 0 & !is.nan(lam)
    result <- numeric(length(lam))
    result[!valid_lam] <- 0  # Set invalid lam values to 0
    
    # Calculate for valid lam values
    result[valid_lam] <- exp(-lam[valid_lam] + r * log(lam[valid_lam]) - (log(lam[valid_lam] / m)^2) / (2 * sig^2)) /
      pmax(lam[valid_lam], 1e-6)
    return(result)
  }
  
  for (r in 1:max_abund) {
    cat("Processing r =", r, "m =", m, "sig =", sig, "\n")
    tryCatch({
      if (r <= 10) {
        result <- integrate(integrand, lower = 0.01, upper = Inf, r = r)
        if (result$message == "OK") {
          probs[r] <- result$value
          cat("Integration successful for r =", r, "result =", result$value, "\n")
        } else {
          warning(paste("Integration failed at r =", r, ":", result$message))
          probs[r] <- NA
        }
      } else {
        probs[r] <- 1 / (sqrt(2 * pi) * sig * r) * exp(-log(r / m)^2 / (2 * sig^2))
        cat("Bulmer's approximation used for r =", r, "prob =", probs[r], "\n")
      }
    }, warning = function(w) {
      warning(paste("Integration warning at r =", r, ":", w$message))
      probs[r] <- NA
    }, error = function(e) {
      warning(paste("Integration error at r =", r, ":", e$message))
      probs[r] <- NA
    })
  }
  
  cat("Computed probs:", probs, "\n")
  return(probs)
}

# Function to calculate p0 with vector handling and range constraint
getp0 <- function(params) {
  m <- params[1]
  sig <- params[2]
  
  integrand <- function(lam) {
    valid_lam <- lam > 0 & !is.nan(lam)
    result <- numeric(length(lam))
    result[!valid_lam] <- 0  # Set invalid lam values to 0
    
    # Calculate for valid lam values
    result[valid_lam] <- exp(-lam[valid_lam] - (log(lam[valid_lam] / m)^2) / (2 * sig^2)) /
      pmax(lam[valid_lam], 1e-6)
    return(result)
  }
  
  tryCatch({
    result <- integrate(integrand, lower = 0.01, upper = Inf)
    if (result$message == "OK") {
      cat("p0 integration result:", result$value, "\n")
      return(min(max(result$value, 0), 0.999))  # Constrain p0 within the valid range
    } else {
      warning("Integration for p0 failed:", result$message)
      return(NA)
    }
  }, warning = function(w) {
    warning("Integration warning for p0:", w$message)
    return(NA)
  }, error = function(e) {
    warning("Integration error for p0:", e$message)
    return(NA)
  })
}

# Log-likelihood function with detailed logging
loglike.1 <- function(params) {
  probs <- mkpielou(params)
  p0 <- getp0(params)
  
  if (is.na(p0) || all(is.na(probs))) {
    cat("Invalid p0 or probs\n")
    return(Inf)
  }
  
  totprob <- probs / (1 - p0)
  
  valid_probs <- totprob[abundance_transformed]
  if (any(is.na(valid_probs) | valid_probs <= 0)) {
    cat("Non-finite or zero probabilities encountered\n")
    return(Inf)
  }
  
  return(-sum(log(valid_probs)))
}


# Initial parameters
initial_params <- c(mean(log(abundance_transformed)), sd(log(abundance_transformed)))

# Run optimization
fit <- optim(initial_params, loglike.1, method = "L-BFGS-B", lower = c(0.01, 0.01))


# Output results
cat("Optimization result:", fit$par, "\n")




# WITH RAW COUNTS

abundance_raw_adjusted_log_series <- unname(abundance_raw_adjusted_log_series)
abundance_raw_adjusted_log_series <- (log(abundance_raw_adjusted_log_series))
abundance_raw_adjusted_log_series

library(stats) # For optimization functions

max.abund <- max(abundance_raw_adjusted_log_series) # Maximum observed abundance

# Updated mkpielou function
mkpielou <- function(params) {
  m <- params[1]
  sig <- params[2]
  
  probs <- rep(NA, max.abund)
  
  # Corrected integrand function with normalization
  integrand <- function(lam, r) {
    1 / (sqrt(2 * pi) * sig * gamma(r + 1)) * 
      exp(-lam + r * log(lam) - (log(lam / m)^2) / (2 * sig^2)) / lam
  }
  
  for (r in 1:max.abund) {
    cat("Processing r =", r, "m =", m, "sig =", sig, "\n")
    tryCatch({
      if (r <= 10) {
        result <- integrate(integrand, lower = 0.01, upper = Inf, r = r)
        if (result$message == "OK") {
          probs[r] <- result$value
          cat("Integration successful for r =", r, "result =", result$value, "\n")
        } else {
          warning(paste("Integration failed at r =", r, ":", result$message))
          probs[r] <- NA
        }
      } else {
        # Bulmer's approximation for larger r values
        probs[r] <- 1 / (sqrt(2 * pi) * sig * r) * exp(-log(r / m)^2 / (2 * sig^2))
      }
    }, warning = function(w) {
      warning(paste("Integration warning at r =", r, ":", w$message))
      probs[r] <- NA
    }, error = function(e) {
      warning(paste("Integration error at r =", r, ":", e$message))
      probs[r] <- NA
    })
  }
  
  return(probs)
}

# Function to calculate p0
getp0 <- function(params) {
  m <- params[1]
  sig <- params[2]
  
  integrand <- function(lam) {
    exp(-lam - (log(lam / m)^2) / (2 * sig^2)) / lam
  }
  
  result <- tryCatch({
    integrate(integrand, lower = 0.01, upper = Inf)$value
  }, error = function(e) {
    warning("Integration error for p0:", e$message)
    return(NA)
  })
  
  return(min(max(result, 0), 0.999))
}

# Log-likelihood for one Poisson lognormal distribution
loglike.1 <- function(params) {
  probs <- mkpielou(params)
  p0 <- getp0(params)
  
  if (is.na(p0) || all(is.na(probs))) return(Inf)
  
  totprob <- probs / (1 - p0)
  valid_probs <- totprob[abundance_raw_adjusted_log_series]
  
  if (any(is.na(valid_probs) | valid_probs <= 0)) return(Inf)
  
  return(-sum(log(valid_probs)))
}

# Log-likelihood for two distributions
loglike.2 <- function(params) {
  probs1 <- mkpielou(params[1:2])
  p01 <- getp0(params[1:2])
  probs2 <- mkpielou(params[3:4])
  p02 <- getp0(params[3:4])
  
  mixture_prob <- params[5]
  totprob <- (1 - mixture_prob) * probs2 / (1 - p02) + mixture_prob * probs1 / (1 - p01)
  
  return(-sum(log(totprob[abundance_raw_adjusted_log_series]), na.rm = TRUE))
}

# Log-likelihood for three distributions
loglike.3 <- function(params) {
  probs1 <- mkpielou(params[1:2])
  p01 <- getp0(params[1:2])
  probs2 <- mkpielou(params[3:4])
  p02 <- getp0(params[3:4])
  probs3 <- mkpielou(params[5:6])
  p03 <- getp0(params[5:6])
  
  f1 <- params[7]
  f2 <- (1 - params[7]) * params[8]
  f3 <- (1 - params[7]) * (1 - params[8])
  
  totprob <- f3 * probs3 / (1 - p03) + f2 * probs2 / (1 - p02) + f1 * probs1 / (1 - p01)
  
  return(-sum(log(totprob[abundance_raw_adjusted_log_series]), na.rm = TRUE))
}

# Log-likelihood for four distributions
loglike.4 <- function(params) {
  probs1 <- mkpielou(params[1:2])
  p01 <- getp0(params[1:2])
  probs2 <- mkpielou(params[3:4])
  p02 <- getp0(params[3:4])
  probs3 <- mkpielou(params[5:6])
  p03 <- getp0(params[5:6])
  probs4 <- mkpielou(params[7:8])
  p04 <- getp0(params[7:8])
  
  f1 <- params[9]
  f2 <- (1 - params[9]) * params[10]
  f3 <- (1 - params[9]) * (1 - params[10]) * params[11]
  f4 <- (1 - params[9]) * (1 - params[10]) * (1 - params[11])
  
  totprob <- f4 * probs4 / (1 - p04) + f3 * probs3 / (1 - p03) + f2 * probs2 / (1 - p02) + f1 * probs1 / (1 - p01)
  
  return(-sum(log(totprob[abundance_raw_adjusted_log_series]), na.rm = TRUE))
}

# Fit models
initial_params_1 <- c(mean(abundance_raw_adjusted_log_series), sd(abundance_raw_adjusted_log_series))
fit1 <- optim(initial_params_1, loglike.1, method = "L-BFGS-B", lower = c(0.01, 0.01))

initial_params_2 <- c(initial_params_1, initial_params_1, 0.5)
fit2 <- optim(initial_params_2, loglike.2, method = "L-BFGS-B", lower = c(0.01, 0.01, 0.01, 0.01, 0.01), upper = c(10, 50, 10, 50, 1))

initial_params_3 <- c(initial_params_2, initial_params_1, 0.33, 0.33)
fit3 <- optim(initial_params_3, loglike.3, method = "L-BFGS-B", lower = c(rep(0.01, 6), 0.01, 0.01))

initial_params_4 <- c(initial_params_3, initial_params_1, 0.25, 0.25, 0.25)
fit4 <- optim(initial_params_4, loglike.4, method = "L-BFGS-B", lower = c(rep(0.01, 8), 0.01, 0.01, 0.01))

# Compare models using AIC
aic1 <- 2 * length(fit1$par) + 2 * fit1$value
aic2 <- 2 * length(fit2$par) + 2 * fit2$value
aic3 <- 2 * length(fit3$par) + 2 * fit3$value
aic4 <- 2 * length(fit4$par) + 2 * fit4$value

cat("AIC for Single PLN Model:", aic1, "\n")
cat("AIC for Two-Distribution Model:", aic2, "\n")
cat("AIC for Three-Distribution Model:", aic3, "\n")
cat("AIC for Four-Distribution Model:", aic4, "\n")