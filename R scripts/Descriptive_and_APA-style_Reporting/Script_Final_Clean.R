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




# GAUSSIAN OCTAVE PLOTS

library(mclust)
library(ggplot2)
library(dplyr)

# Fit Gaussian mixture models with different numbers of components (1 to 4)
gaussian_mix_models <- list()
for (k in 1:4) {
  gaussian_mix_models[[k]] <- Mclust(abundance_raw_logtransformed, G = k)
}

# Function to compute predicted octave frequencies from a fitted GMM
get_gmm_octave_freqs <- function(model, abundance_data) {
  # Generate a sequence of log-transformed abundances
  log_abund_seq <- seq(min(abundance_raw_logtransformed), max(abundance_raw_logtransformed), length.out = 1000)
  
  # Extract model parameters
  mu <- model$parameters$mean    # Component means
  sigma <- sqrt(model$parameters$variance$sigmasq)  # Standard deviations
  weights <- model$parameters$pro  # Mixing proportions
  
  # Compute the Gaussian Mixture density
  gmm_density <- rep(0, length(log_abund_seq))
  for (j in 1:length(mu)) {
    gmm_density <- gmm_density + weights[j] * dnorm(log_abund_seq, mean = mu[j], sd = sigma[j])
  }
  
  # Transform densities back to original scale
  original_scale_abund <- exp(log_abund_seq)  # Convert log-abundance back to original abundance
  
  # Compute octaves
  predicted_data <- data.frame(
    abundance = original_scale_abund,
    predicted_density = gmm_density
  )
  predicted_data$octave <- floor(log2(predicted_data$abundance))
  
  # Aggregate predicted density by octave bins
  predicted_octave <- predicted_data %>%
    group_by(octave) %>%
    summarise(Predicted = sum(predicted_density, na.rm = TRUE), .groups = 'drop')
  
  return(predicted_octave)
}

# Process observed octaves
abundance_data <- data.frame(count = exp(abundance_raw_logtransformed))  # Convert back to original scale
abundance_data$octave <- floor(log2(abundance_data$count))

observed_octave <- abundance_data %>%
  group_by(octave) %>%
  summarise(Freq = n(), .groups = 'drop')

# Generate octave plots for each model
for (k in 1:4) {
  predicted_octave <- get_gmm_octave_freqs(gaussian_mix_models[[k]], abundance_data)
  
  # Merge observed and predicted octave frequencies
  octave_data <- left_join(observed_octave, predicted_octave, by = "octave")
  octave_data$Upper_Bound <- 2^(octave_data$octave)
  
  # Generate the plot
  plot_gmm <- ggplot(octave_data, aes(x = factor(Upper_Bound))) +
    geom_bar(aes(y = Freq), stat = "identity", fill = "darkgray", alpha = 0.7) +
    geom_line(aes(y = Predicted, group = 1), color = "black", size = 1.0) +
    scale_x_discrete(name = "Octave",
                     labels = function(x) sapply(x, function(i) format(i, scientific = FALSE))) +
    labs(title = paste("Gaussian Mixture Model (", k, " Components)", sep = ""),
         y = "Number of Species") +
    theme_minimal() +
    theme(text = element_text(size = 12),
          axis.title = element_text(size = 14),
          plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
          legend.position = "bottom",
          legend.title = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank())
  
  # Print the plot
  print(plot_gmm)
  
  # Save the plot
  ggsave(paste0("GMM_", k, "_Components_Octave.png"), plot_gmm, width = 10, height = 6, units = "in")
}
























# USING SADS PACKAGE

# Load the package
library(sads)

# Define the abundance data (make sure data is properly prepared)
abundance_data <- abundance_raw_adjusted  # Replace with your actual data

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









library(sads)
library(ggplot2)
library(dplyr)

# Assuming 'lnorm' model fits best and 'abundance_raw_adjusted' is your species abundance data
best_fit_lnorm <- fitsad(abundance_raw_adjusted, sad = "lnorm")

# Calculate the octave for each species count
abundance_data_lnorm <- data.frame(count = abundance_raw_adjusted)
abundance_data_lnorm$octave <- floor(log2(abundance_data_lnorm$count))

# Aggregate observed data by octave
observed_octave_lnorm <- abundance_data_lnorm %>%
  group_by(octave) %>%
  summarise(Freq = n(), .groups = 'drop')

# Generate predicted probabilities using the log-normal PDF
meanlog_lnorm <- best_fit_lnorm@coef["meanlog"]
sdlog_lnorm <- best_fit_lnorm@coef["sdlog"]

meanlog_lnorm
sdlog_lnorm

max_abund_lnorm <- max(abundance_raw_adjusted)
abund_seq_lnorm <- seq(1, max_abund_lnorm, by = 1)
predicted_freqs_lnorm <- dlnorm(abund_seq_lnorm, meanlog = meanlog_lnorm, sdlog = sdlog_lnorm)
predicted_counts_lnorm <- predicted_freqs_lnorm * length(abundance_raw_adjusted)  # Adjust by total number of observations

# Create a dataframe for predicted counts
predicted_data_lnorm <- data.frame(count = abund_seq_lnorm, predicted = predicted_counts_lnorm)
predicted_data_lnorm$octave <- floor(log2(predicted_data_lnorm$count))

# Aggregate predicted data by octave
predicted_octave_lnorm <- predicted_data_lnorm %>%
  group_by(octave) %>%
  summarise(Predicted = sum(predicted, na.rm = TRUE), .groups = 'drop')

# Merge observed and predicted data for plotting
octave_data_lnorm <- left_join(observed_octave_lnorm, predicted_octave_lnorm, by = "octave")
octave_data_lnorm$Upper_Bound <- 2^(octave_data_lnorm$octave)

# Define the plot
plot_lognormal <- ggplot(octave_data_lnorm, aes(x = factor(Upper_Bound))) +
  geom_bar(aes(y = Freq), stat = "identity", fill = "darkgray", alpha = 0.7) +
  geom_line(aes(y = Predicted, group = 1), color = "black", size = 1.0) +
  scale_x_discrete(name = "Octave",
                   labels = function(x) sapply(x, function(i) format(i, scientific = FALSE))) +
  labs(title = "Species Abundance Octave Plot",
       y = "Number of Species") +
  theme_minimal() +
  theme(text = element_text(size = 12),
        axis.title = element_text(size = 14),
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        legend.position = "bottom",
        legend.title = element_blank(),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

print(plot_lognormal)

# Save the plot
ggsave("LogNormal_PredSAD.png", plot_lognormal, width = 10, height = 6, units = "in")


# Calculate AIC for the log-normal model
aic_lnorm <- AIC(best_fit_lnorm)

# Print the AIC value
print(aic_lnorm)



  # DORNELAS SCRIPTS
  
  library(stats) # For optimization functions
  
  max.abund <- max(abundance_raw) # Maximum observed abundance
  abundance_raw
                   
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
    valid_probs <- totprob[abundance_raw]
    
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
    
    return(-sum(log(totprob[abundance_raw]), na.rm = TRUE))
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
    
    return(-sum(log(totprob[abundance_raw]), na.rm = TRUE))
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
    
    return(-sum(log(totprob[abundance_raw]), na.rm = TRUE))
  }
  
  # Fit models
  initial_params_1 <- c(mean(abundance_raw), sd(abundance_raw))
  fit1 <- optim(initial_params_1, loglike.1, method = "L-BFGS-B", lower = c(0.01, 0.1))
  
  
  
  
  # Additional optimization for models with two or more components
  
    optimize_model <- function(params, loglikelihood, lower_bounds, upper_bounds = NULL, max_iterations = 100) {
    tryCatch({
      optim_result <- optim(
        par = params,
        fn = loglikelihood,
        method = "L-BFGS-B",
        lower = lower_bounds,
        upper = upper_bounds,
        control = list(maxit = max_iterations)
      )
      if (optim_result$convergence != 0) {
        cat("Optimization did not converge: ", optim_result$message, "\n")
        return(NULL)
      }
      return(optim_result)
    }, warning = function(w) {
      cat("Warning during optimization: ", w$message, "\n")
      return(NULL)
    }, error = function(e) {
      cat("Error during optimization: ", e$message, "\n")
      return(NULL)
    })
  }
  
  # Adjust the bounds and initial parameters for each model
  lower_bounds_sigma = 0.1  # More realistic lower bounds for sigma to prevent it from becoming too small
  
  # Two-component model
  fit2 <- optimize_model(
    params = initial_params_2,
    loglikelihood = loglike.2,
    lower_bounds = c(0.01, lower_bounds_sigma, 0.01, lower_bounds_sigma, 0.01),
    upper_bounds = c(10, 50, 10, 50, 1),
    max_iterations = 100
  )
  
  # Three-component model
  fit3 <- optimize_model(
    params = initial_params_3,
    loglikelihood = loglike.3,
    lower_bounds = c(rep(0.01, 6), lower_bounds_sigma, lower_bounds_sigma, 0.01, 0.01),
    max_iterations = 100
  )
  
  # Four-component model
  fit4 <- optimize_model(
    params = initial_params_4,
    loglikelihood = loglike.4,
    lower_bounds = c(rep(0.01, 8), lower_bounds_sigma, lower_bounds_sigma, lower_bounds_sigma, 0.01, 0.01, 0.01),
    max_iterations = 100
  )
  
  # Check if fits were successful and handle each accordingly
  if (is.null(fit2)) {
    cat("Model 2 fitting was halted due to issues.\n")
  } else {
    cat("Model 2 fitted successfully.\n")
  }
  
  if (is.null(fit3)) {
    cat("Model 3 fitting was halted due to issues.\n")
  } else {
    cat("Model 3 fitted successfully.\n")
  }
  
  if (is.null(fit4)) {
    cat("Model 4 fitting was halted due to issues.\n")
  } else {
    cat("Model 4 fitted successfully.\n")
  }
  
  
  # Compare models using AIC
  aic1 <- 2 * length(fit1$par) + 2 * fit1$value
  aic2 <- 2 * length(fit2$par) + 2 * fit2$value
  aic3 <- 2 * length(fit3$par) + 2 * fit3$value
  aic4 <- 2 * length(fit4$par) + 2 * fit4$value
  
  cat("AIC for Single PLN Model:", aic1, "\n")
  cat("AIC for Two-Distribution Model:", aic2, "\n")
  cat("AIC for Three-Distribution Model:", aic3, "\n")
  cat("AIC for Four-Distribution Model:", aic4, "\n")
  
  
  print_model_results <- function() {
    # Helper function to display probabilities
    display_probs <- function(params, model_name) {
      probs <- mkpielou(params)
      cat("\nProbabilities for", model_name, "model:\n")
      print(probs)
    }
    
    # Print optimized parameters and probabilities for the single model
    cat("\nSingle PLN Model Results:\n")
    cat("Optimized parameters: mean =", fit1$par[1], ", sigma =", fit1$par[2], "\n")
    display_probs(fit1$par, "Single PLN")
    
    # Print AIC value
    aic1 <- 2 * length(fit1$par) + 2 * fit1$value
    cat("\nAIC for Single PLN Model:", aic1, "\n")
  }
  
  # Run the function
  print_model_results()
  
  
  library(ggplot2)
  library(dplyr)
  
  # Remove zeros for analysis
  abundance_filtered_pln <- abundance_raw[abundance_raw > 0]
  
  # Calculate the octave for each species count
  abundance_data_pln <- data.frame(count = abundance_filtered_pln)
  abundance_data_pln$octave <- floor(log2(abundance_data_pln$count))
  
  # Aggregate observed data by octave
  observed_octave_pln <- abundance_data_pln %>%
    group_by(octave) %>%
    summarise(Freq = n(), .groups = 'drop')
  
  # Predicted probabilities from fit1 (should define `fit1` appropriately)
  probs_pln <- mkpielou(fit1$par)
  predicted_counts_pln <- probs_pln * length(abundance_filtered_pln)  # Adjust by total number of observations
 
  # Create a dataframe for predicted counts
  predicted_data_pln <- data.frame(count = 1:length(predicted_counts_pln), predicted = predicted_counts_pln)
  predicted_data_pln$octave <- floor(log2(predicted_data_pln$count))
  
  # Aggregate predicted data by octave
  predicted_octave_pln <- predicted_data_pln %>%
    group_by(octave) %>%
    summarise(Predicted = sum(predicted, na.rm = TRUE), .groups = 'drop')
  
  str(predicted_octave_pln)
  observed_octave_pln
  
  # Initialize the predicted frequencies vector with the correct length
  predicted_octave_freqs_pln <- numeric(length(observed_octave_pln$octave))
  predicted_octave_freqs_pln
  
  # Calculate predicted frequencies for each octave
  for (i in seq_along(predicted_octave_freqs_pln)) {
    bin_range <- 2^(i-1):min(2^i - 1, length(probs_pln))
    predicted_octave_freqs_pln[i] <- sum(probs_pln[bin_range])
  }
  
  # Scale to the total number of observed species
  predicted_octave_freqs_pln <- predicted_octave_freqs_pln * sum(observed_octave_pln)
  predicted_octave_freqs_pln
  
  # Create a dataframe for plotting
  octave_data_pln <- data.frame(
    Octave = seq_along(observed_octave_pln$octave),  # Make sure 'octave' is the correct column name
    Observed = observed_octave_pln$Freq,            # Directly reference the 'Freq' column
    Predicted = predicted_octave_freqs_pln,
    Upper_Bound = 2^(seq_along(observed_octave_pln$octave)-1)
  )
  
  # Generate the plot
  plot_poissonlognormal <- ggplot(octave_data_pln, aes(x = factor(Upper_Bound))) +
    geom_bar(aes(y = Observed), stat = "identity", fill = "darkgray", alpha = 0.7) +
    geom_line(aes(y = Predicted, group = 1), color = "black", size = 1.0) +
    scale_x_discrete(name = "Octave",
                     labels = function(x) sapply(x, function(i) format(i, scientific = FALSE))) +
    labs(title = "Species Abundance Octave Plot",
         y = "Number of Species") +
    theme_minimal() +
    theme(text = element_text(size = 12),
          axis.title = element_text(size = 14),
          plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
          legend.position = "bottom",
          legend.title = element_blank(),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank())
  
  print(plot_poissonlognormal)
  
  # Save the plot
  ggsave("PoissonLogNormal_PredSAD.png", plot_poissonlognormal, width = 10, height = 6, units = "in")
  

  
  
  
  
# Function to calculate CDF from probabilities
generate_fitted_cdf <- function(params) {
  probs <- mkpielou(params)
  total_abundance <- sum(exp(abundance_raw_adjusted_log_series))
  
  # Transform probabilities to match total abundance
  fitted_df <- transform_probabilities(probs, total_abundance)
  
  # Calculate CDF
  fitted_df$cdf <- cumsum(fitted_df$expected_frequency) / sum(fitted_df$expected_frequency)
  return(fitted_df)
}

# Generate CDFs for each model
fitted_cdf_1 <- generate_fitted_cdf(fit1$par)
fitted_cdf_2 <- generate_fitted_cdf(fit2$par)
fitted_cdf_3 <- generate_fitted_cdf(fit3$par)
fitted_cdf_4 <- generate_fitted_cdf(fit4$par)

# Function to plot a single CDF comparison
plot_cdf_comparison <- function(observed, fitted_df, model_name, color) {
  observed_cdf <- ecdf(observed)
  
  # Plot observed CDF
  plot(observed_cdf, main = paste("CDF Comparison:", model_name),
       xlab = "Raw Abundance", ylab = "Cumulative Probability", col = "blue", lwd = 2)
  
  # Overlay fitted CDF
  lines(fitted_df$raw_abundance, fitted_df$cdf, col = color, lwd = 2, type = "l")
  
  # Add legend
  legend("bottomright", legend = c("Observed", "Fitted"), col = c("blue", color), lwd = 2)
}

# Plot CDF comparisons for each model
plot_cdf_comparison(exp(abundance_raw_adjusted_log_series), fitted_cdf_1, "Single PLN Model", "red")
plot_cdf_comparison(exp(abundance_raw_adjusted_log_series), fitted_cdf_2, "Two-Distribution Model", "green")
plot_cdf_comparison(exp(abundance_raw_adjusted_log_series), fitted_cdf_3, "Three-Distribution Model", "purple")
plot_cdf_comparison(exp(abundance_raw_adjusted_log_series), fitted_cdf_4, "Four-Distribution Model", "orange")










# Load required libraries
library(sads)
library(ggplot2)
library(dplyr)

# Fit the logseries model using the `fitsad` function from the sads package
best_fit_logseries <- fitsad(abundance_raw_adjusted_log_series, sad = "ls")

# Extract the fitted parameters
N_logseries <- best_fit_logseries@fullcoef["N"]
alpha_logseries <- best_fit_logseries@fullcoef["alpha"]

# Print the parameters
cat("Fitted parameters for Logseries model:\n")
cat("N =", N_logseries, "\n")
cat("Alpha =", alpha_logseries, "\n")

# Calculate the octave for each species count
abundance_data_ls <- data.frame(count = abundance_raw_adjusted_log_series)
abundance_data_ls$octave <- floor(log2(abundance_data_ls$count))

# Aggregate observed data by octave
observed_octave_ls <- abundance_data_ls %>%
  group_by(octave) %>%
  summarise(Freq = n(), .groups = 'drop')

# Generate predicted probabilities using the log-series PDF
max_abund_ls <- max(abundance_raw_adjusted_log_series)
abund_seq_ls <- seq(1, max_abund_ls, by = 1)

# Compute probabilities using the logseries model
predicted_freqs_ls <- dls(abund_seq_ls, N = N_logseries, alpha = alpha_logseries)
predicted_counts_ls <- predicted_freqs_ls * length(abundance_raw_adjusted_log_series)

# Create a dataframe for predicted counts
predicted_data_ls <- data.frame(count = abund_seq_ls, predicted = predicted_counts_ls)
predicted_data_ls$octave <- floor(log2(predicted_data_ls$count))

# Aggregate predicted data by octave
predicted_octave_ls <- predicted_data_ls %>%
  group_by(octave) %>%
  summarise(Predicted = sum(predicted, na.rm = TRUE), .groups = 'drop')


# Merge observed and predicted data for plotting
octave_data_ls <- left_join(observed_octave_ls, predicted_octave_ls, by = "octave")
octave_data_ls$Upper_Bound <- 2^(octave_data_ls$octave)

# Define the plot
plot_logseries <- ggplot(octave_data_ls, aes(x = factor(Upper_Bound))) +
  geom_bar(aes(y = Freq), stat = "identity", fill = "darkgray", alpha = 0.7) +  # Observed data
  geom_line(aes(y = Predicted, group = 1), color = "black", size = 1.0) +  # Fitted model
  scale_x_discrete(name = "Octave",
                   labels = function(x) sapply(x, function(i) format(i, scientific = FALSE))) +
  labs(title = "Species Abundance Octave Plot - Logseries Model",
       y = "Number of Species") +
  theme_minimal() +
  theme(text = element_text(size = 12),
        axis.title = element_text(size = 14),
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        legend.position = "bottom",
        legend.title = element_blank(),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank())

# Print the plot
print(plot_logseries)

# Save the plot
ggsave("Logseries_PredSAD.png", plot_logseries, width = 10, height = 6, units = "in")
