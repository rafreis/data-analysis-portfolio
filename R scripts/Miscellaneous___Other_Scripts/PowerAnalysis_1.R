library(pwr)

## A PRIORI

# Assumptions
alpha <- 0.05
desired_power <- 0.80
num_predictors <- 12  # From the data you provided
expected_r2 <- 0.26  # Expected proportion of variance explained by the predictors

# Calculate f2 effect size for regression: f2 = R^2 / (1 - R^2)
f2 <- expected_r2 / (1 - expected_r2)

# Calculate required sample size
sample_size_result <- pwr.f2.test(u = num_predictors, v = NULL, f2 = f2, 
                                  sig.level = alpha, power = desired_power)

# Calculate the total sample size: n = u + v + 1
required_sample_size <- num_predictors + sample_size_result$v + 1

# Print results
print(sample_size_result)
cat("Total required sample size:", round(required_sample_size), "\n")


## POST HOC

# Given data from ANOVA table (regression)
f_statistic <- 39.245  # F-value from the ANOVA table
df1 <- 12  # numerator degrees of freedom (number of predictors)
df2 <- 606  # denominator degrees of freedom

# Post hoc power analysis for F-test in regression
power_result <- pwr.f2.test(u = df1, v = df2, f2 = f_statistic / df2, sig.level = 0.05)

print(power_result)