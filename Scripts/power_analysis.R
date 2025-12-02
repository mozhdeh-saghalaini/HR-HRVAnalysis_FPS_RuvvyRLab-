##### Last Update: 11/30/2025 ####

# Authored by Mozhdeh Saghalaini: m.saghalaini@gmail.com
# ------------------------------------------------------------------------------

# Statistical Power Analysis

## Description:
# Uses the first imputed dataset from multiple imputation

## Input:
# Final_Analysis_Dataset.csv (from missing_data_analysis.R)

## Output:
# Power_Analysis_Report.txt (power analysis results)

# ------------------------------------------------------------------------------
##### Loading Packages #####

library(tidyverse) # data manipulation
library(pwr)       # power analysis
library(broom)     # Tidy model output

#### Loading Data and Sample Size checking ####
analysis_data <- read_csv("D:/Research/FPS (Ruvvy RLab)/Codes/Output/Final_Analysis_Dataset.csv")
final_n <- n_distinct(analysis_data$id)

cat("Dataset:", nrow(analysis_data), "records\n")
cat("Unique participants:", final_n, "\n")

# Task distribution
task_dist <- analysis_data %>%
  count(task_type)
cat("\nTask distribution:\n")
print(task_dist)

#### Power Analysis for correlation ####

effect_sizes <- c(0.1, 0.3, 0.5)  # Small, medium and large effect sizes

correlation_power <- map_dfr(effect_sizes, function(r) {
  power_result <- pwr.r.test(
    n = final_n,
    r = r,
    sig.level = 0.05,
    alternative = "two.sided"
  )
  
  data.frame(
    effect_size = r,
    power = power_result$power,
    minimum_detectable_r = power_result$r,
    interpretation = case_when(
      r == 0.1 ~ "Small",
      r == 0.3 ~ "Medium", 
      r == 0.5 ~ "Large"
    )
  )
})

print(correlation_power)

# What effect size can we detect with 80% power?
detectable_effect <- pwr.r.test(
  n = final_n,
  power = 0.80,
  sig.level = 0.05
)

cat("\nMinimum detectable correlation with 80% power: r =", 
    round(detectable_effect$r, 3), "\n")

#### Power analysis for regression ####
cat("\n Power analysis for regression \n")

predictor_scenarios <- list(
  "Simple" = 2,    
  "Moderate" = 4, 
  "Complex" = 6   
)

regression_power <- map_dfr(names(predictor_scenarios), function(scenario) {
  u <- predictor_scenarios[[scenario]]
  v <- final_n - u - 1  # df
  
  effect_scenarios <- c(0.02, 0.15, 0.35) 
  
  map_dfr(effect_scenarios, function(f2) {
    power_result <- pwr.f2.test(
      u = u,
      v = v,
      f2 = f2,
      sig.level = 0.05
    )
    
    data.frame(
      scenario = scenario,
      predictors = u,
      effect_size_f2 = f2,
      power = power_result$power,
      interpretation = case_when(
        f2 == 0.02 ~ "Small",
        f2 == 0.15 ~ "Medium",
        f2 == 0.35 ~ "Large"
      )
    )
  })
})

print(regression_power)

#### Power analysis for t-tests (group comparisons) ####
cat("\n Power analysis for t-tests (group comparisons \n")

# For between-group comparisons
t_test_power <- map_dfr(c(0.2, 0.5, 0.8), function(d) {
  power_result <- pwr.t.test(
    n = final_n / 2,  # Assuming equal group sizes
    d = d,
    sig.level = 0.05,
    type = "two.sample"
  )
  
  data.frame(
    cohens_d = d,
    power = power_result$power,
    interpretation = case_when(
      d == 0.2 ~ "Small",
      d == 0.5 ~ "Medium",
      d == 0.8 ~ "Large"
    )
  )
})

print(t_test_power)


#### Sample size requirements ####

# Required sample sizes for 80% power
required_samples <- data.frame(
  analysis_type = c("Correlation (r=0.3)", "Regression (4 predictors, f2=0.15)", "T-test (d=0.5)"),
  required_n = c(
    pwr.r.test(r = 0.3, power = 0.80)$n,
    NA,  # Regression calculation is different
    pwr.t.test(d = 0.5, power = 0.80)$n * 2  # *2 for total sample
  )
)

# For regression
reg_required <- pwr.f2.test(u = 4, f2 = 0.15, power = 0.80)
required_samples$required_n[2] <- reg_required$v + 4 + 1

cat("Sample size requirements for 80% power:\n")
print(required_samples)


#### Power summary ####
# Key correlation power
key_cor_power <- correlation_power %>% 
  filter(effect_size == 0.3) %>% 
  pull(power)

cat("Power to detect medium correlations (r = 0.3):", round(key_cor_power * 100, 1), "%\n")

# Key regression power  
key_reg_power <- regression_power %>% 
  filter(scenario == "Moderate" & effect_size_f2 == 0.15) %>% 
  pull(power)

cat("Power for moderate regression (4 predictors):", round(key_reg_power * 100, 1), "%\n")

cat("Recommended sample for 80% power:", round(required_samples$required_n[1]), "participants\n")

#### Saving power analysis report ####

sink("D:/Research/FPS (Ruvvy RLab)/Codes/Output/Power_Analysis_Report.txt")

cat("Analysis Date:", format(Sys.Date(), "%Y-%m-%d"), "\n")
cat("Final Sample Size:", final_n, "participants\n\n")

cat("CORRELATION POWER ANALYSIS\n")
print(correlation_power)
cat("\n")

cat("REGRESSION POWER ANALYSIS\n")
print(regression_power)
cat("\n")

cat("T-TEST POWER ANALYSIS\n")
print(t_test_power)
cat("\n")

cat("SAMPLE SIZE REQUIREMENTS FOR 80% POWER\n")
print(required_samples)
cat("\n")


sink()
