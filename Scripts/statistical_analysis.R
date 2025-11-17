# Last Update: 11/10/2025

##### Loading Packages #####
library(tidyverse)       
library(readxl)          
library(broom)           
library(performance)     
library(psych)           
library(corrplot)        
library(emmeans)         
library(effects)        
library(patchwork)       
# APA packages
library(apaTables)      
library(sjPlot)          
library(sjlabelled)      
library(sjmisc)          
library(apa)             
library(knitr)          
library(kableExtra)      

#### Loading and preparing data ####
# Load cleaned HRV data
hrv_data <- read_csv("D:/Research/FPS (Ruvvy RLab)/Codes/Output/final_analysis_dataset.csv")

# Load clinical data (PTSD symptoms, anxiety, demographics)
clinical_data <- read_csv("D:/Research/FPS (Ruvvy RLab)/Codes/RawData/Clinical_Data.csv") %>%
  clean_names()

# Merge datasets and prepare for analysis
analysis_data <- hrv_data %>%
  left_join(clinical_data, by = "id") %>%
  mutate(
    sex = as.factor(sex),
    task_type = as.factor(task_type),
    # Create hyperarousal score (adjust based on your UCLA items)
    hyperarousal = ucla_irritable + ucla_concentrate + ucla_sleep + ucla_danger,
    # Ensure no missing values in key variables
    across(c(overall_mean_hr_mean, overall_rmssd_mean, ucla_total), 
           ~ifelse(is.na(.), median(., na.rm = TRUE), .))
  ) %>%
  # Add variable labels for APA tables
  var_labels(
    overall_mean_hr_mean = "Mean Heart Rate (bpm)",
    overall_rmssd_mean = "RMSSD (ms)",
    overall_sdnn_mean = "SDNN (ms)",
    ucla_total = "PTSD Symptom Severity",
    hyperarousal = "Hyperarousal Symptoms",
    age = "Age (years)",
    sex = "Sex",
    pds_total = "Pubertal Development Score"
  )

cat("=== DATA LOADED ===\n")
cat("Participants:", n_distinct(analysis_data$id), "\n")
cat("Total observations:", nrow(analysis_data), "\n")

#### APA Descriptive Statistics Table ####
cat("\n=== APA DESCRIPTIVE STATISTICS ===\n")

# Create APA-style descriptive table
desc_vars <- analysis_data %>%
  select(overall_mean_hr_mean, overall_rmssd_mean, overall_sdnn_mean,
         ucla_total, hyperarousal, age, pds_total)

apa_desc <- apa.descriptives(desc_vars, show.conf.interval = TRUE)

# Print to console
print(apa_desc)

# Save as Word document
apaTables::apa.desc.table(desc_vars, 
                          filename = "03_Output/APA_Tables/descriptive_statistics.doc",
                          table.number = 1)

# Create descriptive table by sex
desc_by_sex <- analysis_data %>%
  group_by(sex) %>%
  select(overall_mean_hr_mean, overall_rmssd_mean, ucla_total, age) %>%
  describe() %>%
  as.data.frame() %>%
  rownames_to_column("Variable") %>%
  mutate(Group = rep(c("Female", "Male"), each = 4))

# Format for APA
desc_by_sex_apa <- desc_by_sex %>%
  mutate(across(where(is.numeric), ~round(., 2))) %>%
  select(Variable, Group, mean, sd, min, max)

print(kable(desc_by_sex_apa, format = "simple", caption = "Descriptive Statistics by Sex"))

#### Hypothesis 1a: Sex Differences in HR and HRV - APA Format ####
cat("\n=== HYPOTHESIS 1a: SEX DIFFERENCES IN HR/HRV (APA FORMAT) ===\n")

# T-tests with APA formatting
t_rmssd <- t.test(overall_rmssd_mean ~ sex, data = analysis_data)
t_hr <- t.test(overall_mean_hr_mean ~ sex, data = analysis_data)

# APA format for t-tests
apa_t_rmssd <- apa::t_test(t_rmssd, print = FALSE)
apa_t_hr <- apa::t_test(t_hr, print = FALSE)

cat("RMSSD Sex Difference:\n")
cat(apa_t_rmssd, "\n\n")

cat("Heart Rate Sex Difference:\n")
cat(apa_t_hr, "\n")

# Effect sizes with CI
cohens_d_rmssd <- effectsize::cohens_d(overall_rmssd_mean ~ sex, data = analysis_data)
cohens_d_hr <- effectsize::cohens_d(overall_mean_hr_mean ~ sex, data = analysis_data)

cat("Effect Sizes (Cohen's d) with 95% CI:\n")
cat("RMSSD: d =", round(cohens_d_rmssd$Cohens_d, 2), 
    "[", round(cohens_d_rmssd$CI_low, 2), ",", round(cohens_d_rmssd$CI_high, 2), "]\n")
cat("Heart Rate: d =", round(cohens_d_hr$Cohens_d, 2),
    "[", round(cohens_d_hr$CI_low, 2), ",", round(cohens_d_hr$CI_high, 2), "]\n")

#### Hypothesis 1b: PTSD Symptoms and Autonomic Dysregulation - APA Format ####
cat("\n=== HYPOTHESIS 1b: PTSD SYMPTOMS AND AUTONOMIC DYSREGULATION (APA FORMAT) ===\n")

# Correlation analysis with APA formatting
cor_hr_ptsd <- cor.test(analysis_data$overall_mean_hr_mean, analysis_data$ucla_total)
cor_rmssd_ptsd <- cor.test(analysis_data$overall_rmssd_mean, analysis_data$ucla_total)

# APA format for correlations
apa_cor_hr <- apa::cor_apa(cor_hr_ptsd, print = FALSE)
apa_cor_rmssd <- apa::cor_apa(cor_rmssd_ptsd, print = FALSE)

cat("Correlation - HR vs PTSD Symptoms:\n")
cat(apa_cor_hr, "\n\n")

cat("Correlation - RMSSD vs PTSD Symptoms:\n")
cat(apa_cor_rmssd, "\n")

# Linear models for regression tables
model_1b_hr <- lm(ucla_total ~ overall_mean_hr_mean + sex + age, data = analysis_data)
model_1b_rmssd <- lm(ucla_total ~ overall_rmssd_mean + sex + age, data = analysis_data)

# Create APA regression tables
cat("\n=== REGRESSION TABLES ===\n")

# Using sjPlot for publication-ready tables
tab_model(model_1b_hr, 
          title = "Table 2. Regression of PTSD Symptoms on Heart Rate",
          file = "03_Output/APA_Tables/regression_hr_predicting_ptsd.doc")

tab_model(model_1b_rmssd,
          title = "Table 3. Regression of PTSD Symptoms on RMSSD", 
          file = "03_Output/APA_Tables/regression_rmssd_predicting_ptsd.doc")

# Print simplified version to console
cat("Heart Rate Predicting PTSD Symptoms:\n")
sjPlot::tab_model(model_1b_hr, show.se = TRUE, show.ci = FALSE)

cat("\nRMSSD Predicting PTSD Symptoms:\n") 
sjPlot::tab_model(model_1b_rmssd, show.se = TRUE, show.ci = FALSE)

#### Objective 2a: Screen Covariates - APA Format ####
cat("\n=== OBJECTIVE 2a: SCREEN COVARIATES (APA FORMAT) ===\n")

model_2a <- lm(ucla_total ~ age + sex + pds_total, data = analysis_data)

# APA table for covariate screening
tab_model(model_2a,
          title = "Table 4. Covariate Screening for PTSD Symptoms",
          file = "03_Output/APA_Tables/covariate_screening.doc")

cat("Covariate Screening Model:\n")
sjPlot::tab_model(model_2a, show.se = TRUE, show.ci = FALSE)

#### Objective 2b: HRV × Hyperarousal Interaction - APA Format ####
cat("\n=== OBJECTIVE 2b: HRV × HYPERAROUSAL INTERACTION (APA FORMAT) ===\n")

# Center variables
analysis_data <- analysis_data %>%
  mutate(
    rmssd_c = scale(overall_rmssd_mean, scale = FALSE),
    hyperarousal_c = scale(hyperarousal, scale = FALSE)
  )

# Moderation model
model_2b <- lm(ucla_total ~ rmssd_c + hyperarousal_c + rmssd_c:hyperarousal_c + 
                 sex + age, data = analysis_data)

# APA table for moderation analysis
tab_model(model_2b,
          title = "Table 5. Moderation Analysis: HRV × Hyperarousal on PTSD Symptoms",
          file = "03_Output/APA_Tables/moderation_analysis.doc")

cat("Moderation Analysis - HRV × Hyperarousal:\n")
sjPlot::tab_model(model_2b, show.se = TRUE, show.ci = FALSE)

# Simple slopes analysis if interaction significant
interaction_p <- tidy(model_2b) %>% 
  filter(term == "rmssd_c:hyperarousal_c") %>% 
  pull(p.value)

if(interaction_p < 0.05) {
  cat("\n=== SIMPLE SLOPES ANALYSIS (APA FORMAT) ===\n")
  
  simple_slopes <- emtrends(model_2b, ~hyperarousal_c, var = "rmssd_c",
                            at = list(hyperarousal_c = c(-1, 0, 1))) # ±1 SD
  
  # APA format for simple slopes
  apa_simple_slopes <- apa::emmeans(simple_slopes, print = FALSE)
  cat("Simple Slopes Analysis:\n")
  print(apa_simple_slopes)
}

#### APA Correlation Matrix ####
cat("\n=== APA CORRELATION MATRIX ===\n")

# Create correlation matrix with p-values
cor_vars <- analysis_data %>%
  select(overall_mean_hr_mean, overall_rmssd_mean, overall_sdnn_mean,
         ucla_total, hyperarousal, age, pds_total)

# APA correlation table
apa.cor.table(cor_vars, 
              filename = "03_Output/APA_Tables/correlation_matrix.doc",
              table.number = 6)

# Also create using apaTables for alternative format
apaTables::apa.cor.table(cor_vars,
                         filename = "03_Output/APA_Tables/correlation_matrix_alt.doc",
                         table.number = 7)

#### Create Comprehensive Results Summary ####
cat("\n=== COMPREHENSIVE RESULTS SUMMARY ===\n")

# Create a summary dataframe of all key results
results_summary <- bind_rows(
  # Hypothesis 1a - Sex differences
  data.frame(
    Hypothesis = "1a - Sex Difference RMSSD",
    Test = "t-test",
    Statistic = paste0("t(", t_rmssd$parameter, ") = ", round(t_rmssd$statistic, 2)),
    p_value = round(t_rmssd$p.value, 3),
    Effect_Size = paste0("d = ", round(cohens_d_rmssd$Cohens_d, 2)),
    CI_95 = paste0("[", round(cohens_d_rmssd$CI_low, 2), ", ", round(cohens_d_rmssd$CI_high, 2), "]")
  ),
  data.frame(
    Hypothesis = "1a - Sex Difference HR",
    Test = "t-test", 
    Statistic = paste0("t(", t_hr$parameter, ") = ", round(t_hr$statistic, 2)),
    p_value = round(t_hr$p.value, 3),
    Effect_Size = paste0("d = ", round(cohens_d_hr$Cohens_d, 2)),
    CI_95 = paste0("[", round(cohens_d_hr$CI_low, 2), ", ", round(cohens_d_hr$CI_high, 2), "]")
  ),
  # Hypothesis 1b - Correlations
  data.frame(
    Hypothesis = "1b - HR-PTSD Correlation",
    Test = "Pearson correlation",
    Statistic = paste0("r(", cor_hr_ptsd$parameter, ") = ", round(cor_hr_ptsd$estimate, 2)),
    p_value = round(cor_hr_ptsd$p.value, 3),
    Effect_Size = "",
    CI_95 = paste0("[", round(cor_hr_ptsd$conf.int[1], 2), ", ", round(cor_hr_ptsd$conf.int[2], 2), "]")
  ),
  data.frame(
    Hypothesis = "1b - RMSSD-PTSD Correlation", 
    Test = "Pearson correlation",
    Statistic = paste0("r(", cor_rmssd_ptsd$parameter, ") = ", round(cor_rmssd_ptsd$estimate, 2)),
    p_value = round(cor_rmssd_ptsd$p.value, 3),
    Effect_Size = "",
    CI_95 = paste0("[", round(cor_rmssd_ptsd$conf.int[1], 2), ", ", round(cor_rmssd_ptsd$conf.int[2], 2), "]")
  )
)

# Print results summary
print(kable(results_summary, format = "simple", 
            caption = "Summary of Hypothesis Testing Results"))

# Save results summary
write_csv(results_summary, "03_Output/APA_Tables/hypothesis_testing_summary.csv")

#### Save All Models for Future Reference ####
cat("\n=== SAVING ANALYSIS OBJECTS ===\n")

analysis_objects <- list(
  data = analysis_data,
  models = list(
    model_1b_hr = model_1b_hr,
    model_1b_rmssd = model_1b_rmssd, 
    model_2a = model_2a,
    model_2b = model_2b
  ),
  tests = list(
    t_rmssd = t_rmssd,
    t_hr = t_hr,
    cor_hr_ptsd = cor_hr_ptsd,
    cor_rmssd_ptsd = cor_rmssd_ptsd
  )
)

saveRDS(analysis_objects, "03_Output/Analysis/complete_analysis_objects.rds")

cat("\n=== APA ANALYSIS COMPLETE ===\n")
cat("All analyses completed with APA formatting\n")
cat("Word documents with publication-ready tables saved in: 03_Output/APA_Tables/\n")
cat("Key files generated:\n")
cat("- descriptive_statistics.doc (Table 1)\n")
cat("- regression_hr_predicting_ptsd.doc (Table 2)\n") 
cat("- regression_rmssd_predicting_ptsd.doc (Table 3)\n")
cat("- covariate_screening.doc (Table 4)\n")
cat("- moderation_analysis.doc (Table 5)\n")
cat("- correlation_matrix.doc (Table 6)\n")