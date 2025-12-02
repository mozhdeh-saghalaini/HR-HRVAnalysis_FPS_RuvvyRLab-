##### Last Update: 12/2/2025 ####

# Authored by Mozhdeh Saghalaini: m.saghalaini@gmail.com
# ------------------------------------------------------------------------------

# Little's Test for Missing Data Analysis and Multiple Imputation

## Description:
# Based on Dr. Grasser's missing data analysis approach

## Input:
# Merged_Data_Before_Imputation.csv  (from data_merging.R)

## Outputs:
# 1. Final_Analysis_Dataset.csv (imputed dataset)
# 2. Data_Codebook.csv (variables documentation)
# 3. Missingness_Report.csv (missing data patterns)

# ------------------------------------------------------------------------------
##### Loading Packages #####

library(tidyverse) # data manipulation and visualization
library(naniar)    # For missing data visualization
library(mice)      # For multiple imputation
library(finalfit)  # For missing data tests
library(janitor)   # data cleaning process

#### Loading Merged Dataset ####
analysis_data <- read_csv("D:/Research/FPS (Ruvvy RLab)/Codes/Output/Merged_Data_Before_Imputation.csv")

#### Missing data analysis ####

# # Selecting key variables for imputation 
# imputation_vars <- analysis_data %>% 
#   select(id, sex, task_type, age_subjective, pds_total,
#          overall_mean_hr_mean, overall_rmssd_mean, overall_sdnn_mean,
#          trauma_exposure, ptsd_total, anxiety_total, ucla_arousal_react)

# Keeping all the variables for now
imputation_vars <- analysis_data 

#### Checking missingness patterns but plotting ####

# checking to see if missingness relates to other variables
explanatory <- c("sex", "age_subjective") 
dependent <- c("overall_rmssd_mean", "ptsd_total", "anxiety_total") 

# Plot
tryCatch({
  imputation_vars %>% 
    missing_pairs(dependent, explanatory)
}, error = function(e) {
  cat("Note: Could not create missing pairs plot -", e$message, "\n")
})

missing_summary <- imputation_vars %>%
  summarise(across(everything(), ~sum(is.na(.))/n()*100)) %>%
  pivot_longer(everything(), names_to = "variable", values_to = "pct_missing") %>%
  arrange(desc(pct_missing))

print(missing_summary)
write_csv(missing_summary, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Missingness_Report.csv")

#### MCAR Test ####

mcar_data <- imputation_vars %>% select(-id) # removing IDs
mcar_result <- mcar_test(mcar_data)
cat("Little's MCAR test p-value:", mcar_result$p.value, "\n")

if(mcar_result$p.value < 0.05) {
  cat("May not be missing completely at random (p =", round(mcar_result$p.value, 3), ")\n")
} else {
  cat("Seems to be missing completely at random (p =", round(mcar_result$p.value, 3), ")\n")
}

#### Multiple Imputation ####

impute_ready <- imputation_vars %>%
  select(-id) %>%
  mutate(
    sex = as.factor(sex),
    task_type = as.factor(task_type)
  )

imputed_data <- mice(impute_ready, 
                     m = 5,           # Create 5 imputed datasets, I should adjust this based on our data missingness (5 is for low missginness like around 10 percent)
                     maxit = 50,      # iterations
                     method = 'pmm',  # Predictive mean matching
                     seed = 123)    

# Save the complete imputation object 
saveRDS(imputed_data, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Imputed_Data_Object.rds")

# Plot to see if imputed values match real data distribution
densityplot(imputed_data)

# Using the first imputed dataset for your main analysis
analysis_data_final <- complete(imputed_data, 1) %>% 
  bind_cols(imputation_vars %>% select(id)) %>% 
  left_join(analysis_data %>% 
              select(id, collection_date, source_file, pct_usable_segments, 
                     total_usable_segments, n_segments, task_version), 
            by = "id")


cat("Final dataset has", nrow(analysis_data_final), "rows.\n")

# # Previous version: simply deleting variables with high % of missingness
# # Applying missing data threshold
# variables_to_keep <- missing_summary %>% filter(pct_missing < 0.20) %>% pull(variable)
# analysis_data_final <- analysis_data %>% select(all_of(variables_to_keep))
# cat("Remained variables (< 20% missing data):", length(variables_to_keep), "/", ncol(analysis_data), "\n")


#### Final Data Export ####

write_csv(analysis_data_final, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Final_Analysis_Dataset.csv")

# Create codebook
codebook <- data.frame(
  Variable = names(analysis_data),
  Description = c(
    # Metadata
    "Participant ID",
    "Sex (M/F)",
    "Task type (AQ or EXT)",
    "Task version (1 or 2)",
    "Collection date",
    "Source file name",
    "Data sheet name",
    
    # Segment-level HRV metrics (example: seg1_mean_hr, seg2_mean_hr, …)
    "Segment-level mean HR", 
    "Segment-level RMSSD",
    "Segment-level SDNN",
    
    # Overall summary metrics
    "Overall mean HR across segments",
    "SD of HR across segments",
    "Overall mean RMSSD across segments",
    "SD of RMSSD across segments",
    "Overall mean SDNN across segments",
    "SD of SDNN across segments",
    
    # Phasic metrics (baseline, early, late, reactivity indices)
    "Baseline HR",
    "Early task HR",
    "Late task HR",
    "Change in HR (late - early)",
    "Early HR reactivity (early - baseline)",
    "Late HR reactivity (late - baseline)",
    
    "Baseline RMSSD",
    "Early task RMSSD",
    "Late task RMSSD",
    "Change in RMSSD (late - early)",
    "Early RMSSD reactivity (early - baseline)",
    "Late RMSSD reactivity (late - baseline)",
    
    "Baseline SDNN",
    "Early task SDNN",
    "Late task SDNN",
    "Change in SDNN (late - early)",
    "Early SDNN reactivity (early - baseline)",
    "Late SDNN reactivity (late - baseline)",
    
    # Quality indicators
    "% usable segments (HR)",
    "Total usable segments (HR)",
    
    # Subjective questionnaire variables
    "Sex from subjective data",
    "Trauma exposure (HTQ)",
    "PTSD total (UCLA)",
    "Anxiety total (SCARED)",
    "UCLA intrusion subscale",
    "UCLA avoidance subscale",
    "UCLA cognitive alterations subscale",
    "UCLA arousal/reactivity subscale",
    "UCLA dissociative subscale",
    "SCARED panic/somatic subscale",
    "SCARED generalized anxiety subscale",
    "SCARED social anxiety subscale",
    "SCARED separation anxiety subscale",
    "SCARED school avoidance subscale",
    "Trauma: death threats",
    "Trauma: victimization",
    "Trauma: accident/injury",
    "Cumulative LEC score",
    "Developmental stage (PDS total)",
    
    # Verification
    "Sex match flag (objective vs subjective)"
  ),
  Type = sapply(analysis_data, class),
  stringsAsFactors = FALSE
)

write_csv(codebook, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Data_Codebook.csv")
cat("Codebook saved: Data_Codebook.csv\n")


cat("Final analysis dataset saved: Final_Analysis_Dataset.csv\n")

cat("\nDataset dimensions:", nrow(analysis_data_final), "rows ×", ncol(analysis_data_final), "columns\n")
cat("Unique participants:", n_distinct(analysis_data_final$id), "\n")


