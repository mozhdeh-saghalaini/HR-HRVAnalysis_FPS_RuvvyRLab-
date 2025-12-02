##### Last Update: 11/30/2025 ####

# Authored by Mozhdeh Saghalaini: m.saghalaini@gmail.com
# ------------------------------------------------------------------------------

# ECG Data Merging With Subjective Data

## Description:
# This merges cleaned ECG data with subjective data

## Inputs:
# 1. Cleaned_ECG_Data.csv (from data_wrangling.R)
# 2. subjective_data.xlsx 

## Output:
# Merged_Data_Before_Imputation.csv

# ------------------------------------------------------------------------------

## warning: the current output will allocate separate rows for EXT and ACQ of the same partcipant
## - consult this with Dr. Grasser, whether it's better to combine EXT/ACQ files of the same 
## participant in one row or this version is fine


##### Loading Packages #####

library(tidyverse) # data manipulation and visualization
library(readxl)    # Reading .xlsx files   
library(lubridate) # Data handling
library(janitor)   # data cleaning process

#### Loading ECG data ####
hrv_clean <- read_csv("D:/Research/FPS (Ruvvy RLab)/Codes/Output/Cleaned_ECG_Data.csv")

cat("Cleaned_ECG_Data:", nrow(hrv_clean), "records\n")
print(table(hrv_clean$task_type))
cat("Unique participants in Objective data:", n_distinct(hrv_clean$id), "\n")


#### Loading subjective Data ####

# I assumed the name of the file is "subjective_data.xlsx" and it's adjusted based on the Empty Subjective Vars and ... files 
subjective_data <- read_excel("D:/Research/FPS (Ruvvy RLab)/Codes/RawData/subjective_data.xlsx") %>%
  clean_names() %>%
  
  select(
    id = sid,                        # For combining subjective and objevtive data of each participant
    
    # age_subjective = age_at_visit,   # For verification
    sex_subjective = sex,            # For verification
    
    trauma_exposure = htq,           # Harvard Trauma Questionnaire
    ptsd_total = ucla,               # UCLA PTSD total
    anxiety_total = scared,          # SCARED total
    
    # UCLA subscales
    ucla_intrusion = ucla_clusterb,
    ucla_avoidance = ucla_clusterc,
    ucla_cog_alternations = ucla_clusterd,
    ucla_arousal_react = ucla_clustere,
    ucla_dissociative = ucla_dissociative,
    
    # Anxiety subscales
    scared_panic_somatic,                 
    scared_gad,                      # Generalized anxiety  
    scared_social_anxiety,   
    scared_separation_anxiety, 
    scared_school_avoidance,
    
    # Trauma-related variables
    death_threats, 
    victimization, 
    accident_injury, 
    cumulative_lec,
    
    # Developmental 
    pds_total                        # There is no PDS variable in the Empty file Dr. Grasser sent!!!!! check this
    
  )
cat("subjective_data:", nrow(subjective_data), "records\n")
cat("Unique participants in Subjective data:", n_distinct(subjective_data$id), "\n")


#### Data Merging ####

analysis_data <- hrv_clean %>%
  left_join(subjective_data, by = "id") %>%
  
  # verification to ensure datasets are correclty merged
  mutate(
    sex = as.factor(sex),
    task_type = as.factor(task_type),
    sex_match = ifelse(sex == sex_subjective, TRUE, FALSE)
  )

# Checking for merging issues
unmatched_ecg <- hrv_clean %>% 
  filter(!id %in% subjective_data$id) %>%
  nrow()

unmatched_subjective <- subjective_data %>% 
  filter(!id %in% hrv_clean$id) %>% 
  nrow()

cat("ECG records without matching subjective data:", unmatched_ecg, "\n")
cat("Subjective records without matching ECG data:", unmatched_subjective, "\n")

cat("Sex mismatches:", sum(!analysis_data$sex_match, na.rm = TRUE), "\n")

if(sum(!analysis_data$sex_match, na.rm = TRUE) > 0) {
  mismatches <- analysis_data %>% 
    filter(!sex_match) %>%
    select(id, sex, sex_subjective, task_type)
  print(mismatches)
}

#### Merged Data Summary ####

cat("Merged dataset:", nrow(analysis_data), "records\n")
cat("Unique participants in merged data:", n_distinct(analysis_data$id), "\n")
cat("Variables in merged dataset:", ncol(analysis_data), "\n")
print(table(analysis_data$task_type))

#### Saving Merged Data ####

write_csv(analysis_data, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Merged_Data_Before_Imputation.csv")
cat("\nMerged data saved: Merged_Data_Before_Imputation.csv\n")



