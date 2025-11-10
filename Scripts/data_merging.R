# Last Update: 11/10/2020

##### Loading Packages #####

library(tidyverse)
library(janitor)
library(dylyr)
library(readxl)    


#### Extra Data cleaning and quality xhecks ####
hrv_final <- read_csv("D:/Research/FPS (Ruvvy RLab)/Codes/Output/Cleaned_Objective_HR_data.csv")

if(nrow(hrv_final) > 0) {
  
  cat("\n Dtaa cleaning\n")
  
  # Check for completely missing files or variables
  missing_ids <- sum(is.na(hrv_final$id)) 
  cat("Files with missing IDs:", missing_ids, "\n")
  
  # Select only the metrics we need
  hrv_data_clean <- hrv_final %>%
    select(
      
      # Metadata 
      id, sex, task_type, task_version, collection_date,
      source_file,
      
      #### Selecting metrics (I should double check the metrics' labels) ####   
      mean_hr, mean_rmssd, mean_sdnn) %>%
    
    # Removing the rows that all the metrics are missing
    filter(!if_all(c(mean_hr, mean_rmssd), is.na))
  
  cat("Original files:", nrow(hrv_final), "\n")
  cat("After cleaning:", nrow(hrv_data_clean), "\n")
  cat("Files removed:", nrow(hrv_final) - nrow(hrv_data_clean), "\n")
  
  #### Save the cleaned dataset ####
  write_csv(hrv_data_clean, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Selected_Objective_HR_data.csv")
  
  cat("File: Selected_Objective_HR_data.csv\n")
  
} else {
  cat("\n No data cleaned !\n")
}


#### Subjective and objective data merging ####
# Load the subjective data !!!! I assumed the name of the file is "subjective_data.xlsx" but hsould check and refine this 
subjective_data <- read_excel("01_Raw_Data/subjective_data.xlsx") %>% clean_names()

# Merging the datasets by participant id
full_data <- hrv_data %>% left_join(subjective_data, by = "id")
glimpse(full_data)

# Creating key variables for analysis (sample)
analysis_data <- full_data %>%
  filter(task == "AQ" & round == "1") %>% # Focus on AQ1
  mutate(
    # Creating a hyperarousal score (sum of relevant UCLA items)
    hyperarousal_score = ucla_irritable + ucla_concentrate + ucla_sleep + ucla_danger, # check this with the available data
    # Ensure sex is a factor
    sex = as.factor(sex)
  )

# Save the final analysis dataset
write_csv(analysis_data, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/final_analysis_dataset.csv")


