##### Last Update: 11/16/2025 ####
# Authored by Mozhdeh Saghalaini: m.saghalaini@gmail.com

# Some of the notes are for double checking and some are for me to remmeber stuff
# This code integrats data wrangling, cleaning, and merging
# It processes the raw MindWare files and merges them with subjective questionnaire data 
# and extracts HRV metrics from the "HRV Stats" sheet, calculates overall averages, and performs quality checks based on the word that Dr. Grasser sent me. 
# methodology (MAD outlier detection, missing data handling), and creates a final analysis-ready dataset. 
# The output is a clean dataset with both subjective and objective variables.

# Outputs:
# 1. Cleaned_ECG_Data.csv - processed HR/HRV metrics with quality indicators
# 2. Final_Analysis_Dataset.csv - merged ECG + questionnare data 
# 3. Data_Codebook.csv - Documentation of all variables and their meanings

# This code handles multiple participants and segments,
# -includes robust data quality checks (outliers, missing data, physiological plausibility),

##### Loading Packages #####

library(tidyverse) # data manipulation and visualization
library(readxl)    # Reading .xlsx files   
library(purrr)     # functional programming
library(stringr)   # String manipulation
library(lubridate) # Data handling
library(janitor)   # data cleaning process
library(robust)    # Robust outlier detection


#### Function: Reading Mindware Files#### 

read_mindware_files <- function(filename) {
  tryCatch({
    
    #### Sheet Selection ####
    segment_sheet_name <- "HRV Stats" # This might need some adjustments
    
    # Read the specified sheet
    d <- read_excel(filename, sheet = segment_sheet_name)
    
    
    #### Extract components from files' names ####
    file_name <- basename(filename)
    participant_data <- stringr::str_match(
      file_name, 
      "^(\\d+)_F31_([MF])_(AQ|EXT)(\\d)_([0-9]{8})_.*_(\\d+)_(\\d+)_(\\d+)\\.xlsx$"
    )
    
    if (any(is.na(participant_data))) {
      warning(paste("Mismatch filename pattern in:", file_name))
      return(NULL)
    }
    
    # Extract participant metadata
    id <- participant_data[2]              # such as "3010"
    sex <- participant_data[3]             # "M" or "F"
    task_type <- participant_data[4]       # "AQ" or "EXT"
    task_version <- participant_data[5]    # "1" or "2"
    collection_date <- participant_data[6] %>% lubridate::mdy() # such as "09072021" and then to "2021-09-07"
    
    
    #### Determining row numbers ####
    # I set these based on the zoom recording of our meeting but these should be checked and adjusted if needed. 
    
    metric_rows <- list(
      
      # Segments' info
      start_number = 1,        
      start_event = 2,
      start_time = 3,
      end_event = 4,         
      end_time = 5,           
      segment_duration = 6,  
      
      # Mean metrics
      mean_hr = 9,            
      mean_ibi = 11,           
      n_rs_found = 12,      
      
      # ECG timing
      first_ecg_r = 17,       
      last_ecg_r = 18,       
      # first_r_to_l= ?, # not quite sure what is this
      
      # Time Domain metrics
      sdnn = 20,              
      avnn = 21,             
      rmssd = 22,             
      nn50 = 23,              
      pnn50 = 24             
    )
    
    # Number of segments 
    n_segments <- ncol(d) - 1 # label column eliminated (mostly 17 segments)
    
    
    #### Creating base data frame with the metadata ####
    
    result_data <- data.frame(
      id = id,
      sex = sex,
      task_type = task_type, 
      task_version = task_version,
      collection_date = collection_date,
      source_file = file_name,
      data_sheet = segment_sheet_name,
      n_segments = n_segments,
      stringsAsFactors = FALSE
    )
    
    
    #### Extracting TRiggers ####
    # I am not sure about the way that triggers are presented in the MindWare outputs, so this section should be completed later

    
    #### Extract all metrics for each segment ####
    for(metric_name in names(metric_rows)) {
      
      row_num <- metric_rows[[metric_name]]
      
      # Extracting values of segments
      metric_values <- tryCatch({as.numeric(d[row_num, 2:ncol(d)])}, 
                                error = function(e) {values <- as.character(d[row_num, 2:ncol(d)]) # suitable for strings like "NA"
        
                                sapply(values, function(x) {
                                  if(x == "N/A") return(NA)
                                  if(grepl("^-?\\d+\\.?\\d*$", x)) return(as.numeric(x))
                                  return(NA)
                                })
      })
      
      # Adding each segment's value as a separate colunm
      for(seg in 1:n_segments) {
        if(length(metric_values) >= seg) {
          col_name <- paste0("seg", seg, "_", metric_name)
          result_data[[col_name]] <- metric_values[seg]
        } else {
          col_name <- paste0("seg", seg, "_", metric_name)
          result_data[[col_name]] <- NA
        }
      }
    }
    
    
    #### Calculating additional summary metrics ####
    key_metrics <- c("mean_hr", "rmssd", "sdnn")
    
    for(metric in key_metrics) {
      
      # Getting segment values for this metric
      segment_cols <- paste0("seg", 1:n_segments, "_", metric)
      values <- sapply(segment_cols, function(col) {
        if(col %in% names(result_data)) result_data[[col]] else NA
      })
      
      # overall mean and variablinity
      result_data[[paste0("overall_", metric, "_mean")]] <- mean(values, na.rm = TRUE)
      result_data[[paste0("overall_", metric, "_sd")]] <- sd(values, na.rm = TRUE)
    }
    
    
    #### Data quality checks ####
    # Usable segments percentage (those with non-NA mean_hr)
    
    hr_cols <- paste0("seg", 1:n_segments, "_mean_hr")
    hr_values <- sapply(hr_cols, function(col) {
      if(col %in% names(result_data)) result_data[[col]] else NA
    })
    
    result_data$pct_usable_segments <- mean(!is.na(hr_values)) * 100
    result_data$total_usable_segments <- sum(!is.na(hr_values))
    
    cat("Processed file:", file_name, "- Segments:", n_segments, 
        "- Usable:", result_data$total_usable_segments, "\n")
    
    return(result_data)
    
  }, error = function(e) {
    warning(paste("Error reading file:", filename, "-", e$message))
    return(NULL)
  })
}

#### Main Processing Pipeline ####

cat("=== Starting The Proccess ===\n")

file_list <- list.files(path = "D:/Research/FPS (Ruvvy RLab)/Codes/RawData/MindWare_Files", 
                        pattern = "*.xlsx", full.names = TRUE)

cat("Found", length(file_list), "MindWare files to process\n")

# Process all files with error handling
hrv_data_list <- map(file_list, safely(read_mindware_files))
hrv_raw <- hrv_data_list %>% map("result") %>% compact() %>% bind_rows()


#### Data cleaning and QC ####

cat("\n=== Data Cleaning & QC ===\n")

if(nrow(hrv_raw) > 0) {
  
  # Checking for missing IDs
  missing_ids <- sum(is.na(hrv_raw$id))
  cat("Files with missing IDs:", missing_ids, "\n")
  
  
  #### Outlier detection using median absolute deviation ####
  
  detect_physiological_outliers <- function(data) {

    ecg_metrics <- c("overall_mean_hr_mean", "overall_rmssd_mean", "overall_sdnn_mean")
    outliers_detected <- list()
    
    for(metric in ecg_metrics) {
      if(metric %in% names(data)) {
        mad_val <- mad(data[[metric]], na.rm = TRUE)
        median_val <- median(data[[metric]], na.rm = TRUE)
        outlier_idx <- which(abs(data[[metric]] - median_val) > 3 * mad_val)
        outliers_detected[[metric]] <- outlier_idx
        
        if(length(outlier_idx) > 0) {
          cat("Outliers detected in", metric, ":", length(outlier_idx), "cases\n")
        }
      }
    }
    return(outliers_detected)
  }
  
  # Reporting outliers
  outliers <- detect_physiological_outliers(hrv_raw)
  
  
  #### Selecting key variables ####
  
  hrv_clean <- hrv_raw %>% select(
      # metadata 
      id, sex, task_type, task_version, collection_date, source_file,
      # overall summary metrics (for primary analysis, I should double check the metrics' labels)
      overall_mean_hr_mean, overall_rmssd_mean, overall_sdnn_mean,
      # data quality indicators
      pct_usable_segments, total_usable_segments, n_segments
    ) %>%
    # Removing rows with missing key metrics and IDs
    filter(!if_all(c(overall_mean_hr_mean, overall_rmssd_mean), is.na)) %>%
    filter(!is.na(id)) %>%
    # Physiological plausibility range
    filter(overall_mean_hr_mean >= 40 & overall_mean_hr_mean <= 180,  
           overall_rmssd_mean >= 10 & overall_rmssd_mean <= 200)      
  
  
  #### Data quality summary####
  
  cat("\n=== Data Qually Ssummary ===\n")
  cat("Original files processed:", nrow(hrv_raw), "\n")
  cat("After quality filtering:", nrow(hrv_clean), "\n")
  cat("Files removed:", nrow(hrv_raw) - nrow(hrv_clean), "\n")
  cat("Unique participants:", n_distinct(hrv_clean$id), "\n")
  
  # Data quality metrics
  quality_summary <- hrv_clean %>%
    summarise(
      avg_usable_segments = mean(total_usable_segments, na.rm = TRUE),
      pct_high_quality = mean(pct_usable_segments >= 80) * 100,
      complete_cases = sum(!is.na(overall_mean_hr_mean) & !is.na(overall_rmssd_mean))
    )
  
  cat("Average usable segments:", round(quality_summary$avg_usable_segments, 1), "\n")
  cat("High quality files (≥ 80% usable):", round(quality_summary$pct_high_quality, 1), "%\n")
  
  # Saving cleaned data
  write_csv(hrv_clean, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Cleaned_ECG_Data.csv")
  cat("Cleaned data saved: Cleaned_ECG_Data.csv\n")
  
} else {
  cat("No data processed \n")
  stop("Processing failed - no data available")
}


#### Merging with subjective Data ####

cat("\n=== Merging with subjective Data ===\n")

# Loading the subjective data !!!! I assumed the name of the file is "subjective_data.xlsx" but I hsould check and refine this as needed)
subjective_data <- read_excel("D:/Research/FPS (Ruvvy RLab)/Codes/RawData/subjective_data.xlsx") %>%
  clean_names() %>%
  # Selecting subjective variables
  select(
    id,
    trauma_exposure = htq_total,           # Harvard Trauma Questionnaire total
    ptsd_total = ucla_total,               # UCLA PTSD Reaction Index total
    anxiety_total = scared_total,          # SCARED total score
    ucla_irritable,                        # Irritability/anger item
    ucla_concentrate,                      # Concentration problems item  
    ucla_sleep,                            # Sleep difficulties item
    ucla_danger,                           # Hypervigilance/danger item
    pds_total,                             # Pubertal Development Scale
    age
  )

# Merging
analysis_data <- hrv_clean %>%
  left_join(subjevtive_data, by = "id") %>%
    mutate(
    sex = as.factor(sex),
    task_type = as.factor(task_type),
    hyperarousal_score = ucla_irritable + ucla_concentrate + ucla_sleep + ucla_danger,
    # Ensuring key variables' format
    across(c(age, pds_total, trauma_exposure, ptsd_total, anxiety_total), as.numeric)
  )


#### Missing data (based on the word file that Dr. Grasser sent me)####

cat("\n=== Missing Data ===\n")

# Checking missing data patterns
missing_summary <- analysis_data %>%
  summarise(across(everything(), ~sum(is.na(.))/n())) %>%
  pivot_longer(everything(), names_to = "variable", values_to = "pct_missing") %>%
  arrange(desc(pct_missing))

print(missing_summary)

# Applying missing data threshold
variables_to_keep <- missing_summary %>% filter(pct_missing < 0.20) %>% pull(variable)
analysis_data_final <- analysis_data %>% select(all_of(variables_to_keep))

cat("Remained variables (< 20% missing data):", length(variables_to_keep), "/", ncol(analysis_data), "\n")


#### Final Data Export ####

cat("\n=== Final Data Export ====\n")

# Saving final analysis dataset
write_csv(analysis_data_final, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Final_Analysis_Dataset.csv")

# Creating codebook for documentation
codebook <- data.frame(
  Variable = names(analysis_data_final),
  Description = c(
    "Participant ID",
    "Sex",
    "Task type (AQ/EXT)", 
    "Task version",
    "Data collection date",
    "Source filename",
    "Mean HR across segments (bpm)",
    "Mean RMSSD across segments (ms)",
    "Mean SDNN across segments (ms)", 
    "Percentage of usable segments",
    "Number of usable segments",
    "Total segments available",
    "Trauma exposure count (HTQ)",
    "PTSD symptom severity (UCLA)",
    "Anxiety symptom severity (SCARED)",
    "Hyperarousal symptoms (UCLA subscale)",
    "Pubertal development score",
    "Age"
  )[1:length(names(analysis_data_final))]
)

write_csv(codebook, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Data_Codebook.csv")

cat("Final analysis dataset saved: Final_Analysis_Dataset.csv\n")
cat("Data codebook saved: Data_Codebook.csv\n")
cat("Processing completed.\n")
cat("\nDataset dimensions:", nrow(analysis_data_final), "rows ×", ncol(analysis_data_final), "columns\n")
cat("Unique participants:", n_distinct(analysis_data_final$id), "\n")



