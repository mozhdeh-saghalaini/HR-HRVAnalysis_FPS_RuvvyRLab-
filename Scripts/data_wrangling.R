##### Last Update: 12/2/2025 ####

# Authored by Mozhdeh Saghalaini: m.saghalaini@gmail.com
# ------------------------------------------------------------------------------
# ECG Data Wrangling and Processing

## Description:
# This code processes raw MindWare ECG files from the FPS paradigm
# Extracts HRV metrics from "HRV Stats" sheet 
# Each participant will have separate rows for ACQ and EXT tasks (two rows for each participant)

## Input:
# MindWare outputs (.xlsx) listed in the MindWare_Files folder

## Output:
# Cleaned_ECG_Data.csv - processed HR/HRV metrics with quality indicators
# ------------------------------------------------------------------------------

##### Loading Packages #####

library(tidyverse) # data manipulation and visualization
library(readxl)    # Reading .xlsx files   
library(purrr)     # functional programming
library(stringr)   # String manipulation
library(lubridate) # Data handling
library(janitor)   # data cleaning process
library(robustbase)# Robust outlier detection


#### Function: Reading Mindware Files #### 

read_mindware_files <- function(filename) {
  tryCatch({
    
    #### Sheet Selection ####
    segment_sheet_name <- "HRV Stats" 
    
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
    # adjusted based on an actual MindWare output (11/28/2025)
    
    metric_rows <- list(
      
      segment_duration = 9,

      mean_hr = 56,            
      # mean_ibi = 59,           
      # n_rs_found = 60,      
      
      sdnn = 65,             
      rmssd = 67
      # avnn = 67,
      # nn50 = 69,              
      # pnn50 = 70             
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
      
      all_values <- as.character(d[row_num, 2:ncol(d)])
      non_na_values <- which(!is.na(all_values) & all_values != "" & all_values != "N/A")
      
      if(length(non_na_values) > 0) {
        last_real_segment <- max(non_na_values)
        metric_values <- all_values[1:last_real_segment]
      } else {
        metric_values <- rep(NA, n_segments)
      }
      
      metric_values <- sapply(metric_values, function(x) {
        if (is.na(x)) return(NA)
        if(x == "N/A") return(NA)
        if(grepl("^-?\\d+\\.?\\d*$", x)) return(as.numeric(x))
        return(NA)
      })

      
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
    
    ## Overall means of key metrics
    key_metrics <- c("mean_hr", "rmssd", "sdnn")
    
    for(metric in key_metrics) {
      
      segment_cols <- paste0("seg", 1:n_segments, "_", metric)
      values <- sapply(segment_cols, function(col) {
        if(col %in% names(result_data)) result_data[[col]] else NA
      })
      
      result_data[[paste0("overall_", metric, "_mean")]] <- mean(values, na.rm = TRUE)
      result_data[[paste0("overall_", metric, "_sd")]] <- sd(values, na.rm = TRUE)
    }
    
    ## Phasic metrics for the paradigm (late/early EXT/ACQ)
    all_segment_numbers <- 1:n_segments
    
    baseline_segment <- paste0("seg1_", metric)
    task_segments <- all_segment_numbers[all_segment_numbers > 1]
    
    # Early and late phases
    split_point <- ceiling(length(task_segments) / 2)
    early_task_segments <- task_segments[1:split_point]
    late_task_segments <- task_segments[(split_point + 1):length(task_segments)]
    
    early_cols <- paste0("seg", early_task_segments, "_", metric)
    late_cols <- paste0("seg", late_task_segments, "_", metric)
    all_task_cols <- paste0("seg", task_segments, "_", metric)
    
    # Phasic averages
    result_data[[paste0(metric, "_baseline")]] <- result_data[[baseline_segment]]
    
    result_data[[paste0(metric, "_task_early")]] <- rowMeans(
      result_data[early_cols], na.rm = TRUE
    )
    
    result_data[[paste0(metric, "_task_late")]] <- rowMeans(
      result_data[late_cols], na.rm = TRUE
    )
    
    # Calculate indices based on Early vs. Late and Early vs. Baseline and Late vs. Baseline
    result_data[[paste0(metric, "_task_change")]] <- 
      result_data[[paste0(metric, "_task_late")]] - result_data[[paste0(metric, "_task_early")]]
    
    result_data[[paste0(metric, "_reactivity_early")]] <- 
      result_data[[paste0(metric, "_task_early")]] - result_data[[paste0(metric, "_baseline")]]
    
    result_data[[paste0(metric, "_reactivity_late")]] <- 
      result_data[[paste0(metric, "_task_late")]] - result_data[[paste0(metric, "_baseline")]]
    
    
    
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
  
  
  #### Outlier detection using median absolute deviation (MAD) ####
  
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
  
  # This approach drops only number of segemnts and the rest of the variables are remained:
  hrv_clean <- hrv_raw %>%
    select(-n_segments) %>% 
    filter(!if_all(c(overall_mean_hr_mean, overall_rmssd_mean), is.na)) %>%
    filter(!is.na(id)) %>%
    filter(overall_mean_hr_mean >= 40 & overall_mean_hr_mean <= 180,
           overall_rmssd_mean >= 10 & overall_rmssd_mean <= 200)
  
  # This one selects only a few of variables:
  # hrv_clean <- hrv_raw %>% select(
  #     # metadata 
  #     id, sex, task_type, task_version, collection_date, source_file,
  #     # overall summary metrics (for primary analysis, I should double check the metrics' labels)
  #     overall_mean_hr_mean, overall_rmssd_mean, overall_sdnn_mean,
  #     # data quality indicators
  #     pct_usable_segments, total_usable_segments, n_segments
  #   ) %>%
  #   # Removing rows with missing key metrics and IDs
  #   filter(!if_all(c(overall_mean_hr_mean, overall_rmssd_mean), is.na)) %>%
  #   filter(!is.na(id)) %>%
  #   # Physiological plausibility range
  #   filter(overall_mean_hr_mean >= 40 & overall_mean_hr_mean <= 180,  
  #          overall_rmssd_mean >= 10 & overall_rmssd_mean <= 200)      
  
  
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
  
  # Task distribution
  task_distribution <- hrv_clean %>%
    count(task_type)
  cat("Task distribution:\n")
  print(task_distribution)
  
  # Saving cleaned data
  write_csv(hrv_clean, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Cleaned_ECG_Data.csv")
  cat("Cleaned data saved: Cleaned_ECG_Data.csv\n")
  
} else {
  cat("No data processed \n")
  stop("Processing failed - no data available")
}


