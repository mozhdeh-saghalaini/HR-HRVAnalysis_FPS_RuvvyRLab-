# Last Update: 11/10/2020

##### Loading Packages #####

library(tidyverse) 
library(readxl)    
library(purrr) 
library(stringr)
library(lubridate)

#### Function: Reading Mindware Files#### 

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
    
    id <- participant_data[2]              # such as "3010"
    sex <- participant_data[3]             # "M" or "F"
    task_type <- participant_data[4]       # "AQ" or "EXT"
    task_version <- participant_data[5]    # "1" or "2"
    collection_date <- participant_data[6] %>% lubridate::mdy() # such as "09072021" and then to "2021-09-07"
    
    #### Determining Row Numbers ####
    
    # I set this based on the zoom recording of our meeting but should be checked and adjusted if needed. 
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
      first_r_to_l= 17, # not quite sure what is this
      
      # Time Domain metrics
      sdnn = 20,              
      avnn = 21,             
      rmssd = 22,             
      nn50 = 23,              
      pnn50 = 24             
    )
    
    # Number of segments 
    n_segments <- ncol(d) - 1 # label column eliminated
    
    # Creating base data frame 
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
    
    #### Calculating additional metrics ####
    key_metrics <- c("mean_hr", "rmssd", "sdnn")
    
    for(metric in key_metrics) {
      
      # Getting segment values for this metric
      segment_cols <- paste0("seg", 1:n_segments, "_", metric)
      values <- sapply(segment_cols, function(col) {
        if(col %in% names(result_data)) result_data[[col]] else NA
      })
      
      # overall mean
      result_data[[paste0("overall_", metric, "_mean")]] <- mean(values, na.rm = TRUE)
      result_data[[paste0("overall_", metric, "_sd")]] <- sd(values, na.rm = TRUE)
    }
    
    #### Data quality checks ####
    # Usable segments percent (those with non-NA mean_hr)
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

#### Main Pipeline ####
file_list <- list.files(path = "D:/Research/FPS (Ruvvy RLab)/Codes/RawData/MindWare_Files", 
                        pattern = "*.xlsx", full.names = TRUE)

cat("Found", length(file_list), "MindWare files to process\n")

# Processing all files
hrv_data_list <- map(file_list, safely(read_mindware_files))

# Combining into one file
hrv_final <- hrv_data_list %>% map("result") %>% compact() %>% bind_rows()

#### Final checks and saving ####

if(nrow(hrv_final) > 0) {
  # Summary 
  cat("\n Processing summary:\n")
  cat("Total files processed:", nrow(hrv_final), "\n")
  cat("Unique participants:", n_distinct(hrv_final$id), "\n")
  cat("Average segments per file:", mean(hrv_final$n_segments, na.rm = TRUE), "\n")
  cat("Average usable segments:", mean(hrv_final$total_usable_segments, na.rm = TRUE), "\n")
  
  # Show column structure
  cat("\n Data structure:\n")
  cat("Metadata columns:", names(hrv_final)[1:8], "\n")
  cat("Overall metrics:", grep("^overall_", names(hrv_final), value = TRUE), "\n")
  cat("Sample segment columns:", names(hrv_final)[9:15], "...\n")
  
  # Save the final data
  write_csv(hrv_final, "D:/Research/FPS (Ruvvy RLab)/Codes/Output/Cleaned_Objective_HR_data.csv")
  
  cat("Cleaned_Objective_HR_data.csv created successfully!\n")

} else {
  cat("\n No data processed! \n")
}




