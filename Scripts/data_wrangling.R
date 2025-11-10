# Last Update: 11/10/2020

##### Loading Packages #####
# The required packages are already installed
library(tidyverse) 
library(readxl)    
library(janitor)   
library(purrr) 
library(stringr)
library(lubridate)

##### Function: Reading Mindware Files ##### 
read_mindware_files <- function(filename) {
  tryCatch({

    d <- read_excel(filename)
    
    # Extract components from filenames
    file_name <- basename(filename)
    participant_data <- stringr::str_match(
      file_name, 
      "^(\\d+)_F31_([MF])_(AQ|EXT)(\\d)_([0-9]{8})_.*_(\\d+)_(\\d+)_(\\d+)\\.xlsx$"
    )
    
    if (any(is.na(participant_data))) {
      warning(paste("Missmatch filename Patterin in:", file_name))
      return(NULL)
    }
    
    id <- participant_data[2]              # such as "3010"
    sex <- participant_data[3]             # "M" or "F"
    task_type <- participant_data[4]       # "AQ" or "EXT"
    task_version <- participant_data[5]    # "1" or "2"
    collection_date <- participant_data[6] %>% lubridate::mdy() # such as "09072021" and then to "2021-09-07"

    # Creating a data frame with the metrics
    metrics_df <- data.frame(
      # Time and Basic Metrics
      time = t(as.vector(d[36, 2:ncol(d)])) %>% as.numeric(),
      epoch = 1:length(t(as.vector(d[36, 2:ncol(d)]))),
      
      # Heart Rate Metrics
      hr = t(as.vector(d[37, 2:ncol(d)])) %>% as.numeric(),
      mean_hr = t(as.vector(d[37, 2:ncol(d)])) %>% as.numeric(),
      
      # HRV Time-Domain Metrics
      rmssd = t(as.vector(d[52, 2:ncol(d)])) %>% as.numeric(),
      sdnn = t(as.vector(d[51, 2:ncol(d)])) %>% as.numeric(),
      pnn50 = t(as.vector(d[53, 2:ncol(d)])) %>% as.numeric(),

      # IBI and Interval Metrics
      ibi = t(as.vector(d[39, 2:ncol(d)])) %>% as.numeric(),
      mean_ibi = t(as.vector(d[39, 2:ncol(d)])) %>% as.numeric(),
      
      # Frequency Domain Metrics
      hf_power = tryCatch({
        t(as.vector(d[54, 2:ncol(d)])) %>% as.numeric()
      }, error = function(e) rep(NA, length(t(as.vector(d[36, 2:ncol(d)]))))),
      
      lf_power = tryCatch({
        t(as.vector(d[55, 2:ncol(d)])) %>% as.numeric()
      }, error = function(e) rep(NA, length(t(as.vector(d[36, 2:ncol(d)]))))),
      
      vlf_power = tryCatch({
        t(as.vector(d[56, 2:ncol(d)])) %>% as.numeric()
      }, error = function(e) rep(NA, length(t(as.vector(d[36, 2:ncol(d)]))))),
      
      lf_hf_ratio = tryCatch({
        t(as.vector(d[57, 2:ncol(d)])) %>% as.numeric()
      }, error = function(e) rep(NA, length(t(as.vector(d[36, 2:ncol(d)]))))),
      
      # Data from the filename
      id = id,
      sex = sex,
      task_type = task_type,
      task_version = task_version,
      collection_date = collection_date_parsed,
      source_file = file_name,
      
      stringsAsFactors = FALSE
    )
    
    # Remove rows where time is NA (empty columns in original data)
    metrics_df <- metrics_df %>% filter(!is.na(time))
    
    return(metrics_df)
    
  }, error = function(e) {
    warning(paste("Error reading file:", filename, "-", e$message))
    return(NULL)
  })
}


# # Extract components from files' names
# participant_data <- stringr::str_match(file_name, "^(\\d+)_F31_([MF])_(AQ|EXT)(\\d)_([0-9]{8})_.*_(\\d+)_(\\d+)_(\\d+)\\.xlsx$")
# 
# # Extract components
# id        <- participant_data[2]  # "3010"
# sex       <- participant_data[3]  # "M"
# task_type      <- participant_data[4]  # "AQ"
# task_version     <- participant_data[5]  # "1"
# collection_date  <- participant_data[6]  # "09072021"



