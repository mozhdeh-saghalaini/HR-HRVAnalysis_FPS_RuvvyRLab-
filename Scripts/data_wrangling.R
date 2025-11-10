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
    metrics <- data.frame(
      
      # time serie data
      time_data <- t(as.vector(d["TimeElapsed Time row", 2:ncol(d)])) %>% as.numeric(),
      hr_data <- t(as.vector(d["HR row", 2:ncol(d)])) %>% as.numeric(),
      rmssd_data <- t(as.vector(d["RMSSD row", 2:ncol(d)])) %>% as.numeric(),
      sdnn_data <- t(as.vector(d["SDNN row", 2:ncol(d)])) %>% as.numeric(),
      rsa_data <- t(as.vector(d["RSA row", 2:ncol(d)])) %>% as.numeric(),
      ibi_data <- t(as.vector(d["IBI row", 2:ncol(d)])) %>% as.numeric(),
      pnn50_data <- t(as.vector(d["pNN50 row", 2:ncol(d)])) %>% as.numeric(),
      
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



