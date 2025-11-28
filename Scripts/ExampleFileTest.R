library(tidyverse) # data manipulation and visualization
library(readxl)    # Reading .xlsx files   
library(purrr)     # functional programming
library(stringr)   # String manipulation
library(lubridate) # Data handling
library(janitor)   # data cleaning process
library(robust)    # Robust outlier detection
library(naniar)    # For missing data visualization
library(mice)      # For multiple imputation
library(finalfit)  # For missing data tests
segment_sheet_name <- "HRV Stats" 

d <- read_excel("D:/Research/FPS (Ruvvy RLab)/Meeting Contents/meeting5-18Nov/Example MindWare Output File HRV.xlsx", sheet = segment_sheet_name)
#### Extract components from files' names ####
file_name <- "Example MindWare Output File HRV.xlsx"
file_name <- basename(file_name)
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


