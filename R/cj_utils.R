name_cleaner <- function(df){
  names(df) <- tolower(names(df))
  names(df) <- str_replace_all(names(df), "\\r\\n", "")
  names(df) <- str_replace_all(names(df), "\\\"", "")
  names(df) <- str_replace_all(names(df), "\\.$", "")
  names(df) <- str_replace_all(names(df), "\\.", "_")
  names(df) <- str_replace_all(names(df), "\\-", "_")
  names(df) <- str_replace_all(names(df), "\\%", "pct")
  names(df) <- str_replace_all(names(df), "\\?$", "")
  names(df) <- str_replace_all(names(df), "\\+", "")
  names(df) <- str_replace_all(names(df), "\\#", "")
  names(df) <- str_replace_all(names(df), "\\(|\\)|,", "")
  names(df) <- str_replace_all(names(df), "\\/", "_")
  names(df) <- str_replace_all(names(df), "\\s", "_")
  names(df) <- str_replace_all(names(df), "_+", "_")
  names(df) <- str_replace_all(names(df), "_$", "")
  names(df) <- str_replace_all(names(df), "^_", "")
  return(df)
}

excel_pct <- function(x) {
  z <- as.numeric(format(round(x, 3), nsmall = 2))
  return(z)
}

strip_currency <- function(x) {
  z <- as.numeric(str_replace_all(string = x,
                                  pattern = "[\\$,]",
                                  replacement = ""))
}

get_file <- function(file_path, filename_pattern = ".xlsx$|.xls$") {
  files <- file.info(list.files(path = file_path, 
                                pattern = filename_pattern, 
                                full.names = TRUE)) %>%
    rownames_to_column(var = "filename") %>%
    arrange(desc(mtime)) %>%
    mutate(filename = tolower(filename))
  return(files[1, 1])
}


read_two_row_headers <- function(path, 
                                 sheet = 1, 
                                 sep = " ") {
  # Read only the first two rows, no column names
  hdr <- read_excel(path, 
                    sheet = sheet, 
                    col_names = FALSE, 
                    n_max = 2)
  
  # Combine row 1 and row 2 into single header strings
  combined <- hdr %>%
    summarise(across(everything(), ~ str_c(., collapse = sep))) |>
    unlist(use.names = FALSE)
  
  # Clean to snake_case
  cleaned <- combined |>
    name_cleaner()
  
  return(cleaned)
}