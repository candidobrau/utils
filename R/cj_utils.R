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

# Deal with annoying Excel column headers split across two rows
clean_report_names <- function(df,
                               pattern,
                               new_names,
                               strict = TRUE,
                               fixed = FALSE) {
  # df:        data frame / tibble
  # pattern:   pattern to match column names (passed to grep)
  # new_names: character vector of replacement names, in order
  # strict:    if TRUE, error when there are fewer matches than new_names
  # fixed:     passed to grep(); if TRUE, treat pattern as literal string
  
  duped_names <- names(df)
  
  # Find indices of columns whose names match the pattern
  idx <- grep(pattern, duped_names, fixed = fixed)
  n_matches <- length(idx)
  n_replacements <- length(new_names)
  
  if (n_matches == 0) {
    warning("clean_report_names(): no column names matched pattern: ", pattern)
    return(df)
  }
  
  if (n_replacements == 0) {
    warning("clean_report_names(): new_names is empty; nothing to do.")
    return(df)
  }
  
  if (strict && n_matches < n_replacements) {
    stop(
      "clean_report_names(): pattern '", pattern,
      "' matched only ", n_matches, " columns, but ",
      n_replacements, " replacement names were provided."
    )
  }
  
  # Only replace as many as we have both matches and new_names for
  k <- min(n_matches, n_replacements)
  
  duped_names[idx[seq_len(k)]] <- new_names[seq_len(k)]
  
  names(df) <- duped_names
  df
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