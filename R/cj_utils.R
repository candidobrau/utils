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
    arrange(desc(ctime)) %>%
    mutate(filename = tolower(filename))
  return(files[1, 1])
}
