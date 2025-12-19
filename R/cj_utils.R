# null coalescing operator
`%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

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

# Old get file
# get_file <- function(file_path, filename_pattern = ".xlsx$|.xls$") {
#   files <- file.info(list.files(path = file_path, 
#                                 pattern = filename_pattern, 
#                                 full.names = TRUE)) %>%
#     rownames_to_column(var = "filename") %>%
#     arrange(desc(mtime)) %>%
#     mutate(filename = tolower(filename))
#   return(files[1, 1])
# }

# New get_file
get_file <- function(file_path, filename_pattern = "\\.xlsx$|\\.xls$") {
  stopifnot(dir.exists(file_path))

  files <- list.files(
    path = file_path,
    pattern = filename_pattern,
    full.names = TRUE,
    ignore.case = TRUE
  )

  if (length(files) == 0) {
    cli::cli_abort(
      "No matching files found in {.path {file_path}} (pattern: {filename_pattern})"
    )
  }

  info <- file.info(files)
  info$filename <- rownames(info)

  info <- info |>
    tibble::as_tibble() |>
    arrange(desc(mtime))

  chosen <- info$filename[1]

  log_info(
    "get_file(): {nrow(info)} matching file(s) in '{file_path}' (pattern: '{filename_pattern}'); using newest: '{basename(chosen)}' ({format(info$mtime[1])})"
  )

  chosen
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

# Parse Benchmark part numbers in Oxide CPN and revision
parse_benchmark_part_number <- function(benchmark_part_number, 
                                        left_pad_revision = FALSE) {
  # 1) Extract CPN: 3 digits, dash, 7 digits
  cpn_raw <- str_extract(benchmark_part_number, "\\d{3}-\\d{7}")
  
  # If CPN parsing fails, recycle the original benchmark_part_number
  cpn <- if_else(is.na(cpn_raw) & !is.na(benchmark_part_number),
                 benchmark_part_number,
                 cpn_raw
  )
  
  # 2) Extract version chunk: V + digits, anywhere like -V001-, -V1_, etc.
  revision_string <- str_extract(benchmark_part_number,
                                 "(?<=^|-)V\\d+(?=$|[-_])"
  )
  
  # Just the digits
  revision_digits <- str_extract(revision_string, "\\d+")
  
  # 3) Apply revision rules:
  #    - no Vxxx suffix        -> revision 1
  #    - V1 / V01 / V001       -> 1
  #    - V002                  -> 2, etc.
  #    - NA benchmark_part_number     -> NA revision
  revision_int <- case_when(is.na(benchmark_part_number) ~ NA_integer_,
                            is.na(revision_digits)       ~ 1L,
                            TRUE                         ~ as.integer(revision_digits)
  )
  
  # 4. Optional left-padding to width 3
  revision <- if (left_pad_revision) {
    # Keep NA as NA_character_
    ifelse(
      is.na(revision_int),
      NA_character_,
      stringr::str_pad(revision_int, width = 3, pad = "0")
    )
  } else {
    revision_int
  }
  
  # Return as tibble so mutate() can splice columns
  tibble(
    cpn      = cpn,
    revision = revision
  )
}


parse_oxide_serial_number <- function(serial_number, keep_original = TRUE) {
  # serial_number: vector of serial number strings
  # keep_original: if TRUE, include the original in the output tibble

  tmp <- tibble::tibble(sn_raw = serial_number)

  parsed <- tmp |>
    tidyr::separate(
      sn_raw,
      into   = c("schema", "cpn", "revision", "serial"),
      sep    = ":",
      fill   = "left",
      extra  = "drop",
      remove = FALSE
    )

  if (!keep_original) {
    parsed <- dplyr::select(parsed, -sn_raw)
  } else {
    parsed <- dplyr::rename(parsed, serial_number = sn_raw)
  }

  parsed
}

write_excel_to_desktop <- function(x, filename, overwrite = TRUE) {
  # x must be a named list: name -> data.frame/tibble

  # Expand Desktop path safely
  desktop <- path.expand("~/Desktop")

  if (!dir.exists(desktop)) {
    cli::cli_abort("Desktop directory does not exist: {desktop}")
  }

  out_path <- file.path(desktop, filename)

  cli::cli_alert_info(glue::glue("Creating Excel workbook: {out_path}"))

  openxlsx::saveWorkbook(
    x,
    file = out_path,
    overwrite = overwrite
  )

  cli::cli_alert_success(glue::glue("Excel file written to Desktop: {out_path}"))

  invisible(out_path)
}

# Function to add a sheet to a wb object with common features
add_wb_sheet <- function(wb_object,
                         wb_sheet_name,
                         tab_color,
                         header_style,
                         header_names = NULL,
                         col_widths,
                         data_df,
                         date_cols_style = NULL,
                         date_cols = NULL) {
  # Add a worksheet
  openxlsx::addWorksheet(
    wb = wb_object,
    sheetName = wb_sheet_name,
    zoom = 100,
    tabColour = tab_color
  )

  cli::cli_alert_success(glue(
    "Created {wb_sheet_name} tab"
  ))

  # Rename columns for display
  if (!is.null(header_names)) {
    # Only rename columns that actually exist
    valid_renames <- header_names[header_names %in% names(data_df)]

    if (length(valid_renames) == 0) {
      cli::cli_inform(glue::glue(
        "No valid header names found for renaming in {wb_sheet_name}; keeping defaults."
      ))
      data_to_write <- data_df
    } else {
      data_to_write <- data_df |>
        dplyr::rename(!!!valid_renames)
    }
  } else {
    data_to_write <- data_df
  }

  # Add data to worksheet
  openxlsx::writeData(
    wb = wb_object,
    sheet = wb_sheet_name,
    x = data_to_write,
    withFilter = TRUE
  )

  cli::cli_alert_success(glue(
    "Added data to {wb_sheet_name} tab: {nrow(data_df)} rows × {ncol(data_df)} columns"
  ))

  # Set column widths
  openxlsx::setColWidths(
    wb = wb_object,
    sheet = wb_sheet_name,
    cols = seq_along(data_df),
    widths = col_widths
  )

  # Add short date style to any date columns
  if (!is.null(date_cols) && !is.null(date_cols_style)) {
    # Allow either names or numeric indices
    if (is.character(date_cols)) {
      cols_idx <- match(date_cols, names(data_df))
      cols_idx <- cols_idx[!is.na(cols_idx)]
    } else {
      cols_idx <- date_cols
    }

    if (length(cols_idx) > 0) {
      # writeData puts header in row 1, data in rows 2:(nrow + 1)
      openxlsx::addStyle(
        wb         = wb_object,
        sheet      = wb_sheet_name,
        style      = date_cols_style,
        rows       = 2:(nrow(data_df) + 1),
        cols       = cols_idx,
        gridExpand = TRUE,
        stack      = TRUE
      )
    }
  }

  # Add header row style
  openxlsx::addStyle(
    wb = wb_object,
    sheet = wb_sheet_name,
    style = header_style,
    rows = 1,
    cols = seq_along(data_to_write),
    gridExpand = TRUE,
    stack = TRUE
  )

  invisible(wb_object)
}
