#' Covert dates in YYYYMMDD format to DDMMMYYYY
#'
#' Covert dates in YYYYMMDD format to DDMMMYYYY, not month will be in all caps.
#'
#' @param my_date a string or Date in YYYYMMDD format
#'
#' @returns a string in DDMMMYYYY format
#' @export
#'
format_date_DDMMMYYYY <- function(my_date) {
  my_date %>%
    as.Date(format = "%Y%m%d") %>%
    format(format = "%d%b%Y") %>%
    stringr::str_to_upper()
}
