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


#' Apply normal rounding rules
#'
#' This is a function to round numeric vectors to the chosen number of
#' decimal places, following the rule "Round to nearest, ties away from
#' zero
#'
#' The function was adapted from
#' \url{https://stackoverflow.com/questions/12688717/round-up-from-5}
#'
#' @param x a numeric value to round
#' @param digits the number of digits to round to, integer
#'
#' @returns a rounded numeric value
#' @export
#'
round2 <- function(x, digits = 0) {
  posneg = sign(x)
  z = abs(x)*10^digits
  z = z + 0.5 + sqrt(.Machine$double.eps)
  z = trunc(z)
  z = z/10^digits
  z*posneg
}
