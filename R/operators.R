#' Operator for not equal comparison with NA handling
#'
#' Provides element-wise comparison of two vectors to determine if the values
#' are NOT equal. Comparison of two non-`NA` values is equivalent to the result
#' of `value1 == value2`.`NA` compared to a non-`NA` values will return `TRUE`
#' and `NA` compared to `NA` will return `FALSE`.
#'
#' @param e1,e2 vectors of equal length to compare that can be of any class
#'
#' @returns a vector of length equal to `e1` and `e2` containing logical values
#' of `TRUE` or `FALSE` only (no `NA` values).
#'
#' @export
#'
#' @examples
#' # the below will return `c(FALSE, TRUE, TRUE, FALSE)`
#' c(1, 1, 1, NA) %neqna% c(1, 2, NA, NA)
#'
`%neqna%` <- function(e1, e2) {
  (e1 != e2 |
     (is.na(e1) & !is.na(e2)) |
     (is.na(e2) & !is.na(e1))
   ) &
    !(is.na(e1) & is.na(e2))
}


#' Operator for equal comparison with NA handling
#'
#' Provides element-wise comparison of two vectors to determine if the values
#' are equal. Comparison of two non-`NA` values is equivalent to the result
#' of `value1 == value2`.`NA` compared to a non-`NA` values will return `TRUE`
#' and `NA` compared to `NA` will return `FALSE`.
#'
#' @param e1,e2 vectors of equal length to compare that can be of any class
#'
#' @returns a vector of length equal to `e1` and `e2` containing logical values
#' of `TRUE` or `FALSE` only (no `NA` values).
#'
#' @export
#'
#' @examples
#' # the below will return `c(FALSE, TRUE, TRUE, FALSE)`
#' c(1, 1, 1, NA) %eqna% c(1, 2, NA, NA)
#'
`%eqna%` <- function(e1, e2) {
  (e1 == e2 |
     (is.na(e1) & !is.na(e2)) |
     (is.na(e2) & !is.na(e1))
  ) &
    !(is.na(e1) & is.na(e2))
}
