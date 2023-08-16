#' Bind multiple data frames by rows ignore empty
#'
#' This is a improved version of [dplyr::bind_rows()] function.
#' To handle case that empty csv data file has different data structure,
#' this [bind_row_non_empty()] function omit data frame which are empty.
#'
#' @param ... data frames to combine
#'
#' @returns a data frame
#' @export
#'
bind_rows_non_empty <- function(...) {
  arguments <- list(...)
  non_empty <- list()
  for (arg in arguments) {
    if (dim(arg)[1] >= 1) {
      non_empty <- c(non_empty, list(arg))
    }
  }
  res <- do.call(dplyr::bind_rows, non_empty)
  return(res)
}
