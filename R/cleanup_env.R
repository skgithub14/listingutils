#' Clean-up global environment and search path
#'
#' Clears all objects from the global environment and restarts R
#'
#' Should be used after each report is run so R/R Studio is ready for the next
#' report
#'
#' @returns nothing
#' @export
#'
cleanup_env <- function () {
  rm(list = ls(envir = .GlobalEnv), envir = .GlobalEnv)
  .rs.restartR()
}
