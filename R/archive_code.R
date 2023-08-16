#' Archive an R script
#'
#' Copies the an R script to a data archive folder
#'
#' @param code_fname a string, the file name of the code to archive
#' @param code_dir a string, the directory the code lives in
#' @param code_archive_dir a string, the directory the code should be copied to
#'
#' @returns nothing
#' @export
#'
archive_code <- function (code_fname, code_dir, code_archive_dir) {
  file.copy(from = file.path(code_dir, code_fname),
            to = file.path(code_archive_dir, code_fname),
            overwrite = TRUE,
            copy.date = TRUE)
}
