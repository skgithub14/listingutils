#' Get file Generated Date
#'
#' Return the date of a specific file generated date
#'
#' @param path absolute file path
#'
#' @returns a Date
#' @export
#'
get_file_date <- function(path){
  file <- file.info(path)
  return(as.Date(file$ctime))
}
