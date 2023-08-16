#' #' Import data
#' #'
#' #' Import data sets within a folder based on the file type and file name
#' #'
#' #' @param source_dir a string of directory contains dataset files
#' #' @param filetype a string of file type, support sas7bdat, csv, rds, xlsx
#' #' @param filenames a vector contains file names
#' #' @param ... additional arguments passed to one of the reading funtions
#' #'
#' #' @returns a list of data frames
#' #'
#' #' @export
#' #'
#' import_data <- function(source_dir, filetype, filenames=c(), ...){
#'   if (length(filenames) == 0){
#'     pattern = filetype
#'   }
#'   else{
#'     pattern <- paste(paste0(filenames, filetype), sep="", collapse="|")
#'   }
#'
#'   source_files_dir <- list.files(path = source_dir, pattern=pattern)
#'   source <- list()
#'
#'   for (file in source_files_dir){
#'     source_file_dir <- file.path(source_dir, file)
#'     if (filetype == ".sas7bdat") {
#'       temp = haven::read_sas(source_file_dir, ...)
#'     }
#'     else if((filetype == ".csv") | (filetype == ".CSV")){
#'       temp = read.csv(source_file_dir, ...)
#'     }
#'     else if (filetype == ".rds"){
#'       temp = readRDS(source_file_dir, ...)
#'     }
#'     else if (filetype == ".xlsx"){
#'       temp = readxl::read_excel(source_file_dir, ...)
#'     }
#'     else if (filetype == ".xls"){
#'       temp = readxl::read_excel(source_file_dir, ...)
#'     }
#'     else{
#'       stop("The file type is not supported. The supported type is .csv and .sas7bdat")
#'     }
#'     source <- c(source, setNames(list(temp), tools::file_path_sans_ext(file)))
#'     assign(tools::file_path_sans_ext(file), temp)
#'   }
#'   return(source)
#' }
#'
