#' Find ID duplicates
#'
#' Determines which observations of a data frame are duplicates based on the key
#' provided in the arguments. The function does not change the dimension of
#' the given data frame but instead creates a new column named by argument col.
#' The content of newly created column is specified by the message argument.
#'
#' @param df Data Frames to find duplicates
#' @param keys keys to determine duplicates
#' @param message content of column indicates duplication
#' @param col column name of generated column for duplication
#'
#' @returns a modified version of `df`
#'
#' @export
#'
find_duplicates <- function(df, keys, message, col = Discrepancy) {
  df <- df %>%
    dplyr::group_by(dplyr::across({{keys}})) %>%
    dplyr::mutate({{col}} := dplyr::row_number()) %>%
    dplyr::mutate({{col}} := dplyr::if_else({{col}} > 1,
                                            message,
                                            NA_character_)
    ) %>%
    dplyr::ungroup()

  return(df)
}
