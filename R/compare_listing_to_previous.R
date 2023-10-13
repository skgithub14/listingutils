#' Compare a data frame listing to a previous version
#'
#' Compares a data frame listing to a previous version by annotating which rows
#' were deleted, changed, or are new.
#'
#' An error will be raised if a row cannot be categorized into deleted, changed,
#' or new.
#'
#' @param listing a data frame of with the new data
#' @param prev_report a data frame with the old data
#'
#' @returns a data frame
#' @export
#'
compare_listing_to_previous <- function (listing, prev_report) {

  if(nrow(listing) == 0) {
    return(listing)
  }

  if (is.null(prev_report)) {
    listing$`Discrepancy Change` <- "New"

  } else if (nrow(prev_report) == 0) {
    listing$`Discrepancy Change` <- "New"

  } else {
    comp <- compareDF::compare_df(df_new = listing,
                                  df_old = prev_report,
                                  group_col = colnames(listing),
                                  stop_on_error = F,
                                  keep_unchanged_rows = T)
    listing_c <- comp$comparison_df %>%
      dplyr::group_by(grp) %>%
      dplyr::mutate(
        grp_cnt = dplyr::row_number(),
        `Discrepancy Change` = dplyr::case_when(
          all("=" == chng_type) ~ "Old",
          all(grp_cnt == 1) & chng_type == "-" ~ "Deleted",
          all(grp_cnt == 1) & chng_type == "+" ~ "New",
          max(grp_cnt) == 2 & all(chng_type %in% c("+", "-")) ~ "Changed",
          TRUE ~ NA_character_
        )
      ) %>%
      dplyr::ungroup() %>%

      # check for unhandled Discrepancy Change cases
      {
        if (any(is.na(.$`Discrepancy Change`))) {
          stop("There are unhandled cases in `listing_c$`Discrepancy Change` marked as `NA`.")
        } else {
          .
        }
      } %>%

      # remove deleted records and records from the previous run
      dplyr::filter(`Discrepancy Change` != "Deleted" &
               !(`Discrepancy Change` == "Changed" & chng_type == "-") &
               !(`Discrepancy Change` == "Unchanged" & grp_cnt == 2))
  }

  return(listing)
}


#' Compare a data frame listing to a previous version
#'
#' Compares a data frame listing to a previous version by annotating which rows
#' were deleted, changed, or are new.
#'
#' An error will be raised if a row cannot be categorized into deleted, changed,
#' or new.
#'
#' @param listing a data frame of with the new data
#' @param prev_report a data frame with the old data
#' @param unique_row_id a string, the name of the column that links unique row
#' IDs between `listing` and `prev_report`, usually a concatenation of other
#' column values that collectively make a unique ID; default is `"Unique_Row_ID"`
#' @param finding_col a string, the name of the column that describes row
#' discrepancies or issues, default is `"Finding"`
#' @param change_col a string, the name of the column to record changes; default
#' is `"Discrepancy Change"`
#'
#' @returns a data frame
#' @export
#'
categorize_finding_change <- function (listing,
                                       prev_report,
                                       unique_row_id = "Unique_Row_ID",
                                       finding_col = "Finding",
                                       change_col = "Discrepancy Change") {

  if(nrow(listing) == 0) return(listing)

  if (is.null(prev_report)) {
    listing[[change_col]] <- "New"

  } else if (nrow(prev_report) == 0) {
    listing[[change_col]] <- "New"

  } else {
    prev_report_cols <- dplyr::select(prev_report,
                                      tidyselect::all_of(c(unique_row_id,
                                                           finding_col)))

    unique_row_ids_in_prev <- unique(prev_report[[unique_row_id]])

    listing <- listing %>%
      dplyr::left_join(prev_report_cols,
                       by = unique_row_id,
                       suffix = c("", ".prev")) %>%
      dplyr::mutate(
        "{change_col}" := dplyr::case_when(
          !(!!rlang::sym(unique_row_id)) %in% unique_row_ids_in_prev ~ "New",
          (!!rlang::sym(finding_col)) == (!!rlang::sym(paste0(finding_col, ".prev"))) ~ "Old",
          (!!rlang::sym(finding_col)) %neqna% (!!rlang::sym(paste0(finding_col, ".prev"))) ~ "Changed",
          TRUE ~ NA_character_
        )
      ) %>%
      dplyr::select(-tidyselect::all_of(paste0(finding_col, ".prev")))

  }

  return(listing)
}
