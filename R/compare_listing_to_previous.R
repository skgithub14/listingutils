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

  if (is.null(prev_report)) {
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


