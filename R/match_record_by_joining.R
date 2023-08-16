#' This function is designed to support DM data reconciliation report.
#' To reconciliate the date, EDC records and vendor records need to be matched
#' together. Most of the time, there is no direct id join key and instead
#' requires composite join key such as subject, visit, date, time, timepoint.
#' In the case, some of the composite key is not perfectly matches, this
#' function try different set of composite keys to match the records.
#'
#' @param df1 Data Frames to match (usually EDC data)
#' @param df2 Data Frame to be mathced (usually vendor data)
#' @param bys Sets of composite keyes used to match
#' @param distinct_keys a character vector of distinct keys
#' @param suffix a character vector of length two of column suffixes to append
#' to `df1` and `df2` columns, respectively
#' @param na_matches whether to match `NA` values, see the dplyr join functions
#'
#' @returns a data frame
#'
#' @export
#'
match_record_by_joining <- function(df1,
                                    df2,
                                    bys,
                                    distinct_keys,
                                    suffix = c(".edc", ".vendor"),
                                    na_matches="never") {


    colnames(df1)<-paste0(colnames(df1), suffix[1])
    colnames(df2) <- paste0(colnames(df2), suffix[2])
    col_num <- ncol(df1)

    key_sets <- c()
    df <- df1
    cnt <- 1

    for (i in 1:length(bys)) {

      by = bys[[i]]
      distinct_key = distinct_keys[[i]]
      temp <- paste0(names(by),suffix[1])
      by = paste0(by, suffix[2])
      distinct_key = paste0(distinct_key,suffix[2])
      names(by) = temp
      group = unname(by)
      key_set = unname(distinct_key)
      key_sets = c(key_sets, key_set)

      df <- df %>%
        dplyr::left_join(
          df2 %>% dplyr::select({{distinct_key}}) %>%
            dplyr::group_by_at(group) %>%
            dplyr::slice(1),
          by = {{by}},
          suffix = c("", as.character(cnt)),
          na_matches = na_matches,
          keep = TRUE
        )

      cnt <- cnt + 1
    }

    print(key_sets)

    actual <- tibble::tibble()

    for (key in unique(key_sets)) {

      if (dim(actual)[1] == 0) {

        actual <- df %>%
          dplyr::select(c((col_num + 1):ncol(df))) %>%
          dplyr::select(tidyselect::starts_with(key)) %>%
          dplyr::transmute({{key}} := dplyr::coalesce(
            !!!dplyr::select(., tidyselect::everything())
            )
          )

      } else {

        actual <- actual %>%
          dplyr::bind_cols(
            df %>%
              dplyr::select(c((col_num + 1):ncol(df))) %>%
              dplyr::select(tidyselect::starts_with(key)) %>%
              dplyr::transmute({{key}} := dplyr::coalesce(
                !!!dplyr::select(., tidyselect::everything())
                )
              )
          )

      }
    }

    print(names(actual))

    res <- df %>%
      dplyr::select(1:col_num) %>%
      dplyr::bind_cols(actual)

    return(res)
  }
