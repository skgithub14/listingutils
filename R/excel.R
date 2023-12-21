#' Make an Excel workbook
#'
#' Creates an [openxlsx::createWorkbook()] object with one or more sheets and
#' populates header information and data for each sheet.
#'
#' @inheritParams setup_multi_sheet_workbook
#' @inheritParams populate_workbook
#'
#' @returns an `openxlsx` `Workbook` object
#' @export
#'
#' @seealso [setup_workbook()], [setup_multi_sheet_workbook()],
#' [populate_worksheet()], [populate_workbook()]
#'
create_excel_listing <- function (sheet_data, sheet_headers) {

  wb <- setup_multi_sheet_workbook(sheet_data)
  wb <- populate_workbook(wb = wb,
                          sheet_data = sheet_data,
                          sheet_headers = sheet_headers)
  return(wb)
}


#' Set-up Excel workbook with one worksheet
#'
#' Creates an [openxlsx::createWorkbook()] object, adds a worksheet, and sets
#' column widths to 19. This function should be used when creating workbooks
#' with one and only one sheet. For multi-sheet workbooks, see
#' [setup_multi_sheet_workbook()].
#'
#' No data is written to the workbook. The `data` argument is just for
#' configuring options for the correct number of columns.
#'
#' @param sheet_name string, the sheet name to create
#' @param data data frame of the data for `sheet_name`
#' @param col_widths a numeric value or `"auto"`, the column widths for all
#'  worksheets; default is `19`.
#'
#' @returns an `openxlsx` `Workbook` object
#' @export
#'
#' @seealso [setup_multi_sheet_workbook()]
#'
setup_workbook <- function (sheet_name, data, col_widths = 19) {

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, sheetName = sheet_name)
  openxlsx::setColWidths(wb,
                         sheet = sheet_name,
                         cols = 1:ncol(data),
                         widths = col_widths)
  return(wb)
}


#' Set-up Excel workbook with multiple worksheets
#'
#' Creates an [openxlsx::createWorkbook()] object and adds blank worksheets.
#'
#' No data is written to the workbook. The `data` argument is just for
#' configuring options for the correct number of columns.
#'
#' @param sheet_data a named list of data frames containing the data for each
#'  worksheet where the list names are the sheet names
#' @inheritParams setup_workbook
#'
#' @returns an `openxlsx` `Workbook` object
#' @export
#'
#' @seealso [setup_workbook()]
#'
setup_multi_sheet_workbook <- function (sheet_data, col_widths = 19) {

  wb <- openxlsx::createWorkbook()

  purrr::iwalk(sheet_data, ~ {
    openxlsx::addWorksheet(wb, sheetName = .y)
    openxlsx::setColWidths(wb,
                           sheet = .y,
                           cols = 1:ncol(.x),
                           widths = col_widths)
  })

  return(wb)
}


#' Writes header data to Excel workbook
#'
#' Writes the meta data to the top rows of an Excel workbook and formats them
#'
#' @param wb an `openxlsx` `Workbook` object
#' @param header a list of strings where each element corresponds to one row in
#' the header
#' @inheritParams setup_workbook
#'
#' @returns an `openxlsx` `Workbook` object
#' @export
#'
#' @seealso [add_worksheet_data()], [setup_workbook()]
#'
add_worksheet_header <- function (wb, sheet_name, header, data) {

  for (hd in seq_along(header)) {

    openxlsx::writeData(wb, sheet = sheet_name, x = header[[hd]], startRow = hd)
    openxlsx::mergeCells(wb, sheet = sheet_name, rows = hd, cols = 1:ncol(data))

    if (hd == 1) {
      study_style <- openxlsx::createStyle(fontSize = 15,
                                           textDecoration = "bold")
      openxlsx::addStyle(wb,
                         sheet = sheet_name,
                         study_style,
                         rows = 1,
                         cols = 1:ncol(data),
                         gridExpand = TRUE)
    }
  }

  return(wb)
}


#' Writes a data frame to an Excel workbook with formatting
#'
#' Writes a data frame to the specified sheet in an Excel workbook and applies
#' formatting, conditional formatting, freeze panes, and filters.
#'
#' @inheritParams add_worksheet_header
#' @inheritParams setup_workbook
#' @param start_row integer, the row number at which to write the column header
#' of `data`
#' @param discrepancy_col a string, the column name for discrepancy change
#' tracking between report versions, default is `"Discrepancy Change"`. If set
#' to `NULL`, conditional formatting for discrepancy change will not be applied.
#' @param match_text a string, text that identifies a data match in
#' `discrepancy_col`, default is `"Data match"`
#' @param finding_col a string, the column name for reconciliation findings,
#' default is `"Finding"`. If set to `NULL`, conditional formatting for
#' discrepancy change will not be applied.
#'
#' @returns a modified [openxlsx::createWorkbook()] object
#' @export
#'
#' @seealso [add_worksheet_header()], [setup_workbook()]
#'
add_worksheet_data <- function (wb,
                                sheet_name,
                                data,
                                start_row,
                                discrepancy_col = "Discrepancy Change",
                                match_text = "Data match",
                                finding_col = "Finding") {

  # set table header style
  headerStyle <- openxlsx::createStyle(border = "TopBottomLeftRight",
                                       wrapText = TRUE,
                                       halign = "center",
                                       valign = "center",
                                       textDecoration = "bold",
                                       fgFill = "#b3d9e5")

  # add the data
  openxlsx::writeData(wb,
                      sheet = sheet_name,
                      x = data,
                      colNames = TRUE,
                      startRow = start_row,
                      borders = "all",
                      headerStyle = headerStyle,
                      withFilter = T)

  ## table body formatting
  body_start <- start_row + 1
  body_end <- start_row + nrow(data)
  body_rows <- seq(body_start, body_end)

  openxlsx::freezePane(wb, sheet = sheet_name, firstActiveRow = body_start)

  # add styles depending on if a column has a Date class or not
  col_classes <- purrr::modify(colnames(data), ~ {
    cl <- class(data[[.]])[1]
    if (cl %in% c("POSIXlt", "POSIXt", "POSIXct")){
      cl <- "POSIX"
    }
    return(cl)
  })
  contentStyleGeneral <- openxlsx::createStyle(border = "TopBottomLeftRight",
                                               wrapText = TRUE,
                                               halign = "center",
                                               valign = "center")
  contentStyleDates <- openxlsx::createStyle(border = "TopBottomLeftRight",
                                             wrapText = TRUE,
                                             halign = "center",
                                             valign = "center",
                                             numFmt = "ddmmmyyyy")
  contentStyleDateTimes <- openxlsx::createStyle(border = "TopBottomLeftRight",
                                                 wrapText = TRUE,
                                                 halign = "center",
                                                 valign = "center",
                                                 numFmt = "ddmmmyyyy hh:mm:ss")
  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = contentStyleGeneral,
                     rows = body_rows,
                     cols = which(!col_classes %in% c("Date", "POSIX")),
                     gridExpand = TRUE)
  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = contentStyleDates,
                     rows = body_rows,
                     cols = which(col_classes == "Date"),
                     gridExpand = TRUE)
  openxlsx::addStyle(wb,
                     sheet = sheet_name,
                     style = contentStyleDateTimes,
                     rows = body_rows,
                     cols = which(col_classes == "POSIX"),
                     gridExpand = TRUE)

  ## conditional formatting
  # mark non-Data Matches in the Finding column
  if (!is.null(finding_col)) {
    if (finding_col %in% colnames(data)) {
      FindingNotMatchStyle <- openxlsx::createStyle(fontColour = "#FF0000",
                                                    bgFill = "#FFFF00")
      openxlsx::conditionalFormatting(wb,
                                      sheet = sheet_name,
                                      rows = body_rows,
                                      cols = which(colnames(data) == finding_col),
                                      rule = paste0("!=\"", match_text,"\""),
                                      type = "expression",
                                      style = FindingNotMatchStyle)
    }
  }

  # mark "Changes" in the Change column
  if (!is.null(discrepancy_col)) {
    if (discrepancy_col %in% colnames(data)) {
      ChangeChangeStyle <- openxlsx::createStyle(fontColour = "#FF0000",
                                                 bgFill = "#FFFF00")
      openxlsx::conditionalFormatting(wb,
                                      sheet = sheet_name,
                                      rows = body_rows,
                                      cols = which(colnames(data) == discrepancy_col),
                                      rule = "Change",
                                      type = "contains",
                                      style = ChangeChangeStyle)

      # mark "New" entries in the Change column
      ChangeNewStyle <- openxlsx::createStyle(fontColour = "#00FF00",
                                              bgFill = "#FFFF00")
      openxlsx::conditionalFormatting(wb,
                                      sheet = sheet_name,
                                      rows = body_rows,
                                      cols = which(colnames(data) == discrepancy_col),
                                      rule = "New",
                                      type = "contains",
                                      style = ChangeNewStyle)
    }
  }

  return(wb)
}


#' Populates a worksheet in an Excel workbook with header info and data
#'
#' Wraps [add_worksheet_header()] and [add_worksheet_data()] to populate an
#' entire sheet's worth of content.
#'
#' The sheet must already have been created in the workbook
#'
#' @inheritParams add_worksheet_header
#' @inheritParams setup_workbook
#'
#' @returns an `openxlsx` Workbook object
#' @export
#'
#' @seealso [add_worksheet_header()], [add_worksheet_data()], [setup_workbook()]
#'
populate_worksheet <- function (wb, sheet_name, data, header) {

  wb <- add_worksheet_header(wb = wb,
                             sheet_name = sheet_name,
                             header = header,
                             data = data)

  wb <- add_worksheet_data(wb = wb,
                           data = data,
                           sheet_name = sheet_name,
                           start_row = length(header) + 1)

  return(wb)
}


#' Populate workbook data for all worksheets
#'
#' Wraps [populate_worksheet()] to populate all sheets in an `openxlsx` Workbook
#' object.
#'
#' @inheritParams add_worksheet_header
#' @inheritParams setup_multi_sheet_workbook
#' @param sheet_headers a named list of lists where each name is a sheet
#' name in `wb` and each element is the header data for a worksheet in `wb`. If
#' `length(sheet_headers) == 1`, `sheet_headers` will be recycled for every
#' sheet in `wb`. `names(sheet_data)` must match `names(sheet_headers)`.
#'
#' @returns an `openxlsx` Workbook object
#' @export
#'
#' @seealso [populate_worksheet()], [setup_multi_sheet_workbook()]
#'
populate_workbook <- function (wb, sheet_data, sheet_headers) {

  # validate data is same length as headers or that length of sheet_headers == 1,
  #  (sheet_headers will be recycled if length == 1)
  if (length(sheet_data) != length(sheet_headers)) {
    if (length(sheet_headers) != 1) {
      stop("`length(sheet_headers)` must equal 1 or `length(sheet_data)`")
    }
  }

  # validate sheet names in the data match the headers, or that length of
  #  sheet_headers == 1 (sheet_headers will be recycled if length == 1)
  if (length(sheet_data) == length(sheet_headers) &
      length(sheet_headers) != 1) {

    in_data_only <- setdiff(names(sheet_data), names(sheet_headers))
    in_headers_only <- setdiff(names(sheet_headers), names(sheet_data))
    mismatched <- c(in_data_only, in_headers_only)

    if (length(mismatched) > 0) {
      stop("`names(sheet_data)` must match `names(sheet_headers)`, else `length(sheet_headers)` must be 1")
    }
  }

  # populate the worksheets
  if (length(sheet_headers) == 1) {
    for (sheet_name in names(sheet_data)) {
      wb <- populate_worksheet(wb = wb,
                               sheet_name = sheet_name,
                               data = sheet_data[[sheet_name]],
                               header = sheet_headers[[1]])

    }
  } else if (length(sheet_headers) > 1) {
    for (sheet_name in names(sheet_data)) {
      wb <- populate_worksheet(wb = wb,
                               sheet_name = sheet_name,
                               data = sheet_data[[sheet_name]],
                               header = sheet_headers[[sheet_name]])

    }
  } else {
    stop("length(sheet_headers) must be > 0")
  }

  return(wb)
}


#' Write a data frame or tibble to Excel worksheet with features and formatting
#'
#' @description Writes a data frame/tibble to an Excel worksheet with the
#'   following features:
#'  - Column labels/descriptions above the header
#'  - Filters
#'  - Freeze panes
#'  - Rounding
#'  - Conversion to percentage
#'  - Date and date/time formatting
#'  - Header and cell styling (alignment, wrap, borders)
#'  - Column widths (default, narrow, wide, and extra wide widths)
#'
#' @param wb the `openxlsx` `Workbook` object to modify
#' @param sheet_name a string, the sheet name in `wb` to modify
#' @param df the data frame or tibble to write
#' @param start_row a numeric value, the row to start writing the table; if
#'   `col_labels` is not `NULL` then the column labels will be written to this
#'   row with the column names the row below them, otherwise the column names
#'   will be written here
#' @param col_labels an optional named character vector of column labels where
#'   names are the column names in `df`; if present, column labels are added to
#'   the row directly above the column titles
#' @param col_label_row_ht an optional numeric value, if `col_labels` is not
#'   `NULL`, sets the height of the column label row; default is `60`
#' @param withFilter a logical value indicating if column filtering should be
#'   turned on; default is `TRUE`
#' @param freeze_header a logical value indicating whether the header row should
#'   be frozen; default is `TRUE`
#' @param last_frozen_col an optional string, the column name in `df` which
#'   should be the last frozen column; default is `NULL` in which case no
#'   columns will be frozen
#' @param num_cols an optional named numeric vector where the names are column
#'   names in `df` that are numeric columns and the values are number of decimal
#'   places each column should be rounded to; default is `NULL`, in which case
#'   any numeric columns not listed in `num_cols` or `perc_cols` will use
#'   Excel's general cell format
#' @param perc_cols an optional named numeric vector where the names are column
#'   names in `df` that are percentage columns and the values are number of
#'   decimal places each column should be rounded to; default is `NULL`, in
#'   which case any numeric columns not specified in `num_cols` or `perc_cols`
#'   will use Excel's general cell format. Note, any column specified in this
#'   argument will be multiplied by 100.
#' @param dateFormat a string specifying the format of Date class columns using
#'   the `openxlsx` specification; default is `"ddmmmyyyy"`
#' @param dateTimeFormat a string specifying the format of the date/time class
#'   columns (`c("POSIXlt", "POSIXt", "POSIXct")`) using the `openxlsx`
#'   specification; default is `"ddmmmyyyy hh:mm:ss"`
#' @param border a string specifying the cell borders for all cells using the
#'   `openxlsx` border specification; default is `"TopBottomLeftRight"`
#' @param wrapText a logical value indicating if all cells should wrap text;
#'   default is `TRUE`
#' @param halign,valign string values indicating the horizontal and vertical
#'   cell alignment for all cells using the `openxlsx` specification; default is
#'   `center`
#' @param fontSize a numeric value specifying the font size for all cells;
#'   default is `11`
#' @param header_textDecoration a string describing the text styling for the
#'   column name and label rows using the `openxlsx` specification; default is
#'   `"bold"`
#' @param header_fgFill a string containing a color hex code which will be used
#'   to color the column name and label rows; default is `"#b3d9e5"` (light
#'   blue)
#' @param default_wd a numeric value specifying the default column width;
#'   default is `20`; note if `narrow_wd`, `wide_wd`, or `xwide_wd` are `NULL`,
#'   they will be adjusted relative to `default_wd` (see their parameter details
#'   for more information)
#' @param narrow_cols,wide_cols,xwide_cols optional character vectors of column
#'   names which will be made narrower, wider or extra wide; default is `NULL`
#' @param narrow_wd,wide_wd,xwide_wd optional numeric values which specify the
#'   column widths for the columns specified in the `narrow_cols`, `wide_cols`,
#'   and `xwide_cols` arguments, respectively; default is `NULL` in which case,
#'   `narrow_wd = 0.5 * default_wd`, `wide_wd = 1.5 * default_wd`, and
#'   `xwide_wd = 2 * default_wd`
#'
#' @returns a modified `openxlsx` `Workbook` object
#' @export
#'
write_data_table_to_sheet <- function(wb,
                                      sheet_name,
                                      df,
                                      start_row,

                                      col_labels = NULL,
                                      col_label_row_ht = 60,

                                      withFilter = TRUE,
                                      freeze_header = TRUE,
                                      last_frozen_col = NULL,

                                      num_cols = NULL,
                                      perc_cols = NULL,
                                      dateFormat = "ddmmmyyyy",
                                      dateTimeFormat = "ddmmmyyyy hh:mm:ss",

                                      border = "TopBottomLeftRight",
                                      wrapText = TRUE,
                                      halign = "center",
                                      valign = "center",
                                      fontSize = 11,
                                      header_textDecoration = "bold",
                                      header_fgFill = "#b3d9e5",

                                      default_wd = 20,
                                      narrow_cols = NULL,
                                      wide_cols = NULL,
                                      xwide_cols = NULL,
                                      narrow_wd = NULL,
                                      wide_wd = NULL,
                                      xwide_wd = NULL) {

  headerStyle <- openxlsx::createStyle(
    border = border,
    wrapText = wrapText,
    halign = halign,
    valign = valign,
    fontSize = fontSize,
    textDecoration = header_textDecoration,
    fgFill = header_fgFill
  )

  # optionally, add a row of column labels above the main table
  if (!is.null(col_labels)) {

    # check all column names are present
    missing_labels <- setdiff(colnames(df), names(col_labels))
    if (length(missing_labels) > 0) {
      stop("There are column names in df without a label in col_labels")
    }

    # ensure the label list is in the same order as the column names
    col_labels <- col_labels[colnames(df)]

    # write the column labels row as the 1st row of the table
    purrr::iwalk(col_labels,
                 ~ openxlsx::writeData(wb,
                                       sheet = sheet_name,
                                       .x,
                                       startRow = start_row,
                                       startCol = which(colnames(df) == .y)))
    openxlsx::addStyle(wb,
                       sheet = sheet_name,
                       style = headerStyle,
                       rows = start_row,
                       cols = 1:ncol(df))

    # restrict height of label row
    openxlsx::setRowHeights(wb,
                            sheet = sheet_name,
                            rows = start_row,
                            heights = col_label_row_ht)

    # set the start row of the data table header
    table_start_row <- start_row + 1

  } else {
    table_start_row <- start_row
  }

  # add the data with header and filters
  openxlsx::writeData(wb,
                      sheet = sheet_name,
                      x = df,
                      colNames = TRUE,
                      startRow = table_start_row,
                      headerStyle = headerStyle,
                      withFilter = withFilter)

  # format the table body, if there is data
  if (nrow(df) > 0){

    # apply default style to all rows and columns
    contentStyleGeneral <- openxlsx::createStyle(
      border = border,
      wrapText = wrapText,
      halign = halign,
      valign = valign,
      fontSize = fontSize
    )
    openxlsx::addStyle(wb,
                       sheet = sheet_name,
                       style = contentStyleGeneral,
                       rows = table_start_row + seq(1, nrow(df)),
                       cols = seq(1, ncol(df)),
                       gridExpand = TRUE)

    # apply numeric style with user specified rounding by column name
    if (!is.null(num_cols)) {
      purrr::iwalk(num_cols, {
        round_format <- paste0("0.", rep("0", .x))
        contentStyleNumRnd <- openxlsx::createStyle(
          border = border,
          wrapText = wrapText,
          halign = halign,
          valign = valign,
          fontSize = fontSize,
          numFmt = round_format
        )
        openxlsx::addStyle(wb,
                           sheet = sheet_name,
                           style = contentStyleNumRnd,
                           rows = table_start_row + seq(1, nrow(df)),
                           cols = which(colnames(df) %in% .y),
                           gridExpand = TRUE)
      })
    }

    # apply percent style with user specified rounding by column name
    if (!is.null(perc_cols)) {
      purrr::iwalk(perc_cols, {
        round_format <- paste0("0.", rep("0", .x), "%")
        contentStylePercRnd <- openxlsx::createStyle(
          border = border,
          wrapText = wrapText,
          halign = halign,
          valign = valign,
          fontSize = fontSize,
          numFmt = round_format
        )
        openxlsx::addStyle(wb,
                           sheet = sheet_name,
                           style = contentStylePercRnd,
                           rows = table_start_row + seq(1, nrow(df)),
                           cols = which(colnames(df) %in% .y),
                           gridExpand = TRUE)
      })
    }

    ## add date and time styles
    col_classes <- purrr::modify(colnames(df), ~ class(df[[.]])[1])

    # dates
    contentStyleDates <- openxlsx::createStyle(
      border = border,
      wrapText = wrapText,
      halign = halign,
      valign = valign,
      fontSize = fontSize,
      numFmt = dateFormat
    )
    openxlsx::addStyle(wb,
                       sheet = sheet_name,
                       style = contentStyleDates,
                       rows = table_start_row + seq(1, nrow(df)),
                       cols = which(col_classes == "Date"),
                       gridExpand = TRUE)

    # date/times
    contentStyleDateTimes <- openxlsx::createStyle(
      border = border,
      wrapText = wrapText,
      halign = halign,
      valign = valign,
      fontSize = fontSize,
      numFmt = dateTimeFormat
    )
    openxlsx::addStyle(
      wb,
      sheet = sheet_name,
      style = contentStyleDateTimes,
      rows = table_start_row + seq(1, nrow(df)),
      cols = which(col_classes %in% c("POSIXlt", "POSIXt", "POSIXct")),
      gridExpand = TRUE
    )
  }

  # freeze column names and user specified left most columns
  if (!is.null(last_frozen_col) & freeze_header) {
    openxlsx::freezePane(
      wb,
      sheet = sheet_name,
      firstActiveRow = table_start_row + 1,
      firstActiveCol = which(colnames(df) == last_frozen_col) + 1
    )

    # freeze left panes only
  } else if (!is.null(last_frozen_col) & !freeze_header) {
    openxlsx::freezePane(
      wb,
      sheet = sheet_name,
      firstActiveCol = which(colnames(df) == last_frozen_col) + 1
    )

    # freeze header only
  } else if (is.null(last_frozen_col) & freeze_header) {
    openxlsx::freezePane(
      wb,
      sheet = sheet_name,
      firstActiveRow = table_start_row + 1
    )
  }

  # set default columns widths
  openxlsx::setColWidths(wb,
                         sheet = sheet_name,
                         cols = 1:ncol(df),
                         widths = default_wd)

  ## column widths
  # calculate and check column widths
  if (is.null(narrow_wd)) narrow_wd <- default_wd * 0.5
  if (narrow_wd >= default_wd) warning("narrow_wd >= default_wd")

  if (is.null(wide_wd)) wide_wd <- default_wd * 1.5
  if (wide_wd <= default_wd) warning("wide_wd <= default_wd")

  if (is.null(xwide_wd)) xwide_wd <- default_wd * 2
  if (xwide_wd <= wide_wd) warning("xwide_wd <= wide_wd")

  # change user specified column widths from default
  purrr::walk2(list(narrow_cols, wide_cols, xwide_cols),
               list(narrow_wd  , wide_wd  , xwide_wd  ),
               ~ {
                   if (!is.null(.x)) {
                     openxlsx::setColWidths(wb,
                                            sheet = sheet_name,
                                            cols = which(colnames(df) %in% .x),
                                            widths = .y)
                   }
                 }
  )

  return(wb)
}
