#' Make an Excel workbook
#'
#' Creates an [openxlsx::createWorkbook()] object with one or more sheets and
#' populates header information and data for each sheet.
#'
#' @inheritParams setup_multi_sheet_workbook
#' @inheritParams populate_workbook
#'
#' @returns an `openslxs` `Workbook` object
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
#' tracking between report versions, default is `"Discrepancy Change"`
#' @param match_text a string, text that identifies a data match in
#' `discrepancy_col`, default is `"Data match"`
#' @param finding_col a string, the column name for reconciliation findings,
#' default is `"Finding"`
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

  # mark "Changes" in the Change column
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



