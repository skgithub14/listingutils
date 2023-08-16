#' Set-up Excel workbook with standard options
#'
#' Creates a `openxlsx::createWorkbook()` object, adds a worksheet, sets column
#' widths to auto and configures options for borders, dates and maximum column
#' widths.
#'
#' No data is written to the workbook. The `dat` argument is just for configuring
#' options for the correct number of columns.
#'
#' @param sheet_name string, the sheet name to create for `dat`
#' @param dat data frame of the data for the sheet
#'
#' @returns a `openxlsx::createWorkbook()` workbook object
setup_workbook <- function (sheet_name, dat) {
  wb <- createWorkbook()
  addWorksheet(wb, sheetName = sheet_name)
  setColWidths(wb, sheet = sheet_name, cols = 1:ncol(dat), widths = 15)
  return(wb)
}


#' Writes header data to Excel workbook
#'
#' Writes the meta data to the top rows of an Excel workbook and formats them
#'
#' @param wb an `openxlsx::createWorkbook()` object
#' @param sheet_name a string, the sheet name in `wb` to write to
#' @param h_data a list where each element corresponds to one row in the header
#' @param dat a data frame of the data that will be written below the header
#'
#' @returns a modified `openxlsx::createWorkbook()` object
add_worksheet_header <- function (wb, sheet_name, h_data, dat) {
  for (hd in seq_along(h_data)) {
    writeData(wb, sheet = sheet_name, x = h_data[[hd]], startRow = hd)
    mergeCells(wb, sheet = sheet_name, rows = hd, cols = 1:ncol(dat))
    if (hd == 1) {
      study_style <- createStyle(fontSize = 15, textDecoration = "bold")
      addStyle(wb,
               sheet = sheet_name,
               study_style,
               rows = 1,
               cols = 1:ncol(dat),
               gridExpand = TRUE)
    }
  }

  return(wb)
}


#' Writes a data frame to Excel workbook with formatting
#'
#' Writes a data frame to the specified sheet in an Excel workbook and applies
#' formatting, conditional formatting, freeze panes, and filters.
#'
#' @param wb an `openxlsx::createWorkbook()` object
#' @param sheet_name string, the sheet name to write to
#' @param dat data frame, the data to write
#' @param start_row integer, the row number at which the write the column header
#' of `dat`
#' @param disrepancy_col a string, the column name for discrepancy change
#' tracking between report versions, default is `"Discrepancy Change"`
#' @param match_text a string, text that identifies data match in
#' `disrepancy_col`, default is `"Data match"`
#' @param finding_col a string, the column name for reconciliation findings,
#' default is `"Finding"`
#'
#' @returns a modified `openxlsx::createWorkbook()` object
add_worksheet_data <- function (wb,
                                sheet_name,
                                dat,
                                start_row,
                                disrepancy_col = "Discrepancy Change",
                                match_text = "Data match",
                                finding_col = "Finding") {

  # set table header style
  headerStyle <- createStyle(border = "TopBottomLeftRight",
                             wrapText = TRUE,
                             halign = "center",
                             valign = "center",
                             textDecoration = "bold",
                             fgFill = "#b3d9e5")

  # add the data
  writeData(wb,
            sheet = sheet_name,
            x = dat,
            colNames = TRUE,
            startRow = start_row,
            borders = "all",
            headerStyle = headerStyle,
            withFilter = T)

  ## table body formatting
  body_start <- start_row + 1
  body_end <- start_row + nrow(dat)
  body_rows <- seq(body_start, body_end)

  freezePane(wb, sheet = sheet_name, firstActiveRow = body_start)

  # add styles depending on if a column has a Date class or not
  col_classes <- modify(colnames(dat), ~ class(dat[[.]]))
  contentStyleGeneral <- createStyle(border = "TopBottomLeftRight",
                                     wrapText = TRUE,
                                     halign = "center",
                                     valign = "center")
  contentStyleDates <- createStyle(border = "TopBottomLeftRight",
                                   wrapText = TRUE,
                                   halign = "center",
                                   valign = "center",
                                   numFmt = "ddmmmyyyy")
  addStyle(wb,
           sheet = sheet_name,
           style = contentStyleGeneral,
           rows = body_rows,
           cols = which(col_classes != "Date"),
           gridExpand = TRUE)
  addStyle(wb,
           sheet = sheet_name,
           style = contentStyleDates,
           rows = body_rows,
           cols = which(col_classes == "Date"),
           gridExpand = TRUE)

  ## conditional formatting
  # mark non-Data Matches in the Discrepancy column
  FindingNotMatchStyle <- createStyle(fontColour = "#FF0000",
                                      bgFill = "#FFFF00")
  conditionalFormatting(wb,
                        sheet = sheet_name,
                        rows = body_rows,
                        cols = which(colnames(dat) == finding_col),
                        rule = paste0("!=\"", match_text,"\""),
                        type = "expression",
                        style = FindingNotMatchStyle)

  # mark "Changes" in the Change column
  ChangeChangeStyle <- createStyle(fontColour = "#FF0000", bgFill = "#FFFF00")
  conditionalFormatting(wb,
                        sheet = sheet_name,
                        rows = body_rows,
                        cols = which(colnames(dat) == disrepancy_col),
                        rule = "Change",
                        type = "contains",
                        style = ChangeChangeStyle)

  # mark "New" entries in the Change column
  ChangeNewStyle <- createStyle(fontColour = "#00FF00", bgFill = "#FFFF00")
  conditionalFormatting(wb,
                        sheet = sheet_name,
                        rows = body_rows,
                        cols = which(colnames(dat) == disrepancy_col),
                        rule = "New",
                        type = "contains",
                        style = ChangeNewStyle)

  return(wb)
}
