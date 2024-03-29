% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/excel.R
\name{write_data_table_to_sheet}
\alias{write_data_table_to_sheet}
\title{Write a data frame or tibble to Excel worksheet with features and formatting}
\usage{
write_data_table_to_sheet(
  wb,
  sheet_name,
  df,
  start_row,
  start_col = 1,
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
  col_label_textDecoration = "bold",
  col_label_fgFill = "#b3d9e5",
  col_label_fontColour = "black",
  header_textDecoration = "bold",
  header_fgFill = "#b3d9e5",
  header_fontColour = "black",
  default_wd = 20,
  narrow_cols = NULL,
  wide_cols = NULL,
  xwide_cols = NULL,
  narrow_wd = NULL,
  wide_wd = NULL,
  xwide_wd = NULL
)
}
\arguments{
\item{wb}{the `openxlsx` `Workbook` object to modify}

\item{sheet_name}{a string, the sheet name in `wb` to modify}

\item{df}{the data frame or tibble to write}

\item{start_row}{a numeric value, the row to start writing the table; if
`col_labels` is not `NULL` then the column labels will be written to this
row with the column names the row below them, otherwise the column names
will be written here}

\item{start_col}{a numeric value, the column to start writing the table;
default is `1`}

\item{col_labels}{an optional named character vector of column labels where
names are the column names in `df`; if present, column labels are added to
the row directly above the column titles}

\item{col_label_row_ht}{an optional numeric value, if `col_labels` is not
`NULL`, sets the height of the column label row; default is `60`}

\item{withFilter}{a logical value indicating if column filtering should be
turned on; default is `TRUE`}

\item{freeze_header}{a logical value indicating whether the header row should
be frozen; default is `TRUE`}

\item{last_frozen_col}{an optional string, the column name in `df` which
should be the last frozen column; default is `NULL` in which case no
columns will be frozen}

\item{num_cols}{an optional named numeric vector where the names are column
names in `df` that are numeric columns and the values are number of decimal
places each column should be rounded to; default is `NULL`, in which case
any numeric columns not listed in `num_cols` or `perc_cols` will use
Excel's general cell format}

\item{perc_cols}{an optional named numeric vector where the names are column
names in `df` that are percentage columns and the values are number of
decimal places each column should be rounded to; default is `NULL`, in
which case any numeric columns not specified in `num_cols` or `perc_cols`
will use Excel's general cell format. Note, any column specified in this
argument will be multiplied by 100.}

\item{dateFormat}{a string specifying the format of Date class columns using
the `openxlsx` specification; default is `"ddmmmyyyy"`}

\item{dateTimeFormat}{a string specifying the format of the date/time class
columns (`c("POSIXlt", "POSIXt", "POSIXct")`) using the `openxlsx`
specification; default is `"ddmmmyyyy hh:mm:ss"`}

\item{border}{a string specifying the cell borders for all cells using the
`openxlsx` border specification; default is `"TopBottomLeftRight"`}

\item{wrapText}{a logical value indicating if all cells should wrap text;
default is `TRUE`}

\item{halign, valign}{string values indicating the horizontal and vertical
cell alignment for all cells using the `openxlsx` specification; default is
`center`}

\item{fontSize}{a numeric value specifying the font size for all cells;
default is `11`}

\item{col_label_textDecoration}{a string describing the text styling for the
column label row using the `openxlsx` specification; default is `"bold"`}

\item{col_label_fgFill}{a string containing a color hex code which will be
used to color the column label rows; default is `"#b3d9e5"` (light blue)}

\item{col_label_fontColour}{a string specifying the font color for the column
label row, default is `"black"`}

\item{header_textDecoration}{a string describing the text styling for the
column name row using the `openxlsx` specification; default is `"bold"`}

\item{header_fgFill}{a string containing a color hex code which will be used
to color the column name row; default is `"#b3d9e5"` (light blue)}

\item{header_fontColour}{a string specifying the font color for the column
header row, default is `"black"`}

\item{default_wd}{a numeric value specifying the default column width;
default is `20`; note if `narrow_wd`, `wide_wd`, or `xwide_wd` are `NULL`,
they will be adjusted relative to `default_wd` (see their parameter details
for more information)}

\item{narrow_cols, wide_cols, xwide_cols}{optional character vectors of column
names which will be made narrower, wider or extra wide; default is `NULL`}

\item{narrow_wd, wide_wd, xwide_wd}{optional numeric values which specify the
column widths for the columns specified in the `narrow_cols`, `wide_cols`,
and `xwide_cols` arguments, respectively; default is `NULL` in which case,
`narrow_wd = 0.5 * default_wd`, `wide_wd = 1.5 * default_wd`, and
`xwide_wd = 2 * default_wd`}
}
\value{
a modified `openxlsx` `Workbook` object
}
\description{
Writes a data frame/tibble to an Excel worksheet with the
  following features:
 - Column labels/descriptions above the header
 - Filters
 - Freeze panes
 - Rounding
 - Conversion to percentage
 - Date and date/time formatting
 - Header and cell styling (alignment, wrap, borders)
 - Column widths (default, narrow, wide, and extra wide widths)
}
