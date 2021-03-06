% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/write_data.R, R/xlsheet.R
\name{write_data}
\alias{write_data}
\alias{write_data.data.frame}
\alias{write_data.default}
\alias{write_data.table}
\alias{write_data.matrix}
\alias{write_data.xlsheet}
\title{Write S3-class object to the worksheet}
\usage{
write_data(
  wb,
  sheet,
  data,
  startCol = 1,
  colNames = TRUE,
  withFilter = TRUE,
  headerStyle = getOption("xlbox.tableHeaderStyle"),
  ...
)

\method{write_data}{data.frame}(wb, sheet, data, ..., dataTable = getOption("xlbox.dataTable"))

\method{write_data}{default}(
  wb,
  sheet,
  data,
  rowNames = TRUE,
  ...,
  dataTable = getOption("xlbox.dataTable")
)

\method{write_data}{table}(
  wb,
  sheet,
  data,
  rowNames = TRUE,
  ...,
  dataTable = getOption("xlbox.dataTable")
)

\method{write_data}{matrix}(wb, sheet, data, ..., dataTable = getOption("xlbox.dataTable"))

\method{write_data}{xlsheet}(
  wb,
  sheet,
  data,
  startRow = 1,
  startCol = 1,
  rowwise = FALSE,
  ...,
  space_between = 1,
  use_names = TRUE
)
}
\arguments{
\item{wb}{A Workbook object containing worksheet.}

\item{sheet}{Worksheet index / name write to.}

\item{data}{xlsheet data list / container to be exported to the same sheet}

\item{startCol}{Start column}

\item{colNames}{If \code{TRUE}, first row is treated as column names}

\item{withFilter}{Add autofilter to header. Can be applied only for single row.}

\item{headerStyle}{Custom style to apply to column names.}

\item{...}{Other parameters from openxlsx::writeData}

\item{dataTable}{Switch between \code{\link[openxlsx]{writeData}} and}

\item{rowNames}{If \code{TRUE}, row names will be also added to xlsx file.}

\item{startRow}{Start row}

\item{rowwise}{By default blocks placed in a column, one under another.
If you want to place datasets side by side, e.g. in a row, set to TRUE.}

\item{space_between}{Space between xlsheet elements (empty rows). Both scalar and vector are allowed.}

\item{use_names}{Use list names as table title before the table}

\item{base}{Simple description of subsample (for scripting purposes).}
}
\description{
Write S3-class object to the worksheet.
S3 class for \code{\link{openxlsx::writeData}} extended to raux analytics output.
}
\section{Methods (by class)}{
\itemize{
\item \code{data.frame}: Write data frame to sheet.

\item \code{default}: Write object to sheet.

\item \code{table}: Write table to sheet.

\item \code{matrix}: Write matrix to sheet.

\item \code{xlsheet}: Write xlsheet -- list of objects to be written in single sheet, block by block
}}

\seealso{
\code{\link[openxlsx]{writeData}} and \code{\link[openxlsx]{writeDataTable}} for additional help
\code{\link[openxlsx]{writeDataTable}}. By default calls \code{\link[openxlsx]{writeData}}.
}
