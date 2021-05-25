#' Write S3-class object to the worksheet
#' @description Write S3-class object to the worksheet.
#' S3 class for \code{\link{openxlsx::writeData}} extended to raux analytics output.
#' @param wb A Workbook object containing worksheet.
#' @param sheet Worksheet index / name write to.
#' @param data Data to be written
#' @param startCol Start column
#' @param startRow Start row
#' @param rowNames If \code{TRUE}, row names will be also added to xlsx file.
#' @param colNames If \code{TRUE}, first row is treated as column names
#' @param withFilter Add autofilter to header. Can be applied only for single row.
#' @param headerStyle Custom style to apply to column names.
#' @param dataTable Switch between \code{\link[openxlsx]{writeData}} and
#' @param base Simple description of subsample (for scripting purposes).
#' @param ... Other parameters from openxlsx::writeData
#' @seealso \code{\link[openxlsx]{writeData}} and \code{\link[openxlsx]{writeDataTable}} for additional help
#' \code{\link[openxlsx]{writeDataTable}}. By default calls \code{\link[openxlsx]{writeData}}.
#' @import openxlsx
#' @export
write_data <- function(wb,
                       sheet,
                       data,
                       startCol = 1,
                       colNames = TRUE,
                       withFilter = TRUE,
                       headerStyle = getOption("xlbox.tableHeaderStyle"),
                       ...) {
  UseMethod("write_data", data)
}

#' @describeIn write_data Write data frame to sheet.
write_data.data.frame <- function(wb, sheet, data, ..., dataTable = getOption("xlbox.dataTable")) {
  if(!dataTable) {
    openxlsx::writeData(wb, sheet, data, ...)
  } else {
    openxlsx::writeDataTable(wb, sheet, data, ...)
  }
}

#' @describeIn write_data Write object to sheet.
write_data.default <- function(wb, sheet, data, rowNames = TRUE, ..., dataTable = getOption("xlbox.dataTable")) {
  write_data.data.frame(wb, sheet, as.matrix(data), rowNames = TRUE, ..., dataTable = dataTable)
}

#' @describeIn write_data Write table to sheet.
write_data.table <- function(wb, sheet, data, rowNames = TRUE, ..., dataTable = getOption("xlbox.dataTable")) {
  write_data.data.frame(wb, sheet, as.matrix(data), rowNames = TRUE, ..., dataTable = dataTable)
}

#' @describeIn write_data Write matrix to sheet.
write_data.matrix <- function(wb, sheet, data, ..., dataTable = getOption("xlbox.dataTable")) {
  write_data.data.frame(wb, sheet, data, ..., dataTable = dataTable)
}
