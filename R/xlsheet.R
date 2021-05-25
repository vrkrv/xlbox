#' Assign xlsheet class
#' @description Apply xlsheet function to the data list / container for export to the same sheet.
#' If input list is named then list names will be inserted on the row before the table (table title).
#' @param data_list List of data
#' @examples
#' library(magrittr)
#' xl_list <- list(mtcars = mtcars, iris = iris)
#' xl_list %>%
#'   xlsheet %>%
#'   write_xlsx
#' @export
xlsheet <- function(data_list) {
  if(!is_bare_list(data_list)) {
    stop("data_list must be bare list of objects!", call. = TRUE)
  }
  class(data_list) <- c(class(data_list), "xlsheet")
  return(data_list)
}

#' Get sheet's last row and column for the data which is already placed on this sheet.
#' @description This function can be used in case you need add smth to the current workbook.
#' @param workbook Workbook object
#' @param sheet Sheet name or index
#' @export
sheet_dim <- function(workbook, sheet = 1) {
  sheet <- if(is.character(sheet)) which(sheet == workbook$sheet_names) else sheet
  if (length(sheet) != 1) {
    warning("Sheet does not exist! Execution stops.")
    return(NULL)
  }
  nrow <- workbook$worksheets[[sheet]]$sheet_data$rows %>% max
  ncol <- workbook$worksheets[[sheet]]$sheet_data$cols %>% max
  return(c(nrow = nrow, ncol = ncol))
}

#' @describeIn  write_data Write xlsheet -- list of objects to be written in single sheet, block by block
#' @param data xlsheet data list / container to be exported to the same sheet
#' @param rowwise By default blocks placed in a column, one under another.
#' If you want to place datasets side by side, e.g. in a row, set to TRUE.
#' @param space_between Space between xlsheet elements (empty rows). Both scalar and vector are allowed.
#' @param use_names Use list names as table title before the table
write_data.xlsheet <- function(wb, sheet, data, startRow = 1, startCol = 1, rowwise = FALSE, ..., space_between = 1, use_names = TRUE) {
  start_row = startRow
  start_col = startCol
  titles = names(data)

  # change use_names to FALSE if unnamed list provided:
  if (is.null(titles)) {
    use_names = FALSE
  }

  for (i in seq_along(data)) {
    if (use_names && nchar(titles[i])) {
      openxlsx::writeData(wb, sheet, titles[i], startRow = start_row, startCol = start_col)
      openxlsx::addStyle(wb, sheet, getOption("xlbox.tableTitleStyle"), rows = start_row, cols = start_col)
      title_space <- 1
    } else {
      title_space <- 0
    }
    write_data(wb, sheet, data[[i]], startRow = start_row + title_space, startCol = start_col, ...)
    if (rowwise) {
      start_col <- sheet_dim(wb, sheet)[2] + startCol + space_between
    } else {
      start_row <- sheet_dim(wb, sheet)[1] + space_between + 1
    }
  }
}
