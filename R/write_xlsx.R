#' Write list of objects to xlsx
#' @description Simple tweak of write.xlsx for lists of data objects compatible with write_data classes.
#' @param data Data Frame or list of data frames. List elements must have unique names! If x is not bare list, then it will be converted to one.
#' @param file Path to output file. If file not specified (NULL),  open preview.
#' @param withFilter If \code{TRUE}, add filters to column name row.
#' @param headerStyle By default brand dark orange gfk's color will be used for headers.
#' @param ... Other options to write_data
#' @seealso \code{\link{write_data}}
#' @examples
#' library(xlbox)
#' library(magrittr)
#'
#' # single data frame export:
#' mtcars %>% write_xlsx
#'
#' # export list of data frames to multiple sheets:
#' df_list <- list(mtcars = mtcars %>% head(), iris = iris %>% head())
#' df_list %>% write_xlsx
#'
#' # export list of data frames to single sheet one under another (in a column):
#' df_list %>% xlsheet %>% write_xlsx
#'
#' # export list of data frames to single sheet one under another (in a row):
#' df_list %>% xlsheet %>% write_xlsx(rowwise = TRUE)
#'
#' ## Example of custom S3 report for lm:
#' library(openxlsx)
#' library(broom)
#' write_data.lm <- function(wb, sheet, data, startRow = 1, startCol = 1, ...) {
#'   fit_info <- glance(fit)
#'   coef_df <- tidy(fit)
#'
#'   srow = startRow
#'   scol = startCol
#'
#'   # place header on the first row and merge first line:
#'   writeData(wb, sheet, "Fit Information", startRow = srow, startCol = scol, ...)
#'   mergeCells(wb, sheet, cols = scol + 1:ncol(fit_info) - 1, rows = srow)
#'   srow <- srow + 1
#'
#'   # place fit_info:
#'   writeData(wb, sheet, fit_info, startRow = srow, startCol = scol, ...)
#'   srow <- srow + nrow(fit_info) + 2
#'
#'   # the same for coefficients:
#'   writeData(wb, sheet, "Coefficients", startRow = srow, startCol = scol, ...)
#'   mergeCells(wb, sheet, cols = scol + 1:ncol(coef_df) - 1, rows = srow)
#'   srow <- srow + 1
#'   # apply accounting style for values:
#'   class(coef_df$p.value) <- "accounting"
#'   writeData(wb, sheet, coef_df, startRow = srow, startCol = scol, ...)
#' }
#'
#' # test how lm will export to xlsx:
#' fit <- lm(Petal.Width ~ . - Species, data = iris)
#' fit %>% write_xlsx
#' @export
#' @seealso \code{\link{xlsheet}}, \code{\link{write_data}}
write_xlsx <- function(data,
                       file = NULL,
                       withFilter = getOption("xlbox.withFilter"),
                       headerStyle = getOption("xlbox.tableHeaderStyle"),
                       ...) {
  # if not bare list convert to one:
  if(!is_bare_list(data)) {
    data = list(Sheet1 = data)
  }

  wb <- openxlsx::createWorkbook()
  # TODO Deal with duplicated names!
  for (d in seq_along(data)) {
    sheetName <- ifelse(is.null(names(data)[d]), sprintf("Sheet%s", d), names(data)[d])
    addWorksheet(wb, sheetName = sheetName)
    write_data(wb, sheetName, data = data[[d]], withFilter = withFilter, headerStyle = headerStyle, ...)
  }

  if (is.null(file)) {
    openxlsx::openXL(wb)
  } else {
    openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  }
}
