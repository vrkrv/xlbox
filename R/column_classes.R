# openxlsx classes --------------------------------------------------------

#' factory function to assign classes
column_class <- function(class_name) {
  function(x) {
    class(x) <- c(class(x), class_name)
    return(x)
  }
}

#' Assign openxlsx format to column
#' @description Explicit openxlsx columnar class assignment.
#' @param x Column
#' @describeIn as_percentage Assign percentage format to column
#' @examples
#' library(magrittr)
#' tab <-
#'  data.frame(
#'    type = c("a", "b"),
#'    abs = c(60, 40),
#'    share = c(.6, .4)
#'  )
#' tab$share <- as_percentage(tab$share)
#'
#' # Column have includes percentage class and will be proper exported to excel sheet:
#' class(tab$share)
#' @export
as_percentage <- column_class("percentage")

#' @describeIn as_percentage Assign currency format to column
#' @export
as_currency <- column_class("currency")

#' @describeIn as_percentage Assign hyperlink format to column
#' @export
as_hyperlink <- column_class("hyperlink")

#' Assign percentage format to column
#' @describeIn as_percentage Assign scientific format to column
#' @export
as_scientific <- column_class("scientific")
