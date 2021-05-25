.onLoad <- function(libname, pkgname) {
  op.xlbox <- list(
    xlbox.tableHeaderStyle = openxlsx::createStyle(fontColour =  "#FFFFFF", fgFill = "#E55A00"),
    xlbox.tableTitleStyle  = openxlsx::createStyle(textDecoration = c("bold", "underline2")),
    xlbox.withFilter = TRUE,
    xlbox.dataTable = FALSE
  )
  options(op.xlbox)
  invisible()
}
