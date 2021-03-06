% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/xlsheet.R
\name{xlsheet}
\alias{xlsheet}
\title{Assign xlsheet class}
\usage{
xlsheet(data_list)
}
\arguments{
\item{data_list}{List of data}
}
\description{
Apply xlsheet function to the data list / container for export to the same sheet.
If input list is named then list names will be inserted on the row before the table (table title).
}
\examples{
library(magrittr)
xl_list <- list(mtcars = mtcars, iris = iris)
xl_list \%>\%
  xlsheet \%>\%
  write_xlsx
}
