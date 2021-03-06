% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/smtools.R
\name{summarise}
\alias{summarise}
\title{Summarise data columns.}
\usage{
summarise(dt, writepath = str_c(getwd(), "/Data Summary.xlsx"),
  append = "Data Summary")
}
\arguments{
\item{dt}{Data table to summarise.}

\item{writepath}{Path + file name to save xlsx; can be \code{NULL}.}

\item{append}{Whether the summary should be appended to the workbook; any 
strings except "Data Summary" mean append; otherwise the summary table is
written on the worksheet "Data Summary".}
}
\value{
Data summary.
}
\description{
Summarise data columns.
}
