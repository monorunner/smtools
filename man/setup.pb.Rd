% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/smtools.R
\name{setup.pb}
\alias{setup.pb}
\title{Progress bar for loops - no need to set up every time.}
\usage{
setup.pb(total)
}
\arguments{
\item{total}{Size of loop.}

\item{ind}{Index in the loop.}

\item{startind}{Starting index of the loop.}
}
\description{
Progress bar for loops - no need to set up every time.
}
\details{
No error catching, especially around interrupted loops and 
start index.
}
