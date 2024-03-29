% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/ppt-utils.R
\name{new_ppt}
\alias{new_ppt}
\title{Generate an empty pptx file}
\usage{
new_ppt(fn)
}
\arguments{
\item{fn}{file name of the new pptx file}
}
\value{
logical value indicating if the operation succeeded
}
\description{
Generate an empty pptx file
}
\examples{
ppt_file <- tempfile(fileext = ".pptx")
new_ppt(ppt_file)
}
