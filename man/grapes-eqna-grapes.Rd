% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/operators.R
\name{\%eqna\%}
\alias{\%eqna\%}
\title{Operator for equal comparison with NA handling}
\usage{
e1 \%eqna\% e2
}
\arguments{
\item{e1, e2}{vectors of equal length to compare that can be of any class}
}
\value{
a vector of length equal to `e1` and `e2` containing logical values
of `TRUE` or `FALSE` only (no `NA` values).
}
\description{
Provides element-wise comparison of two vectors to determine if the values
are equal. Comparison of two non-`NA` values is equivalent to the result
of `value1 == value2`.`NA` compared to a non-`NA` values will return `TRUE`
and `NA` compared to `NA` will return `TRUE`.
}
\examples{
# the below will return `c(TRUE, FALSE, TRUE, TRUE)`
c(1, 1, 1, NA) \%eqna\% c(1, 2, NA, NA)

}
