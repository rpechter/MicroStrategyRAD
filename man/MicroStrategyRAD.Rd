\name{MicroStrategyRAD}
\alias{MicroStrategyRAD}
\title{R Analytic Deployer}
\usage{
  MicroStrategyRAD(envRAD)
}
\arguments{
  \item{envRAD}{R environment that contains all information
  related to the deployment of an R Analytic}
}
\description{
  \code{MicroStrategyRAD} launches the MicroStrategy R
  Analytic Deployer to prepare an R Script for execution by
  MicroStrategy.
}
\details{
  This function launches a user interface which allows the
  user to open an R Script, capture its "signature" (the
  nature of inputs and outputs along with other information
  about how the R analytic should be executed).
}
\examples{
MicroStrategyRAD(envRAD <- new.env())
}
\keyword{microstrategy}

