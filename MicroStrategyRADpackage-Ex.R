pkgname <- "MicroStrategyRADpackage"
source(file.path(R.home("share"), "R", "examples-header.R"))
options(warn = 1)
options(pager = "console")
library('MicroStrategyRADpackage')

assign(".oldSearch", search(), pos = 'CheckExEnv')
cleanEx()
nameEx("MicroStrategyRAD")
### * MicroStrategyRAD

flush(stderr()); flush(stdout())

### Name: MicroStrategyRAD
### Title: R Analytic Deployer
### Aliases: MicroStrategyRAD
### Keywords: microstrategy

### ** Examples

MicroStrategyRAD(envRAD <- new.env())



### * <FOOTER>
###
cat("Time elapsed: ", proc.time() - get("ptime", pos = 'CheckExEnv'),"\n")
grDevices::dev.off()
###
### Local variables: ***
### mode: outline-minor ***
### outline-regexp: "\\(> \\)?### [*]+" ***
### End: ***
quit('no')
