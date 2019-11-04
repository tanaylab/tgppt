.onLoad <- function(libname, pkgname) {
    utils::suppressForeignCheck(c("ph_label"))
    utils::globalVariables(c("ph_label"))
}