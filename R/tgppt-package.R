#' @keywords internal
#' @import dplyr
#' @import tidyr
#' @import ggplot2
#' @importFrom glue glue
#' @import purrr
#' @import stringr
#' @import rvg
#' @import officer
#' @importFrom ggrastr geom_point_rast
#' @importFrom ggrastr geom_boxplot_jitter
"_PACKAGE"

#' @export
#' @inherit ggrastr::geom_point_rast
geom_point_rast <- purrr::partial(ggrastr::geom_point_rast, raster.width = 1, raster.height = 1)

#' @export
#' @inherit ggrastr::geom_boxplot_jitter
geom_boxplot_jitter <- purrr::partial(ggrastr::geom_boxplot_jitter, raster.width = 1, raster.height = 1)

# The following block is used by usethis to automatically manage
# roxygen namespace tags. Modify with care!
## usethis namespace: start
## usethis namespace: end
NULL
