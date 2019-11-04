#' A ggplot2 custom theme (based on 'ArialMT' font)
#'
#' @param size base size of text
#'
#' @export
theme_arial <- function(size = 8) {
    theme_arial <- ggplot2::theme_light() %+replace%
        theme(
            panel.background = element_blank(),
            panel.grid.minor = element_blank(),
            strip.background = element_blank(),
            strip.text = element_text(colour = "black", size = size, family = "ArialMT")
        )
    theme_arial <- theme_arial + theme(
        text = element_text(size = size, family = "ArialMT"),
        axis.text = element_text(size = size, family = "ArialMT"),
        legend.text = element_text(size = size, family = "ArialMT"),
        plot.title = element_text(size = size, family = "ArialMT")
    )

    return(theme_arial)
}

#' Set x-axis vertical labels in ggplot2
#'
#' @export
vertical_labs <- function() {
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}

#' Raster version of geom_point
#'
#' @param ... arguments passed to ggrastr::geom_point_rast
#'  
#' @seealso \code{\link[ggrastr]{geom_point_rast}}
#' @export
geom_point_rast <- purrr::partial(ggrastr::geom_point_rast, raster.width = 1, raster.height = 1)

#' Raster version of geom_boxplot with jittering of outliers
#' 
#' @param ... arguments passed to ggrastr::geom_boxplot_jitter
#' 
#' @seealso \code{\link[ggrastr]{geom_boxplot_jitter}}
#' 
#' @export
geom_boxplot_jitter <- purrr::partial(ggrastr::geom_boxplot_jitter, raster.width = 1, raster.height = 1)
