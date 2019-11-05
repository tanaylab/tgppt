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
#' @inherit ggrastr::geom_point_rast 
#' @examples 
#' library(ggplot2)
#' ggplot() + geom_point_rast(aes(x=rnorm(1000), y=rnorm(1000)), raster.dpi=600)
#' 
#' @export
geom_point_rast <- ggrastr::geom_point_rast

#' Raster version of geom_boxplot with jittering of outliers
#' 
#' @inherit ggrastr::geom_boxplot_jitter
#' @examples 
#' library(ggplot2) 
#' ggplot() + geom_boxplot_jitter(aes(y=rt(1000, df=3), x=as.factor(1:1000 %% 2)), outlier.jitter.width = 0.1, raster = TRUE)
#' 
#' @export
geom_boxplot_jitter <- ggrastr::geom_boxplot_jitter
