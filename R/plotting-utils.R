#' A ggplot2 custom theme (based on 'ArialMT' font)
#'
#' @param size base size of text
#'
#' @return a \code{\link[ggplot2]{theme}} object
#'
#' @examples
#' library(ggplot2)
#' ggplot(mtcars, aes(x = mpg, y = drat)) +
#'     geom_point() +
#'     theme_arial()
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
#' @return a \code{\link[ggplot2]{theme}} object that rotates x-axis labels to 90
#'   degrees
#'
#' @examples
#' library(ggplot2)
#' ggplot(mtcars, aes(x = factor(cyl), y = mpg)) +
#'     geom_boxplot() +
#'     vertical_labs()
#' @export
vertical_labs <- function() {
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}
