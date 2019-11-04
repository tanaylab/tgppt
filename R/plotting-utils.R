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
        legend.text = element_text(size = size, family = "ArialMT")
    )

    return(theme_arial)
}

#' Set x-axis vertical labels in ggplot2
#'
#' @export
vertical_labs <- function() {
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}
