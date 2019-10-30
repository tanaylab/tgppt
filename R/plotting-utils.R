#' A ggplot2 custom theme (based on 'ArialMT' font)
#' 
#' @param size base size of text
#' 
#' @export
arial_theme <- function(size = 8){
    arial_theme <- ggplot2::theme_light() %+replace%
        theme(
            panel.background = element_blank(),
            panel.grid.minor = element_blank(),
            strip.background = element_blank(),
            strip.text = element_text(colour = "black", size = size, family = "ArialMT")
        )
     arial_theme <- arial_theme + theme(
        text = element_text(size = size, family = "ArialMT"),
        axis.text = element_text(size = size, family = "ArialMT"),
        legend.text = element_text(size = size, family = "ArialMT")
    )

    return(arial_theme)
}

#' Set x-axis vertical labels in ggplot2
#' 
#' @export
vertical_labs <- function() {
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}