#' A ggplot2 custom theme (based on 'Arial' font)
#' 
#' @param size base size of text
#' 
#' @export
arial_theme <- function(size = 6){
    arial_theme <- ggplot2::theme_light() %+replace%
        theme(
            panel.background = element_blank(),
            panel.grid.minor = element_blank(),
            strip.background = element_blank(),
            strip.text = element_text(colour = "black", size = 6, family = "Arial")
        )
     arial_theme <- arial_theme + theme(
        text = element_text(size = 6, family = "Arial"),
        axis.text = element_text(size = 6, family = "Arial")
    )

    return(arial_theme)
}

#' Set X axis vertical labels in ggplot2
#' 
#' @export
vertical_labs <- function() {
    theme(axis.text.x = element_text(angle = 90, hjust = 1))
}