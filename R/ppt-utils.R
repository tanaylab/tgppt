#' @export
fp_text_arial <- function(text, ...) {
    list(fpar(ftext(glue(text), fp_text(font.size = 6, font.family = "Arial", ...))))
}

#' @export
ph_with_vg_by_name_template <- function(x, name, template_ppt, code, ggobj = NULL, ...) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_vg_at(x, code, ggobj = ggobj, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...)
}

#' @export
ph_with_img_by_name_template <- function(x, name, template_ppt, src, ...) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_img_at(x, src = src, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...)
}

#' @export
ph_with_gg_by_name_template <- function(x, name, template_ppt, object, ...) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_gg(x, object, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...)
}

#' @export
ph_with_vg_by_name <- function(x, name, code, ggobj = NULL, ...) {
    loc <- slide_summary(x) %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_vg_at(x, code, ggobj = ggobj, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...) %>%
        ph_remove(ph_label = name)
}


#' @export
ph_with_text_by_name_template <- function(x, text, name, template_ppt) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    x %>% ph_with_fpars_at(fpars = fp_text_arial(text), left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy)
}

#' Generate an empty pptx file
#' 
#' @param fn file name of the new pptx file
#' 
#' @return logical value indicating if the operation succeded
#' 
#' @examples 
#' new_ppt()
new_ppt <- function(fn){
    file.copy(system.file("ppt", "template.pptx", package = "alutil"), fn)
}

#' Plot ggplot object to ppt
#' 
#' @param gg ggplot object
#' @param out_ppt output pptx file
#' @param height height in cm
#' @param width width in cm
#' @param left left alignment in cm
#' @param top top alignment in cm
#' @param inches measures are given in inches
#' @param sep_legend plot legend and plot separatly
#' @param new_slide add plot to a new slide
#' 
#' @examples
#' 
#' gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
#' temp_ppt <- tempfile(fileext = ".pptx")
#' plot_gg_ppt(gg, temp_ppt)
#' 
#' @export
plot_gg_ppt <- function(gg, out_ppt, height = 6, width = 6, left = 5, top = 5, inches = FALSE, sep_legend = FALSE, new_slide = FALSE){
    if (sep_legend){
        plot_base_ppt(code = print(gg + theme(legend.position="none")), out_ppt = out_ppt, height = height, width = width, left = left, top = top, inches = inches, new_slide = new_slide)
        plot_base_ppt(code = print(grid::grid.draw(cowplot::get_legend(gg))), out_ppt = out_ppt, height = height, width = width, left = left, top = top, inches = inches, new_slide = new_slide)
    } else {
        plot_base_ppt(code = print(gg), out_ppt = out_ppt, height = height, width = width, left = left, top = top, inches = inches, new_slide = new_slide)    
    }
    
}

#' Plot base R plots to ppt
#' 
#' @param code code that generates plot
#' @param out_ppt output pptx file
#' @param height height in cm
#' @param width width in cm
#' @param left left alignment in cm
#' @param top top alignment in cm
#' @param inches measures are given in inches
#' @param new_slide add plot to a new slide
#' 
#' @examples
#' 
#' temp_ppt <- tempfile(fileext = ".pptx")
#' plot_base_ppt({plot(mtcars$mpg, mtcars$drat)}, temp_ppt)
#' 
#' 
#' @export
plot_base_ppt <- function(code, out_ppt, height = 6, width = 6, left = 5, top = 5, inches = FALSE, new_slide = FALSE){

    cm2inch <- 1
    if (!inches){
        cm2inch <- 2.54
    }

    if (!file.exists(out_ppt)){
        ppt <- read_pptx(system.file("ppt", "template.pptx", package = "alutil"))
    } else {
        ppt <- read_pptx(out_ppt)
        if (new_slide){
            ppt <- ppt %>% add_slide(layout = layout_summary(ppt)$layout[1], master = layout_summary(ppt)$master[1])
        }
    }    
    
    ppt <- ppt %>% ph_with_vg_at(code = code, height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch)
    print(ppt, target = out_ppt)
}