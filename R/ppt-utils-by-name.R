ph_with_vg_by_name_template <- function(x, name, template_ppt, code, ggobj = NULL, ...) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_vg_at(x, code, ggobj = ggobj, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...)
}

ph_with_img_by_name_template <- function(x, name, template_ppt, src, ...) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_img_at(x, src = src, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...)
}

ph_with_gg_by_name_template <- function(x, name, template_ppt, object, ...) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_gg(x, object, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...)
}

ph_with_vg_by_name <- function(x, name, code, ggobj = NULL, ...) {
    loc <- slide_summary(x) %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    ph_with_vg_at(x, code, ggobj = ggobj, left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy, ...) %>%
        ph_remove(ph_label = name)
}


ph_with_text_by_name_template <- function(x, text, name, template_ppt) {
    loc <- template_ppt %>%
        slide_summary() %>%
        filter(ph_label == name) %>%
        slice(1)
    print(loc)
    x %>% ph_with_fpars_at(fpars = fp_text_arial(text), left = loc$offx, top = loc$offy, width = loc$cx, height = loc$cy)
}

fp_text_arial <- function(text, ...) {
    list(fpar(ftext(glue(text), fp_text(font.size = 6, font.family = "Arial", ...))))
}
