#' Add a slide title text box to a pptx object
#'
#' @param ppt an rpptx object
#' @param title character string for the title text
#' @param title_fp an \code{fp_text} object for title formatting
#' @param left left position of the title in the same units as the plot
#' @param top top position of the plot in the same units as the plot
#' @param width width of the title text box in the same units as the plot
#' @param cm2inch conversion factor (2.54 when using cm, 1 when using inches)
#'
#' @return the modified rpptx object
#'
#' @noRd
add_slide_title <- function(ppt, title, title_fp, left, top, width, cm2inch) {
    title_height <- 1.2 / cm2inch
    title_top <- top / cm2inch - title_height

    if (is.null(title_fp)) {
        title_fp <- fp_text(font.family = "ArialMT", font.size = 18, bold = TRUE)
    }

    ppt <- ppt %>% ph_with(
        fpar(ftext(title, prop = title_fp)),
        ph_location(
            left = left / cm2inch,
            top = title_top,
            width = width / cm2inch,
            height = title_height
        )
    )

    ppt
}

#' Resolve the PowerPoint template file path
#'
#' @param template path to a custom template file, or NULL to use defaults
#'
#' @return path to the template file
#'
#' @noRd
get_template <- function(template = NULL) {
    if (!is.null(template)) {
        if (!file.exists(template)) {
            stop("Template file does not exist: ", template)
        }
        return(template)
    }

    opt <- getOption("tgppt.template")
    if (!is.null(opt)) {
        if (!file.exists(opt)) {
            stop("Template file set via options(tgppt.template) does not exist: ", opt)
        }
        return(opt)
    }

    system.file("ppt", "template.pptx", package = "tgppt")
}

#' Generate an empty pptx file
#'
#' @param fn file name of the new pptx file
#' @param template path to a custom pptx template file. When NULL, uses
#'   `getOption("tgppt.template")` if set, otherwise the bundled template.
#'
#' @return logical value indicating if the operation succeeded
#'
#' @examples
#' ppt_file <- tempfile(fileext = ".pptx")
#' new_ppt(ppt_file)
#' @export
new_ppt <- function(fn, template = NULL) {
    file.copy(get_template(template), fn)
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
#' @param sep_legend plot legend and plot separately
#' @param transparent_bg make background of the plot transparent
#' @param rasterize_plot rasterize the plotting panel
#' @param rasterize_legend rasterize the legend panel. Works only when sep_legend=TRUE and rasterize_plot=TRUE
#' @param rasterize_text rasterize text elements (geom_text, geom_label, geom_text_repel, etc.) within the panel. When FALSE (default), text grobs are kept as vector graphics even when rasterize_plot=TRUE.
#' @param rasterize_geoms character vector of geom name prefixes to selectively rasterize
#'   (e.g., \code{c("geom_point", "geom_tile")}). When not NULL, only the matching geom
#'   layers within panel grobs are rasterized to PNG while the remaining geom layers are
#'   rendered as vector DML. Setting this implicitly enables \code{rasterize_plot}.
#' @param res resolution of png used to generate rasterized plot
#' @param new_slide add plot to a new slide
#' @param overwrite overwrite existing file
#' @param legend_height,legend_width,legend_left,legend_top legend position and size in case \code{sep_legend=TRUE}. Similar to to height,width,left and top. By default - legend would be plotted to the right of the plot.
#' @param tmpdir temporary directory to store the rasterized png intermediate files
#' @param template path to a custom pptx template file. When NULL, uses
#'   `getOption("tgppt.template")` if set, otherwise the bundled template.
#' @param title optional slide title string. When not NULL, a title text box is
#'   placed above the plot.
#' @param title_fp an \code{fp_text} object controlling title font properties.
#'   When NULL (the default), uses 18pt bold ArialMT.
#'
#' @return invisible; called for its side effect of writing a pptx file.
#'
#' @examples
#'
#' library(ggplot2)
#' gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
#'     geom_point()
#' temp_ppt <- tempfile(fileext = ".pptx")
#' plot_gg_ppt(gg, temp_ppt)
#' @export
plot_gg_ppt <- function(gg, out_ppt, height = 6, width = 6, left = 5, top = 5, inches = FALSE, sep_legend = FALSE, transparent_bg = TRUE, rasterize_plot = FALSE, rasterize_legend = FALSE, rasterize_text = FALSE, rasterize_geoms = NULL, res = 300, new_slide = FALSE, overwrite = FALSE, legend_height = NULL, legend_width = NULL, legend_left = NULL, legend_top = NULL, tmpdir = tempdir(), template = NULL, title = NULL, title_fp = NULL) {
    cm2inch <- 1
    if (!inches) {
        cm2inch <- 2.54
    }

    if (!is.null(rasterize_geoms)) {
        rasterize_plot <- TRUE
    }

    if (!file.exists(out_ppt) || overwrite) {
        ppt <- read_pptx(get_template(template))
    } else {
        ppt <- read_pptx(out_ppt)
        if (new_slide) {
            ppt <- ppt %>% add_slide(layout = "Blank", master = layout_summary(ppt)$master[1])
        }
    }

    if (!is.null(title)) {
        ppt <- add_slide_title(ppt, title, title_fp, left, top, width, cm2inch)
    }

    if (transparent_bg) {
        gg <- gg + theme(
            panel.background = element_rect(fill = "transparent", color = NA),
            plot.background = element_rect(fill = "transparent", color = NA)
        )
    }

    if (sep_legend) {
        legend_height <- legend_height %||% height
        legend_width <- legend_width %||% width
        legend_top <- legend_top %||% top
        legend_left <- legend_left %||% left + legend_width
    }

    if (rasterize_plot) {
        if (sep_legend) {
            leg <- cowplot::get_legend(gg)
            gg <- gg + theme(legend.position = "none")
        }

        # plot the panel and other elements separately
        gt <- cowplot::as_gtable(gg)

        if (!is.null(rasterize_geoms)) {
            # Selective geom rasterization: split panel children by geom name
            geom_split <- separate_geom_grobs(gt, rasterize_geoms)
            gt_rasterize <- geom_split$rasterize
            gt_vector_geoms <- geom_split$vector

            if (!rasterize_text) {
                # Separate text grobs from the rasterize set
                text_sep <- separate_text_grobs(gt_rasterize)
                gt_rasterize <- text_sep$no_text
                gt_text_only <- text_sep$text_only
            }

            # Rasterize panel from the matched-geom gtable
            gt_panel <- gt_rasterize
            gt_panel$grobs <- gt_rasterize$grobs %>% modify_at(grep("panel", gt_rasterize$layout$name, invert = TRUE), ~ grid::nullGrob())

            fn <- tempfile(fileext = ".png", tmpdir = tmpdir)
            rasterize_grob(gt_panel, fn, width, height, res, cm2inch)

            # Vector geom panel overlay
            gt_vector_panel <- gt_vector_geoms
            gt_vector_panel$grobs <- gt_vector_geoms$grobs %>% modify_at(grep("panel", gt_vector_geoms$layout$name, invert = TRUE), ~ grid::nullGrob())

            # Non-panel elements (axes, labels) as vector graphics
            gt_other <- gt
            gt_other$grobs <- gt$grobs %>% modify_at(grep("panel", gt$layout$name), ~ grid::nullGrob())
            gt_other$grobs <- gt_other$grobs %>% modify_at(grep("background", gt$layout$name), ~ grid::nullGrob())

            # Layer order: rasterized geoms PNG (bottom) -> vector geom DML -> non-panel DML -> text DML (top)
            ppt <- ppt %>% ph_with(external_img(fn), ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))

            vector_geom_overlay <- dml(grid::grid.draw(gt_vector_panel), bg = "transparent")
            ppt <- ppt %>% ph_with(vector_geom_overlay, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))

            plot <- dml(grid::grid.draw(gt_other), bg = "transparent")
            ppt <- ppt %>% ph_with(plot, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))

            if (!rasterize_text) {
                text_overlay <- dml(grid::grid.draw(gt_text_only), bg = "transparent")
                ppt <- ppt %>% ph_with(text_overlay, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))
            }
        } else {
            if (!rasterize_text) {
                # Separate text grobs from non-text grobs in the panel
                separated <- separate_text_grobs(gt)
                gt_no_text <- separated$no_text
                gt_text_only <- separated$text_only

                # Plot the non-text panel to png
                gt_panel <- gt_no_text
                gt_panel$grobs <- gt_no_text$grobs %>% modify_at(grep("panel", gt_no_text$layout$name, invert = TRUE), ~ grid::nullGrob())
            } else {
                # Plot the full panel to png (original behavior)
                gt_panel <- gt
                gt_panel$grobs <- gt$grobs %>% modify_at(grep("panel", gt$layout$name, invert = TRUE), ~ grid::nullGrob())
            }

            fn <- tempfile(fileext = ".png", tmpdir = tmpdir)
            rasterize_grob(gt_panel, fn, width, height, res, cm2inch)

            # Plot other elements as vector graphics
            gt_other <- gt
            gt_other$grobs <- gt$grobs %>% modify_at(grep("panel", gt$layout$name), ~ grid::nullGrob())
            gt_other$grobs <- gt_other$grobs %>% modify_at(grep("background", gt$layout$name), ~ grid::nullGrob())

            ppt <- ppt %>% ph_with(external_img(fn), ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))
            plot <- dml(grid::grid.draw(gt_other), bg = "transparent")
            ppt <- ppt %>% ph_with(plot, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))

            if (!rasterize_text) {
                # Add text grobs as a vector overlay on top
                text_overlay <- dml(grid::grid.draw(gt_text_only), bg = "transparent")
                ppt <- ppt %>% ph_with(text_overlay, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))
            }
        }

        if (sep_legend) {
            if (rasterize_legend) {
                fn_leg <- tempfile(fileext = ".png", tmpdir = tmpdir)
                rasterize_grob(leg, fn_leg, width, height, res, cm2inch)
                ppt <- ppt %>% ph_with(external_img(fn_leg), ph_location(height = legend_height / cm2inch, width = legend_width / cm2inch, left = legend_left / cm2inch, top = legend_top / cm2inch))
            } else {
                legend <- dml(grid::grid.draw(leg), bg = "transparent")
                ppt <- ppt %>% ph_with(legend, ph_location(height = legend_height / cm2inch, width = legend_width / cm2inch, left = legend_left / cm2inch, top = legend_top / cm2inch))
            }
        }
    } else {
        if (sep_legend) {
            p <- dml(ggobj = gg + theme(legend.position = "none"), bg = "transparent")
            ppt <- ppt %>% ph_with(p, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))

            legend <- dml(grid::grid.draw(cowplot::get_legend(gg)), bg = "transparent")
            ppt <- ppt %>% ph_with(legend, ph_location(height = legend_height / cm2inch, width = legend_width / cm2inch, left = legend_left / cm2inch, top = legend_top / cm2inch))
        } else {
            code <- dml(ggobj = gg, bg = "transparent")
            ppt <- ppt %>% ph_with(code, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))
        }
    }

    print(ppt, target = out_ppt)
}

#' Plot multiple ggplot objects on a single slide
#'
#' Arrange several ggplot objects in a grid layout on one PowerPoint slide.
#' Each plot is placed using \code{\link{plot_gg_ppt}}, so all of its options
#' (rasterisation, transparent background, etc.) are available via \code{...}.
#'
#' @param plots a list of ggplot objects
#' @param out_ppt output pptx file
#' @param ncol number of columns in the grid. When NULL, defaults to
#'   \code{ceiling(sqrt(length(plots)))}.
#' @param nrow number of rows in the grid. When NULL, defaults to
#'   \code{ceiling(length(plots) / ncol)}.
#' @param titles optional character vector of titles, one per plot. When not
#'   NULL, \code{titles[[i]]} is passed as \code{title} to
#'   \code{\link{plot_gg_ppt}} for the \emph{i}-th plot.
#' @param height height of each cell in cm. When NULL, computed from
#'   \code{slide_height}, \code{top}, and \code{nrow}.
#' @param width width of each cell in cm. When NULL, computed from
#'   \code{slide_width}, \code{left}, and \code{ncol}.
#' @param left left margin in cm
#' @param top top margin in cm
#' @param slide_width total slide width in cm. When NULL (default), read from
#'   the template.
#' @param slide_height total slide height in cm. When NULL (default), read from
#'   the template.
#' @param template path to a custom pptx template file. When NULL, uses
#'   \code{getOption("tgppt.template")} if set, otherwise the bundled template.
#' @param overwrite overwrite existing file
#' @param new_slide add plots to a new slide
#' @param ... additional arguments passed to \code{\link{plot_gg_ppt}} (e.g.
#'   \code{transparent_bg}, \code{rasterize_plot}).
#'
#' @return invisible NULL
#'
#' @examples
#' library(ggplot2)
#' plots <- list(
#'     ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
#'     ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point()
#' )
#' temp_ppt <- tempfile(fileext = ".pptx")
#' plot_multi_gg_ppt(plots, temp_ppt)
#' @export
plot_multi_gg_ppt <- function(plots, out_ppt, ncol = NULL, nrow = NULL,
                               titles = NULL, height = NULL, width = NULL,
                               left = 0.5, top = 0.5, slide_width = NULL,
                               slide_height = NULL, template = NULL,
                               overwrite = FALSE, new_slide = FALSE, ...) {
    n <- length(plots)
    if (n == 0) {
        stop("plots must be a non-empty list")
    }

    # Read slide dimensions from template if not specified
    if (is.null(slide_width) || is.null(slide_height)) {
        tmpl <- get_template(template)
        ppt_tmp <- read_pptx(tmpl)
        sz <- slide_size(ppt_tmp)
        slide_width <- slide_width %||% (sz$width * 2.54)
        slide_height <- slide_height %||% (sz$height * 2.54)
    }

    ncol <- ncol %||% ceiling(sqrt(n))
    nrow <- nrow %||% ceiling(n / ncol)

    cell_width <- width %||% ((slide_width - 2 * left) / ncol)
    cell_height <- height %||% ((slide_height - 2 * top) / nrow)

    for (i in seq_len(n)) {
        col <- (i - 1) %% ncol
        row <- (i - 1) %/% ncol
        plot_left <- left + col * cell_width
        plot_top <- top + row * cell_height

        title_i <- if (!is.null(titles)) titles[[i]] else NULL

        if (i == 1) {
            plot_gg_ppt(
                plots[[i]], out_ppt,
                height = cell_height, width = cell_width,
                left = plot_left, top = plot_top,
                overwrite = overwrite, new_slide = new_slide,
                template = template, title = title_i, ...
            )
        } else {
            plot_gg_ppt(
                plots[[i]], out_ppt,
                height = cell_height, width = cell_width,
                left = plot_left, top = plot_top,
                overwrite = FALSE, new_slide = FALSE,
                template = template, title = title_i, ...
            )
        }
    }

    invisible(NULL)
}

#' Plot a list of ggplot objects to a PowerPoint file, one per slide
#'
#' Creates a PowerPoint file with one slide per plot. Each plot is rendered
#' via \code{\link{plot_gg_ppt}}, so all of its options (rasterisation,
#' transparent background, etc.) are available via \code{...}.
#'
#' @param plots a list of ggplot objects
#' @param out_ppt output pptx file
#' @param titles optional character vector of slide titles, one per plot.
#'   When not NULL, must have the same length as \code{plots}.
#' @param template path to a custom pptx template file. When NULL, uses
#'   \code{getOption("tgppt.template")} if set, otherwise the bundled template.
#' @param overwrite overwrite existing file
#' @param ... additional arguments passed to \code{\link{plot_gg_ppt}} (e.g.
#'   \code{height}, \code{width}, \code{left}, \code{top},
#'   \code{rasterize_plot}).
#'
#' @return invisible NULL
#'
#' @examples
#' library(ggplot2)
#' plots <- list(
#'     ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
#'     ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point(),
#'     ggplot(mtcars, aes(x = hp, y = mpg)) + geom_point()
#' )
#' temp_ppt <- tempfile(fileext = ".pptx")
#' plot_list_ppt(plots, temp_ppt)
#'
#' # With per-slide titles
#' temp_ppt2 <- tempfile(fileext = ".pptx")
#' plot_list_ppt(plots, temp_ppt2, titles = c("Plot 1", "Plot 2", "Plot 3"))
#' @export
plot_list_ppt <- function(plots, out_ppt, titles = NULL, template = NULL,
                           overwrite = FALSE, ...) {
    if (!is.list(plots)) {
        stop("plots must be a list")
    }
    n <- length(plots)
    if (n == 0) {
        stop("plots must be a non-empty list")
    }
    if (!is.null(titles) && length(titles) != n) {
        stop("titles must have the same length as plots")
    }

    for (i in seq_len(n)) {
        title_i <- if (!is.null(titles)) titles[[i]] else NULL

        if (i == 1) {
            plot_gg_ppt(
                plots[[i]], out_ppt,
                overwrite = overwrite,
                template = template,
                title = title_i, ...
            )
        } else {
            plot_gg_ppt(
                plots[[i]], out_ppt,
                new_slide = TRUE,
                template = template,
                title = title_i, ...
            )
        }
    }

    invisible(NULL)
}

separate_geom_grobs <- function(gt, rasterize_geoms) {
    panel_idx <- grep("panel", gt$layout$name)

    gt_rasterize <- gt
    gt_vector <- gt

    # Null out non-panel grobs in the rasterize gtable (keep only panels)
    non_panel_idx <- grep("panel", gt$layout$name, invert = TRUE)
    gt_rasterize$grobs <- gt$grobs %>% modify_at(non_panel_idx, ~ grid::nullGrob())
    bg_idx <- grep("background", gt$layout$name)
    if (length(bg_idx) > 0) {
        gt_rasterize$grobs <- gt_rasterize$grobs %>% modify_at(bg_idx, ~ grid::nullGrob())
    }

    for (i in panel_idx) {
        panel_grob <- gt$grobs[[i]]
        if (!inherits(panel_grob, "gTree") || is.null(panel_grob$children)) {
            next
        }

        children <- panel_grob$children
        n_children <- length(children)

        rasterize_children <- children
        vector_children <- children

        for (j in seq_len(n_children)) {
            child <- children[[j]]
            matches_geom <- FALSE
            if (!is.null(child$name)) {
                for (prefix in rasterize_geoms) {
                    if (startsWith(child$name, prefix)) {
                        matches_geom <- TRUE
                        break
                    }
                }
            }

            if (matches_geom) {
                vector_children[[j]] <- grid::nullGrob()
            } else {
                rasterize_children[[j]] <- grid::nullGrob()
            }
        }

        gt_rasterize$grobs[[i]] <- grid::setChildren(panel_grob, rasterize_children)
        gt_vector$grobs[[i]] <- grid::setChildren(panel_grob, vector_children)
    }

    list(rasterize = gt_rasterize, vector = gt_vector)
}

separate_text_grobs <- function(gt) {
    text_classes <- c("text", "titleGrob", "textrepeltree", "labelrepeltree", "richtext_grob", "textbox_grob")
    panel_idx <- grep("panel", gt$layout$name)

    gt_no_text <- gt
    gt_text_only <- gt

    # Null out non-panel grobs in the text-only gtable
    gt_text_only$grobs <- gt$grobs %>% modify_at(grep("panel", gt$layout$name, invert = TRUE), ~ grid::nullGrob())
    # Also null the background in the text-only gtable
    bg_idx <- grep("background", gt$layout$name)
    if (length(bg_idx) > 0) {
        gt_text_only$grobs <- gt_text_only$grobs %>% modify_at(bg_idx, ~ grid::nullGrob())
    }

    for (i in panel_idx) {
        panel_grob <- gt$grobs[[i]]
        if (!inherits(panel_grob, "gTree") || is.null(panel_grob$children)) {
            next
        }

        children <- panel_grob$children
        n_children <- length(children)

        no_text_children <- children
        text_only_children <- children

        for (j in seq_len(n_children)) {
            child <- children[[j]]
            is_text <- inherits(child, text_classes)
            if (!is_text && inherits(child, "gTree") && !is.null(child$name)) {
                if (startsWith(child$name, "geom_label")) {
                    is_text <- TRUE
                }
            }

            if (is_text) {
                no_text_children[[j]] <- grid::nullGrob()
            } else {
                text_only_children[[j]] <- grid::nullGrob()
            }
        }

        gt_no_text$grobs[[i]] <- grid::setChildren(panel_grob, no_text_children)
        gt_text_only$grobs[[i]] <- grid::setChildren(panel_grob, text_only_children)
    }

    list(no_text = gt_no_text, text_only = gt_text_only)
}

rasterize_grob <- function(grob, fn, width, height, res, cm2inch) {
    old_dev <- grDevices::dev.cur()
    grDevices::png(fn, width = width / cm2inch, height = height / cm2inch, units = "in", res = res)
    on.exit(utils::capture.output({
        if (old_dev > 1) grDevices::dev.set(old_dev)
    }))
    grid::grid.draw(grob)
    grDevices::dev.off()
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
#' @param overwrite overwrite existing file
#' @param template path to a custom pptx template file. When NULL, uses
#'   `getOption("tgppt.template")` if set, otherwise the bundled template.
#' @param title optional slide title string. When not NULL, a title text box is
#'   placed above the plot.
#' @param title_fp an \code{fp_text} object controlling title font properties.
#'   When NULL (the default), uses 18pt bold ArialMT.
#'
#' @return invisible; called for its side effect of writing a pptx file.
#'
#' @examples
#'
#' temp_ppt <- tempfile(fileext = ".pptx")
#' plot_base_ppt(
#'     {
#'         plot(mtcars$mpg, mtcars$drat)
#'     },
#'     temp_ppt
#' )
#' @export
plot_base_ppt <- function(code, out_ppt, height = 6, width = 6, left = 5, top = 5, inches = FALSE, new_slide = FALSE, overwrite = FALSE, template = NULL, title = NULL, title_fp = NULL) {
    cm2inch <- 1
    if (!inches) {
        cm2inch <- 2.54
    }

    if (!file.exists(out_ppt) || overwrite) {
        ppt <- read_pptx(get_template(template))
    } else {
        ppt <- read_pptx(out_ppt)
        if (new_slide) {
            ppt <- ppt %>% add_slide(layout = "Blank", master = layout_summary(ppt)$master[1])
        }
    }

    if (!is.null(title)) {
        ppt <- add_slide_title(ppt, title, title_fp, left, top, width, cm2inch)
    }

    out <- list()
    out$code <- enquo(code)
    out$bg <- "white"
    out$fonts <- list()
    out$pointsize <- 12
    out$editable <- TRUE
    class(out) <- "dml"


    ppt <- ppt %>% ph_with(out, ph_location(height = height / cm2inch, width = width / cm2inch, left = left / cm2inch, top = top / cm2inch))

    print(ppt, target = out_ppt)
}

#' Add a table to a PowerPoint slide
#'
#' Creates a formatted table on a PowerPoint slide using the flextable package.
#'
#' @param df a data frame to render as a table
#' @param out_ppt output pptx file
#' @param left left position in cm (or inches if \code{inches = TRUE})
#' @param top top position in cm (or inches if \code{inches = TRUE})
#' @param width width of the table in cm (or inches if \code{inches = TRUE})
#' @param height height of the table in cm (or inches if \code{inches = TRUE})
#' @param inches logical; if TRUE, dimensions are in inches instead of cm
#' @param title optional slide title string. When not NULL, a title text box is
#'   placed above the table.
#' @param title_fp an \code{fp_text} object controlling title font properties.
#'   When NULL (the default), uses 18pt bold ArialMT.
#' @param template path to a custom pptx template file. When NULL, uses
#'   \code{getOption("tgppt.template")} if set, otherwise the bundled template.
#' @param new_slide logical; add table to a new slide
#' @param overwrite logical; overwrite existing file
#' @param font_size font size for the table text
#' @param header_bold logical; if TRUE, header row text is bold
#' @param header_bg background colour for the header row
#' @param header_color text colour for the header row
#' @param alternating_rows logical; if TRUE, even rows get a coloured background
#' @param even_row_bg background colour for even rows when
#'   \code{alternating_rows = TRUE}
#'
#' @return invisible; called for its side effect of writing a pptx file.
#'
#' @examples
#' temp_ppt <- tempfile(fileext = ".pptx")
#' add_table_ppt(head(mtcars), temp_ppt)
#' @export
add_table_ppt <- function(df, out_ppt, left = 2, top = 4, width = 25, height = 15,
                           inches = FALSE, title = NULL, title_fp = NULL,
                           template = NULL, new_slide = FALSE, overwrite = FALSE,
                           font_size = 10, header_bold = TRUE,
                           header_bg = "#4472C4", header_color = "white",
                           alternating_rows = TRUE, even_row_bg = "#D9E2F3") {
    if (!requireNamespace("flextable", quietly = TRUE)) {
        stop("Package 'flextable' is required for add_table_ppt(). Install it with: install.packages('flextable')")
    }

    cm2inch <- 1
    if (!inches) {
        cm2inch <- 2.54
    }

    if (!file.exists(out_ppt) || overwrite) {
        ppt <- read_pptx(get_template(template))
    } else {
        ppt <- read_pptx(out_ppt)
        if (new_slide) {
            ppt <- ppt %>% add_slide(layout = "Blank", master = layout_summary(ppt)$master[1])
        }
    }

    if (!is.null(title)) {
        ppt <- add_slide_title(ppt, title, title_fp, left, top, width, cm2inch)
    }

    ft <- flextable::flextable(df)
    ft <- flextable::fontsize(ft, size = font_size, part = "all")
    ft <- flextable::font(ft, fontname = "ArialMT", part = "all")

    if (header_bold) {
        ft <- flextable::bold(ft, part = "header")
    }
    ft <- flextable::bg(ft, bg = header_bg, part = "header")
    ft <- flextable::color(ft, color = header_color, part = "header")

    if (alternating_rows && nrow(df) >= 2) {
        ft <- flextable::bg(ft, i = seq(2, nrow(df), by = 2), bg = even_row_bg)
    }

    ft <- flextable::autofit(ft)

    ppt <- ppt %>% ph_with(
        ft,
        ph_location(
            left = left / cm2inch,
            top = top / cm2inch,
            width = width / cm2inch,
            height = height / cm2inch
        )
    )

    print(ppt, target = out_ppt)
}
