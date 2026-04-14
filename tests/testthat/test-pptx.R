context("PPTX")
test_that("pptx file is generated", {
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_base_ppt(
        {
            plot(mtcars$mpg, mtcars$drat)
        },
        temp_ppt
    )
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot)", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot) when overwrite is true", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, overwrite = TRUE)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("pptx file is generated (ggplot) with separate legend", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, sep_legend = TRUE)
    expect_true(file.exists(temp_ppt))
})

test_that("legend is not created on a separate slide when new_slide is TRUE", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, sep_legend = TRUE, new_slide = TRUE)
    expect_true(file.exists(temp_ppt))
    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("plot on a separate slide when new_slide is TRUE", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, new_slide = TRUE)
    plot_gg_ppt(gg, temp_ppt, new_slide = TRUE)
    expect_true(file.exists(temp_ppt))
    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 2)
})

test_that("pptx file is generated (ggplot) when rasterize_plot is TRUE", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_plot = TRUE)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot) when rasterize_plot is TRUE and tmpdir is set", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    tmp_dir <- file.path(tempdir(), "test")
    dir.create(file.path(tempdir(), "test"))
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_plot = TRUE, tmpdir = tmp_dir)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot) when rasterize_plot is TRUE and there is an open device", {
    grDevices::png(tempfile(fileext = ".png"))
    plot(1)
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_plot = TRUE)
    expect_true(file.exists(temp_ppt))
    dev.off()
})


test_that("pptx file is generated (ggplot) when rasterize_plot is TRUE and sep_legend = TRUE", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_plot = TRUE, sep_legend = TRUE)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot) when rasterize_plot is TRUE, sep_legend = TRUE and rasterize_legend is TRUE", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_plot = TRUE, sep_legend = TRUE, rasterize_legend = TRUE)
    expect_true(file.exists(temp_ppt))
})

test_that("New slide is generated", {
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_base_ppt(
        {
            plot(mtcars$mpg, mtcars$drat)
        },
        temp_ppt
    )
    plot_base_ppt(
        {
            plot(mtcars$mpg, mtcars$drat)
        },
        temp_ppt,
        new_slide = TRUE
    )

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 2)
})

test_that("pptx is generated with new_ppt", {
    temp_ppt <- tempfile(fileext = ".pptx")
    new_ppt(temp_ppt)
    expect_true(file.exists(temp_ppt))
})

test_that("template parameter works with explicit path", {
    bundled <- system.file("ppt", "template.pptx", package = "tgppt")

    temp_ppt <- tempfile(fileext = ".pptx")
    new_ppt(temp_ppt, template = bundled)
    expect_true(file.exists(temp_ppt))

    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point()
    temp_ppt2 <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt2, template = bundled)
    expect_true(file.exists(temp_ppt2))

    temp_ppt3 <- tempfile(fileext = ".pptx")
    plot_base_ppt(
        {
            plot(mtcars$mpg, mtcars$drat)
        },
        temp_ppt3,
        template = bundled
    )
    expect_true(file.exists(temp_ppt3))
})

test_that("tgppt.template option is used when set", {
    old_opt <- getOption("tgppt.template")
    on.exit(options(tgppt.template = old_opt))

    bundled <- system.file("ppt", "template.pptx", package = "tgppt")
    options(tgppt.template = bundled)

    temp_ppt <- tempfile(fileext = ".pptx")
    new_ppt(temp_ppt)
    expect_true(file.exists(temp_ppt))

    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point()
    temp_ppt2 <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt2)
    expect_true(file.exists(temp_ppt2))
})

test_that("non-existent template path raises error", {
    expect_error(
        new_ppt(tempfile(fileext = ".pptx"), template = "/no/such/file.pptx"),
        "Template file does not exist"
    )
})

test_that("non-existent template in option raises error", {
    old_opt <- getOption("tgppt.template")
    on.exit(options(tgppt.template = old_opt))

    options(tgppt.template = "/no/such/file.pptx")
    expect_error(
        new_ppt(tempfile(fileext = ".pptx")),
        "Template file set via options"
    )
})

test_that("plot_gg_ppt with title generates pptx", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, title = "My ggplot title")
    expect_true(file.exists(temp_ppt))
})

test_that("plot_base_ppt with title generates pptx", {
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_base_ppt(
        {
            plot(mtcars$mpg, mtcars$drat)
        },
        temp_ppt,
        title = "My base plot title"
    )
    expect_true(file.exists(temp_ppt))
})

test_that("custom title_fp works", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    custom_fp <- officer::fp_text(font.family = "Courier", font.size = 24, bold = FALSE, italic = TRUE)
    plot_gg_ppt(gg, temp_ppt, title = "Custom styled title", title_fp = custom_fp)
    expect_true(file.exists(temp_ppt))
})

test_that("plot_multi_gg_ppt with 2 plots generates pptx on one slide", {
    plots <- list(
        ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
        ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point()
    )
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_multi_gg_ppt(plots, temp_ppt)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("plot_multi_gg_ppt with 4 plots in 2x2 grid", {
    plots <- list(
        ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
        ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point(),
        ggplot(mtcars, aes(x = hp, y = mpg)) + geom_point(),
        ggplot(mtcars, aes(x = disp, y = hp)) + geom_point()
    )
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_multi_gg_ppt(plots, temp_ppt, ncol = 2, nrow = 2)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("plot_multi_gg_ppt with custom ncol", {
    plots <- list(
        ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
        ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point(),
        ggplot(mtcars, aes(x = hp, y = mpg)) + geom_point()
    )
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_multi_gg_ppt(plots, temp_ppt, ncol = 3)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("plot_list_ppt with 3 plots produces 3 slides", {
    plots <- list(
        ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
        ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point(),
        ggplot(mtcars, aes(x = hp, y = mpg)) + geom_point()
    )
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_list_ppt(plots, temp_ppt)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(ppt), 3)
})

test_that("plot_list_ppt with titles produces correct number of slides", {
    plots <- list(
        ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
        ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point()
    )
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_list_ppt(plots, temp_ppt, titles = c("First", "Second"))
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(ppt), 2)
})

test_that("plot_list_ppt with a single plot works", {
    plots <- list(
        ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point()
    )
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_list_ppt(plots, temp_ppt)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(ppt), 1)
})

test_that("pptx file is generated with rasterize_geoms", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point() +
        geom_line()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_geoms = c("geom_point"))
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated with rasterize_geoms and rasterize_text = FALSE", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
        geom_point() +
        geom_text(aes(label = gear), size = 3)
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_geoms = c("geom_point"), rasterize_text = FALSE)
    expect_true(file.exists(temp_ppt))
})

test_that("add_table_ppt creates a pptx with a table", {
    skip_if_not_installed("flextable")
    temp_ppt <- tempfile(fileext = ".pptx")
    add_table_ppt(head(mtcars), temp_ppt)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("add_table_ppt with title generates pptx", {
    skip_if_not_installed("flextable")
    temp_ppt <- tempfile(fileext = ".pptx")
    add_table_ppt(head(iris), temp_ppt, title = "Iris Data")
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("add_table_ppt with custom styling generates pptx", {
    skip_if_not_installed("flextable")
    temp_ppt <- tempfile(fileext = ".pptx")
    add_table_ppt(
        head(mtcars),
        temp_ppt,
        font_size = 14,
        header_bold = FALSE,
        header_bg = "#FF0000",
        header_color = "black",
        alternating_rows = FALSE
    )
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})
