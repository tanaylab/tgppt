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
