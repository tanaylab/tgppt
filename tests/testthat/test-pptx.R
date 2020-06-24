context("PPTX")
test_that("pptx file is generated", {
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_base_ppt({
        plot(mtcars$mpg, mtcars$drat)
    }, temp_ppt)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot)", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot) when overwrite is true", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, overwrite = TRUE)
    expect_true(file.exists(temp_ppt))

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("pptx file is generated (ggplot) with separate legend", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, sep_legend = TRUE)
    expect_true(file.exists(temp_ppt))
})

test_that("legend is not created on a separate slide", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, sep_legend = TRUE, new_slide = TRUE)
    expect_true(file.exists(temp_ppt))
    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 1)
})

test_that("pptx file is generated (ggplot) when rasterize_plot is TRUE", {
    gg <- ggplot(mtcars, aes(x = mpg, y = drat, color = factor(am))) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, rasterize_plot = TRUE)
    expect_true(file.exists(temp_ppt))
})

test_that("New slide is generated", {
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_base_ppt({
        plot(mtcars$mpg, mtcars$drat)
    }, temp_ppt)
    plot_base_ppt({
        plot(mtcars$mpg, mtcars$drat)
    }, temp_ppt, new_slide = TRUE)

    ppt <- officer::read_pptx(temp_ppt)
    expect_equal(length(table(officer::pptx_summary(ppt)$slide_id)), 2)
})

test_that("pptx is generated with new_ppt", {
    temp_ppt <- tempfile(fileext = ".pptx")
    new_ppt(temp_ppt)
    expect_true(file.exists(temp_ppt))
})
