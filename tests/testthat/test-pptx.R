test_that("pptx file is generated", {
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_base_ppt({plot(mtcars$mpg, mtcars$drat)}, temp_ppt)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot)", {
    gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot) when overwrite is true", {
    gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, overwrite = TRUE)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot) with separate legend", {
    gg <- ggplot(mtcars, aes(x=mpg, y=drat, color=factor(am))) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt(gg, temp_ppt, sep_legend=TRUE)
    expect_true(file.exists(temp_ppt))
})

