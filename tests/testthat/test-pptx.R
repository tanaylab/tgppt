test_that("pptx file is generated", {
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_base_ppt_at({plot(mtcars$mpg, mtcars$drat)}, temp_ppt)
    expect_true(file.exists(temp_ppt))
})

test_that("pptx file is generated (ggplot)", {
    gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
    temp_ppt <- tempfile(fileext = ".pptx")
    plot_gg_ppt_at(gg, temp_ppt)
    expect_true(file.exists(temp_ppt))
})
