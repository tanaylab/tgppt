context("plot utils")
test_that("vertical_labs works", {
    p <- ggplot(mtcars, aes(x = factor(cyl), y = wt)) + geom_col() + vertical_labs()
    expect_is(p, "ggplot")
})

test_that("geom_point_rast works", {
    p <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point_rast()
    expect_is(p, "ggplot")
})

test_that("geom_boxplot_jitter works", {
    p <- ggplot(mtcars, aes(x=factor(cyl), y=drat)) + geom_boxplot_jitter()
    expect_is(p, "ggplot")
})