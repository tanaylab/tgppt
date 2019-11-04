context("plot utils")
test_that("vertical_labs works", {
    p <- ggplot(mtcars, aes(x = factor(cyl), y = wt)) + geom_col() + vertical_labs()
    expect_is(p, "ggplot")
})
