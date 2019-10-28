test_that("theme is set", {
    ggplot2::theme_set(arial_theme(8))
    expect_equal(theme_get()$text$family, "Arial")
})
