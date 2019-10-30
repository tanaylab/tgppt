test_that("theme is set", {
    ggplot2::theme_set(theme_arial(8))
    expect_equal(theme_get()$text$family, "ArialMT")
})
