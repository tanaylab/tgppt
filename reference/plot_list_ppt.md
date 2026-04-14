# Plot a list of ggplot objects to a PowerPoint file, one per slide

Creates a PowerPoint file with one slide per plot. Each plot is rendered
via
[`plot_gg_ppt`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md),
so all of its options (rasterisation, transparent background, etc.) are
available via `...`.

## Usage

``` r
plot_list_ppt(
  plots,
  out_ppt,
  titles = NULL,
  template = NULL,
  overwrite = FALSE,
  ...
)
```

## Arguments

- plots:

  a list of ggplot objects

- out_ppt:

  output pptx file

- titles:

  optional character vector of slide titles, one per plot. When not
  NULL, must have the same length as `plots`.

- template:

  path to a custom pptx template file. When NULL, uses
  `getOption("tgppt.template")` if set, otherwise the bundled template.

- overwrite:

  overwrite existing file

- ...:

  additional arguments passed to
  [`plot_gg_ppt`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md)
  (e.g. `height`, `width`, `left`, `top`, `rasterize_plot`).

## Value

invisible NULL

## Examples

``` r
library(ggplot2)
plots <- list(
    ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
    ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point(),
    ggplot(mtcars, aes(x = hp, y = mpg)) + geom_point()
)
temp_ppt <- tempfile(fileext = ".pptx")
plot_list_ppt(plots, temp_ppt)

# With per-slide titles
temp_ppt2 <- tempfile(fileext = ".pptx")
plot_list_ppt(plots, temp_ppt2, titles = c("Plot 1", "Plot 2", "Plot 3"))
```
