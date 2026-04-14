# Plot multiple ggplot objects on a single slide

Arrange several ggplot objects in a grid layout on one PowerPoint slide.
Each plot is placed using
[`plot_gg_ppt`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md),
so all of its options (rasterisation, transparent background, etc.) are
available via `...`.

## Usage

``` r
plot_multi_gg_ppt(
  plots,
  out_ppt,
  ncol = NULL,
  nrow = NULL,
  titles = NULL,
  height = NULL,
  width = NULL,
  left = 0.5,
  top = 0.5,
  slide_width = NULL,
  slide_height = NULL,
  template = NULL,
  overwrite = FALSE,
  new_slide = FALSE,
  ...
)
```

## Arguments

- plots:

  a list of ggplot objects

- out_ppt:

  output pptx file

- ncol:

  number of columns in the grid. When NULL, defaults to
  `ceiling(sqrt(length(plots)))`.

- nrow:

  number of rows in the grid. When NULL, defaults to
  `ceiling(length(plots) / ncol)`.

- titles:

  optional character vector of titles, one per plot. When not NULL,
  `titles[[i]]` is passed as `title` to
  [`plot_gg_ppt`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md)
  for the *i*-th plot.

- height:

  height of each cell in cm. When NULL, computed from `slide_height`,
  `top`, and `nrow`.

- width:

  width of each cell in cm. When NULL, computed from `slide_width`,
  `left`, and `ncol`.

- left:

  left margin in cm

- top:

  top margin in cm

- slide_width:

  total slide width in cm. When NULL (default), read from the template.

- slide_height:

  total slide height in cm. When NULL (default), read from the template.

- template:

  path to a custom pptx template file. When NULL, uses
  `getOption("tgppt.template")` if set, otherwise the bundled template.

- overwrite:

  overwrite existing file

- new_slide:

  add plots to a new slide

- ...:

  additional arguments passed to
  [`plot_gg_ppt`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md)
  (e.g. `transparent_bg`, `rasterize_plot`).

## Value

invisible NULL

## Examples

``` r
library(ggplot2)
plots <- list(
    ggplot(mtcars, aes(x = mpg, y = drat)) + geom_point(),
    ggplot(mtcars, aes(x = wt, y = qsec)) + geom_point()
)
temp_ppt <- tempfile(fileext = ".pptx")
plot_multi_gg_ppt(plots, temp_ppt)
```
