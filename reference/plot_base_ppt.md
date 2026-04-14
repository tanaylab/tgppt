# Plot base R plots to ppt

Plot base R plots to ppt

## Usage

``` r
plot_base_ppt(
  code,
  out_ppt,
  height = 6,
  width = 6,
  left = 5,
  top = 5,
  inches = FALSE,
  new_slide = FALSE,
  overwrite = FALSE,
  template = NULL,
  title = NULL,
  title_fp = NULL
)
```

## Arguments

- code:

  code that generates plot

- out_ppt:

  output pptx file

- height:

  height in cm

- width:

  width in cm

- left:

  left alignment in cm

- top:

  top alignment in cm

- inches:

  measures are given in inches

- new_slide:

  add plot to a new slide

- overwrite:

  overwrite existing file

- template:

  path to a custom pptx template file. When NULL, uses
  `getOption("tgppt.template")` if set, otherwise the bundled template.

- title:

  optional slide title string. When not NULL, a title text box is placed
  above the plot.

- title_fp:

  an `fp_text` object controlling title font properties. When NULL (the
  default), uses 18pt bold ArialMT.

## Value

invisible; called for its side effect of writing a pptx file.

## Examples

``` r
temp_ppt <- tempfile(fileext = ".pptx")
plot_base_ppt(
    {
        plot(mtcars$mpg, mtcars$drat)
    },
    temp_ppt
)
```
