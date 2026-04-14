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
  overwrite = FALSE
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
