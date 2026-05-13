# Plot ggplot object to ppt

Plot ggplot object to ppt

## Usage

``` r
plot_gg_ppt(
  gg,
  out_ppt,
  height = 6,
  width = 6,
  left = 5,
  top = 5,
  inches = FALSE,
  sep_legend = FALSE,
  transparent_bg = TRUE,
  rasterize_plot = FALSE,
  rasterize_legend = FALSE,
  rasterize_text = FALSE,
  rasterize_geoms = NULL,
  res = 300,
  new_slide = FALSE,
  overwrite = FALSE,
  legend_height = NULL,
  legend_width = NULL,
  legend_left = NULL,
  legend_top = NULL,
  tmpdir = tempdir(),
  template = NULL,
  title = NULL,
  title_fp = NULL
)
```

## Arguments

- gg:

  ggplot object

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

- sep_legend:

  plot legend and plot separately

- transparent_bg:

  make background of the plot transparent

- rasterize_plot:

  rasterize the plotting panel

- rasterize_legend:

  rasterize the legend panel. Works only when sep_legend=TRUE and
  rasterize_plot=TRUE

- rasterize_text:

  rasterize text elements (geom_text, geom_label, geom_text_repel, etc.)
  within the panel. When FALSE (default), text grobs are kept as vector
  graphics even when rasterize_plot=TRUE.

- rasterize_geoms:

  character vector of geom name prefixes to selectively rasterize (e.g.,
  `c("geom_point", "geom_tile")`). When not NULL, only the matching geom
  layers within panel grobs are rasterized to PNG while the remaining
  geom layers are rendered as vector DML. Setting this implicitly
  enables `rasterize_plot`.

- res:

  resolution of png used to generate rasterized plot

- new_slide:

  add plot to a new slide

- overwrite:

  overwrite existing file

- legend_height, legend_width, legend_left, legend_top:

  legend position and size in case `sep_legend=TRUE`. Similar to to
  height,width,left and top. By default - legend would be plotted to the
  right of the plot.

- tmpdir:

  temporary directory to store the rasterized png intermediate files

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

library(ggplot2)
gg <- ggplot(mtcars, aes(x = mpg, y = drat)) +
    geom_point()
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
```
