# Changelog

## tgppt 0.1.0

- Custom template path support: set a default template via
  `options(tgppt.template = "/path/to/template.pptx")` or pass the
  `template` parameter directly to
  [`new_ppt()`](https://tanaylab.github.io/tgppt/reference/new_ppt.md),
  [`plot_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md),
  [`plot_base_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_base_ppt.md),
  and other functions.
- Slide title support: new `title` and `title_fp` parameters on
  [`plot_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md),
  [`plot_base_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_base_ppt.md),
  and
  [`add_table_ppt()`](https://tanaylab.github.io/tgppt/reference/add_table_ppt.md)
  to place a formatted title text box above the content.
- New
  [`plot_multi_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_multi_gg_ppt.md)
  function to arrange multiple ggplot objects in a grid layout on a
  single PowerPoint slide.
- New
  [`plot_list_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_list_ppt.md)
  function to create a PowerPoint file with one slide per plot from a
  list of ggplot objects.
- New
  [`add_table_ppt()`](https://tanaylab.github.io/tgppt/reference/add_table_ppt.md)
  function to render a data frame as a formatted table on a PowerPoint
  slide (requires the flextable package).
- New `rasterize_geoms` parameter in
  [`plot_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md)
  for selective geom rasterization: specify which geom layers to
  rasterize while keeping the rest as vector graphics.

## tgppt 0.0.10

- Added `rasterize_text` parameter to
  [`plot_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md) -
  when `rasterize_plot=TRUE`, text grobs (geom_text, geom_label,
  geom_text_repel, etc.) are now kept as vector graphics by default,
  making them selectable and editable in PowerPoint.
- Added GitHub Actions CI/CD: R-CMD-check, pkgdown site deployment, test
  coverage, and linting.

## tgppt 0.0.9

- Added an option to set the temporary directory in which png files are
  saved.

## tgppt 0.0.8

- Fix: changed template to A4 \# tgppt 0.0.7

- Removed unused dependencies

- Added a `NEWS.md` file to track changes to the package.
