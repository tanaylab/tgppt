# Changelog

## tgppt 0.1.2

- [`plot_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md),
  [`plot_base_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_base_ppt.md)
  and
  [`add_table_ppt()`](https://tanaylab.github.io/tgppt/reference/add_table_ppt.md)
  now refuse a write that would add shape geometry PowerPoint cannot
  open. rvg writes an NA shape position or size as `-2147483648`;
  PowerPoint rejects such a file while LibreOffice renders it, so a
  render check cannot catch it. The usual cause is a zero-area shape -
  in practice `geom_tile(width = 0, height = 0)` used as an invisible
  layer that carries a legend; use
  `geom_point(shape = 21, alpha = 0, colour = "transparent")` instead.
- A refused call leaves the deck unchanged: the result is rendered to a
  temporary file in [`tempdir()`](https://rdrr.io/r/base/tempfile.html)
  and checked before the deck is touched, then copied into it, so
  symlinked decks, permissions, owners and hard links behave as before.
- Only what a call adds is counted. Appending to a deck that is already
  broken, or starting from such a template, is not refused, but warns.
- [`plot_multi_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_multi_gg_ppt.md)
  and
  [`plot_list_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_list_ppt.md)
  check each plot as they place it, so a refusal stops at the offending
  plot and keeps the plots placed before it; with `overwrite = TRUE` the
  old slides are already gone by then.

## tgppt 0.1.1

- New bundled widescreen template (`inst/ppt/widescreen.pptx`, 13.333 x
  7.5 in) alongside the portrait default.
- New exported
  [`tgppt_template()`](https://tanaylab.github.io/tgppt/reference/tgppt_template.md)
  resolver that returns the path to a bundled template by name
  (`"portrait"` or `"widescreen"`). Use
  `template = tgppt_template("widescreen")` (or
  `options(tgppt.template = tgppt_template("widescreen"))`) to get
  widescreen output without copying templates around.
- Fix:
  [`plot_multi_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_multi_gg_ppt.md)’s
  `slide_width` / `slide_height` arguments now also resize the
  underlying slide. Previously they only affected layout calculations,
  so passing widescreen dimensions against a portrait template produced
  an off-slide layout. Existing callers get the more-correct behavior
  automatically.
- Fix:
  [`plot_multi_gg_ppt()`](https://tanaylab.github.io/tgppt/reference/plot_multi_gg_ppt.md)
  with `titles = ...` now bumps `top` to at least 1.2 cm (the title
  height) when the caller’s `top` is smaller, with a warning, so
  first-row titles can no longer fall off the top of the slide.
- Fix: `add_slide_title()` clamps off-slide title positions to `y = 0`
  (with a warning) instead of placing them at a negative coordinate.
  Trips when `top` is less than the title height (1.2 cm or 1.2 in
  depending on units).

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
