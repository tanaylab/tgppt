# tgppt 0.1.1

* New bundled widescreen template (`inst/ppt/widescreen.pptx`, 13.333 x 7.5 in) alongside the portrait default.
* New exported `tgppt_template()` resolver that returns the path to a bundled template by name (`"portrait"` or `"widescreen"`). Use `template = tgppt_template("widescreen")` (or `options(tgppt.template = tgppt_template("widescreen"))`) to get widescreen output without copying templates around.
* Fix: `plot_multi_gg_ppt()`'s `slide_width` / `slide_height` arguments now also resize the underlying slide. Previously they only affected layout calculations, so passing widescreen dimensions against a portrait template produced an off-slide layout. Existing callers get the more-correct behavior automatically.
* Fix: `plot_multi_gg_ppt()` with `titles = ...` now bumps `top` to at least 1.2 cm (the title height) when the caller's `top` is smaller, with a warning, so first-row titles can no longer fall off the top of the slide.
* Fix: `add_slide_title()` clamps off-slide title positions to `y = 0` (with a warning) instead of placing them at a negative coordinate. Trips when `top` is less than the title height (1.2 cm or 1.2 in depending on units).

# tgppt 0.1.0

* Custom template path support: set a default template via `options(tgppt.template = "/path/to/template.pptx")` or pass the `template` parameter directly to `new_ppt()`, `plot_gg_ppt()`, `plot_base_ppt()`, and other functions.
* Slide title support: new `title` and `title_fp` parameters on `plot_gg_ppt()`, `plot_base_ppt()`, and `add_table_ppt()` to place a formatted title text box above the content.
* New `plot_multi_gg_ppt()` function to arrange multiple ggplot objects in a grid layout on a single PowerPoint slide.
* New `plot_list_ppt()` function to create a PowerPoint file with one slide per plot from a list of ggplot objects.
* New `add_table_ppt()` function to render a data frame as a formatted table on a PowerPoint slide (requires the flextable package).
* New `rasterize_geoms` parameter in `plot_gg_ppt()` for selective geom rasterization: specify which geom layers to rasterize while keeping the rest as vector graphics.

# tgppt 0.0.10

* Added `rasterize_text` parameter to `plot_gg_ppt()` - when `rasterize_plot=TRUE`, text grobs (geom_text, geom_label, geom_text_repel, etc.) are now kept as vector graphics by default, making them selectable and editable in PowerPoint.
* Added GitHub Actions CI/CD: R-CMD-check, pkgdown site deployment, test coverage, and linting.

# tgppt 0.0.9 

* Added an option to set the temporary directory in which png files are saved.

# tgppt 0.0.8

* Fix: changed template to A4
# tgppt 0.0.7

* Removed unused dependencies
* Added a `NEWS.md` file to track changes to the package.
