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
