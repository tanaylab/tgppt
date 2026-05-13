# Path to a bundled tgppt PowerPoint template

Resolves a named template to a file path. Useful as the `template`
argument of
[`plot_gg_ppt`](https://tanaylab.github.io/tgppt/reference/plot_gg_ppt.md),
[`plot_multi_gg_ppt`](https://tanaylab.github.io/tgppt/reference/plot_multi_gg_ppt.md),
[`plot_list_ppt`](https://tanaylab.github.io/tgppt/reference/plot_list_ppt.md),
[`plot_base_ppt`](https://tanaylab.github.io/tgppt/reference/plot_base_ppt.md),
and
[`add_table_ppt`](https://tanaylab.github.io/tgppt/reference/add_table_ppt.md).

## Usage

``` r
tgppt_template(name = c("portrait", "widescreen"))
```

## Arguments

- name:

  template name. Built-in options:

  "portrait"

  :   A4 portrait (19.05 x 27.52 cm). Default. Suitable for paper-style
      figures.

  "widescreen"

  :   16:9 widescreen (33.87 x 19.05 cm = 13.333 x 7.5 in). Suitable for
      talks and multi-panel summary slides.

## Value

the absolute path to the template file

## Examples

``` r
tgppt_template()              # portrait
#> [1] "/home/runner/work/_temp/Library/tgppt/ppt/template.pptx"
tgppt_template("widescreen")
#> [1] "/home/runner/work/_temp/Library/tgppt/ppt/widescreen.pptx"

if (FALSE) { # \dontrun{
library(ggplot2)
p <- ggplot(mtcars, aes(wt, mpg)) + geom_point()
plot_gg_ppt(p, "deck.pptx", template = tgppt_template("widescreen"))

# Make widescreen the session-wide default:
options(tgppt.template = tgppt_template("widescreen"))
} # }
```
