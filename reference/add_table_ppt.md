# Add a table to a PowerPoint slide

Creates a formatted table on a PowerPoint slide using the flextable
package.

## Usage

``` r
add_table_ppt(
  df,
  out_ppt,
  left = 2,
  top = 4,
  width = 25,
  height = 15,
  inches = FALSE,
  title = NULL,
  title_fp = NULL,
  template = NULL,
  new_slide = FALSE,
  overwrite = FALSE,
  font_size = 10,
  header_bold = TRUE,
  header_bg = "#4472C4",
  header_color = "white",
  alternating_rows = TRUE,
  even_row_bg = "#D9E2F3"
)
```

## Arguments

- df:

  a data frame to render as a table

- out_ppt:

  output pptx file

- left:

  left position in cm (or inches if `inches = TRUE`)

- top:

  top position in cm (or inches if `inches = TRUE`)

- width:

  width of the table in cm (or inches if `inches = TRUE`)

- height:

  height of the table in cm (or inches if `inches = TRUE`)

- inches:

  logical; if TRUE, dimensions are in inches instead of cm

- title:

  optional slide title string. When not NULL, a title text box is placed
  above the table.

- title_fp:

  an `fp_text` object controlling title font properties. When NULL (the
  default), uses 18pt bold ArialMT.

- template:

  path to a custom pptx template file. When NULL, uses
  `getOption("tgppt.template")` if set, otherwise the bundled template.

- new_slide:

  logical; add table to a new slide

- overwrite:

  logical; overwrite existing file

- font_size:

  font size for the table text

- header_bold:

  logical; if TRUE, header row text is bold

- header_bg:

  background colour for the header row

- header_color:

  text colour for the header row

- alternating_rows:

  logical; if TRUE, even rows get a coloured background

- even_row_bg:

  background colour for even rows when `alternating_rows = TRUE`

## Value

invisible; called for its side effect of writing a pptx file.

## Examples

``` r
temp_ppt <- tempfile(fileext = ".pptx")
add_table_ppt(head(mtcars), temp_ppt)
```
