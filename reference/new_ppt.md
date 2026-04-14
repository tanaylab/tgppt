# Generate an empty pptx file

Generate an empty pptx file

## Usage

``` r
new_ppt(fn, template = NULL)
```

## Arguments

- fn:

  file name of the new pptx file

- template:

  path to a custom pptx template file. When NULL, uses
  `getOption("tgppt.template")` if set, otherwise the bundled template.

## Value

logical value indicating if the operation succeeded

## Examples

``` r
ppt_file <- tempfile(fileext = ".pptx")
new_ppt(ppt_file)
#> [1] TRUE
```
