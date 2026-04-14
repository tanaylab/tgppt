# Generate an empty pptx file

Generate an empty pptx file

## Usage

``` r
new_ppt(fn)
```

## Arguments

- fn:

  file name of the new pptx file

## Value

logical value indicating if the operation succeeded

## Examples

``` r
ppt_file <- tempfile(fileext = ".pptx")
new_ppt(ppt_file)
#> [1] TRUE
```
