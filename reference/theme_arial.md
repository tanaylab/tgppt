# A ggplot2 custom theme (based on 'ArialMT' font)

A ggplot2 custom theme (based on 'ArialMT' font)

## Usage

``` r
theme_arial(size = 8)
```

## Arguments

- size:

  base size of text

## Value

a [`theme`](https://ggplot2.tidyverse.org/reference/theme.html) object

## Examples

``` r
library(ggplot2)
ggplot(mtcars, aes(x = mpg, y = drat)) +
    geom_point() +
    theme_arial()
```
