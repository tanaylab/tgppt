# Set x-axis vertical labels in ggplot2

Set x-axis vertical labels in ggplot2

## Usage

``` r
vertical_labs()
```

## Value

a [`theme`](https://ggplot2.tidyverse.org/reference/theme.html) object
that rotates x-axis labels to 90 degrees

## Examples

``` r
library(ggplot2)
ggplot(mtcars, aes(x = factor(cyl), y = mpg)) +
    geom_boxplot() +
    vertical_labs()
```
