
<!-- README.md is generated from README.Rmd. Please edit that file -->
tgppt
=====

<!-- badges: start -->
[![Lifecycle: experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://www.tidyverse.org/lifecycle/#experimental) [![CRAN status](https://www.r-pkg.org/badges/version/tgppt)](https://CRAN.R-project.org/package=tgppt) <!-- badges: end -->

The goal of tgppt is to provide handy functions to plot directly to powerpoint in R.

Installation
------------

You can install tgppt with:

``` r
remotes::install_github("tanaylab/tgppt")
```

Example
-------

Plot base R directly to a powerpoint presentation:

``` r
library(tgppt)
temp_ppt <- tempfile(fileext = ".pptx")
plot_base_ppt_at({plot(mtcars$mpg, mtcars$drat)}, temp_ppt)
#> [1] "/tmp/Rtmp2ZwgHg/file554b7558b843b.pptx"
```

Plot ggplot to a powerpoint presentation:

``` r
library(tgppt)
gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt_at(gg, temp_ppt)
#> [1] "/tmp/Rtmp2ZwgHg/file554b75253589f.pptx"
```

Create a new powerpoint file:

``` r
library(tgppt)
new_ppt("myfile.pptx")
#> [1] TRUE
```

Use "Arial" font based ggplot theme:

``` r
ggplot2::theme_set(arial_theme(8))
```
