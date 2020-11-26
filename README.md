
<!-- README.md is generated from README.Rmd. Please edit that file -->
tgppt
=====

<!-- badges: start -->
[![Lifecycle: experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://www.tidyverse.org/lifecycle/#experimental) [![CRAN status](https://www.r-pkg.org/badges/version/tgppt)](https://CRAN.R-project.org/package=tgppt) [![Travis build status](https://travis-ci.com/tanaylab/tgppt.svg?branch=master)](https://travis-ci.org/tanaylab/tgppt) [![Codecov test coverage](https://codecov.io/gh/tanaylab/tgppt/branch/master/graph/badge.svg)](https://codecov.io/gh/tanaylab/tgppt?branch=master) <!-- badges: end -->

The goal of tgppt is to provide handy functions to plot directly to powerpoint in R. It is based on the magnificent [officer](https://davidgohel.github.io/officer/) package from David Gohel.

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
#> 
#> Attaching package: 'tgppt'
#> The following object is masked from 'package:alutil':
#> 
#>     vertical_labs
temp_ppt <- tempfile(fileext = ".pptx")
plot_base_ppt({plot(mtcars$mpg, mtcars$drat)}, temp_ppt)
```

Plot ggplot to a powerpoint presentation:

``` r
library(tgppt)
library(ggplot2)
gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
```

Create a new powerpoint file:

``` r
library(tgppt)
new_ppt("myfile.pptx")
```

Use "Arial" font based ggplot theme:

``` r
ggplot2::theme_set(theme_arial(8))
```

Rasterize ggplot plotting area
------------------------------

Powerpoint might crash when plotting a large number of points. To avoid that, set `rasterize_plot` to `TRUE`:

``` r
plot_gg_ppt(gg, temp_ppt, rasterize_plot = TRUE)
```
