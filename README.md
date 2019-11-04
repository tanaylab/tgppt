
<!-- README.md is generated from README.Rmd. Please edit that file -->

# tgppt

<!-- badges: start -->

[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://www.tidyverse.org/lifecycle/#experimental)
[![CRAN
status](https://www.r-pkg.org/badges/version/tgppt)](https://CRAN.R-project.org/package=tgppt)
[![Travis build
status](https://travis-ci.com/tanaylab/tgppt.svg?branch=master)](https://travis-ci.org/tanaylab/tgppt)
[![Codecov test
coverage](https://codecov.io/gh/tanaylab/tgppt/branch/master/graph/badge.svg)](https://codecov.io/gh/tanaylab/tgppt?branch=master)
<!-- badges: end -->

The goal of tgppt is to provide handy functions to plot directly to
powerpoint in R.

## Installation

You can install tgppt with:

``` r
remotes::install_github("tanaylab/tgppt")
```

## Example

Plot base R directly to a powerpoint presentation:

``` r
library(tgppt)
temp_ppt <- tempfile(fileext = ".pptx")
plot_base_ppt({plot(mtcars$mpg, mtcars$drat)}, temp_ppt)
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/Rtmp3OuPdG/file7176a25c4da.pptx"
```

Plot ggplot to a powerpoint presentation:

``` r
library(tgppt)
library(ggplot2)
gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/Rtmp3OuPdG/file717352d17eb.pptx"
```

Create a new powerpoint file:

``` r
library(tgppt)
new_ppt("myfile.pptx")
#> [1] TRUE
```

Use “Arial” font based ggplot theme:

``` r
ggplot2::theme_set(theme_arial(8))
```

## Note on plotting a large number of points

Powerpoint might crash when plotting a large number of points, so it is
advised to use the a rasterized version of geom\_point:

``` r
library(tgppt)
library(ggplot2)
gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point_rast()
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/Rtmp3OuPdG/file7175d67f481.pptx"
```

The same problem might occur when plotting a boxplot with many outliers:

``` r
library(tgppt)
library(ggplot2)
gg <- ggplot(mtcars, aes(x=factor(cyl), y=drat)) + geom_boxplot_jitter()
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/Rtmp3OuPdG/file717106e21c.pptx"
```
