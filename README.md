
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
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/RtmpoOcEx7/file16e3772d48e9.pptx"
```

Plot ggplot to a powerpoint presentation:

``` r
library(tgppt)
library(ggplot2)
gg <- ggplot(mtcars, aes(x=mpg, y=drat)) + geom_point()
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/RtmpoOcEx7/file16e37c34ee3f.pptx"
```

Create a new powerpoint file:

``` r
library(tgppt)
new_ppt("myfile.pptx")
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
gg <- ggplot() + geom_point_rast(aes(x=rnorm(1000), y=rnorm(1000)))
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/RtmpoOcEx7/file16e34fe2d700.pptx"
```

The same problem might occur when plotting a boxplot with many outliers:

``` r
library(tgppt)
library(ggplot2)
gg <- ggplot() + geom_boxplot_jitter(aes(y=rt(1000, df=3), x=as.factor(1:1000 %% 2)), outlier.jitter.width = 0.1, raster = TRUE)
temp_ppt <- tempfile(fileext = ".pptx")
plot_gg_ppt(gg, temp_ppt)
#> [1] "/private/var/folders/13/t30dpv4n43qcb40g1mdzwgq00000gp/T/RtmpoOcEx7/file16e3a71b796.pptx"
```
