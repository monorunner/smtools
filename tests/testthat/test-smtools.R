context("Test smtools")
library(smtools)

test_that("clean.names works", {
  expect_equal(clean.names(data.table(AsjkB=1)), data.table(asjkb=1))
  expect_equal(clean.names(data.table(sa_ds_1=1)), data.table(sa.ds1=1))
})

test_that("excel.names works",{
  expect_equal(excel.names(data.table(po.zf=1)), data.table(`Po Zf`=1))
})
