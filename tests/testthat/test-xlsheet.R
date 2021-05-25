library(xlbox)
library(magrittr)
library(openxlsx)

list_named <- list(A = data.frame(x = 1:2, y = c("a", "b")),
                   B = data.frame(x = 1:10, y = letters[1:10]))

test_that("Export for xlsheet class works", {
  tmp <- tempfile()
  on.exit(unlink(tmp))

  # test type
  # list_named %>% xlsheet %>% class %>% expect_equal(c("list", "xlsheet"))
  list_named %>% xlsheet %>% expect_s3_class("xlsheet")

  # test named list exported
  list_named %>% xlsheet %>% write_xlsx(tmp)
  expect_equal("A", read.xlsx(tmp, rows = 1, cols = 1) %>% names)
  expect_equal(list_named$A, read.xlsx(tmp, rows = 2:4, cols = 1:2))
  expect_equal("B", read.xlsx(tmp, rows = 6, cols = 1) %>% names)
  expect_equal(list_named$B, read.xlsx(tmp, rows = 6 + 1:11, cols = 1:2))

  # test unnamed list exported
  list_named %>% `names<-`(NULL) %>% xlsheet %>% write_xlsx(tmp)
  expect_equal(list_named$A, read.xlsx(tmp, rows = 1:3, cols = 1:2))
  expect_equal(list_named$B, read.xlsx(tmp, rows = 4 + 1:11, cols = 1:2))

})

