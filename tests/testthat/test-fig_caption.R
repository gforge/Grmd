fn <- "./tests/testthat/Full_test_suite.Rmd"
rmarkdown::render(fn, output_file = "./tmp.html")

file.remove("./tests/testthat/docx.css")
unlink("./tests/testthat/tmp_files", recursive = TRUE)
