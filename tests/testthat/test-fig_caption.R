fn <- system.file("./tests/testthat/Full_test_suite.Rmd", package = "Grmd")

if (file.exists(fn)) {
  # TODO: test passes but must figure out how to find the file when running test check
  output_file <- tempfile(fileext = ".html")
  rmarkdown::render(fn, output_file = output_file)
}
