library(tesseract)
tesseract()
tesseract_download("eng")
tesseract_info()
text <- ocr(file.path("C:", "Users", "SchroedR", "Documents",
                      "Projects", "SUS Roll Consolidation",
                      "trim-analysis", "21pt.jpg"))
