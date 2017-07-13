################################################################################
#
# Read the QlikView daily inventory file.  The data comes from the Inventory
# cube, and the SUS Rollstock bookmark
#
sinv <- read.xlsx(file.path(proj_root, "Data", 
                            "QV-daily-inventory.xlsx"),
                  detectDates = T)

# Convert to long format
