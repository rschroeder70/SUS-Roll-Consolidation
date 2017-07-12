#------------------------------------------------------------------------------
#
# Look at Elk Grove and build a table of the consumption by Width for each Die
#
cons <- read.xlsx(file.path(proj_root, "Data", 
                            "QV SUS Roll Consumption 2015-2017.xlsx"),
                  detectDates = T)

cons <- cons %>%
  mutate(Plant = paste0("PLT0", Plant),
         Year = as.character(year(MonthYear)),
         Board.Width = as.numeric(Board.Width)) %>%
  filter(Board.Width > 0)

cons <- cons %>%
  filter(Year == "2016")

cons <- cons %>% 
  filter(Plant == "PLT00009" & Year == "2016")

cons <- cons %>%
  rename(Material.Descr = Material.Description)

cons <- cons %>%
  mutate(M.I = str_sub(Material.Descr, 1, 1))

cons <- cons %>%
  filter(M.I %in% c("I", "M"))

cons <- cons %>% 
  mutate(Width = ifelse(M.I == "I",
                        str_sub(Material.Descr, 10, 17),
                        str_sub(Material.Descr, 11, 14)))
cons$Width <- as.numeric(cons$Width)

cons <- cons %>%
  mutate(Caliper = ifelse(M.I == "I",
                          str_sub(Material.Descr, 3, 4),
                          str_sub(Material.Descr, 4, 5)))
cons <- cons %>%
  mutate(Grade = ifelse(M.I == "I",
                        str_sub(Material.Descr, 5, 8),
                        str_sub(Material.Descr, 6, 9)))

cons <- cons %>%
  mutate(Wind = str_sub(Material.Descr, 
                        str_length(Material.Descr), -1))

cons <- cons %>%
  mutate(Dia = ifelse(M.I == "I",
                      str_sub(Material.Descr, 19, 20),
                      str_sub(Material.Descr, 16, 19)))

cons <- cons %>%
  group_by(Plant, PlantName, Year, Grade, Caliper, Wind, Width, 
           Die, Die.Description) %>%
  summarize(Tons = sum(TON))

# Summarize by Die for Ed LUsche
cons_ed <- cons %>%
  group_by(Plant, PlantName, Year, Die, Die.Description) %>%
  summarize(Tons = sum(TON)) %>%
  ungroup()

cons_ed <- cons_ed %>%
  separate(Die, c("Die1", "Die2", "Die3"), sep = "/")

cons_ed <- cons_ed %>%
  gather("Die", "n", Die1:Die3)

cons_ed <- cons_ed %>%
  select(-Die, -Tons) %>%
  rename(Die = n) %>%
  filter(!is.na(Die))
  
cons_ed <- cons_ed %>%
  distinct(Plant, Die, .keep_all = TRUE)

copy.table(cons_ed)

# Does the die have both OM and PK?
cons <- cons %>%
  group_by(Die) %>%
  mutate(OMPK = ifelse(Grade %in% c("OMXX", "PKXX") &
                         n_distinct(Grade) > 1, TRUE, FALSE))

# If both, then make the Grade OMXX to ship from Macon for freight savings
cons <- cons %>%
  mutate(Grade.Fix = ifelse(OMPK, "OMXX", Grade))

cons <- cons %>%
  group_by(Plant, PlantName, Year, Die, Die.Description, Grade.Fix, 
           Caliper, Wind, Width) %>%
  summarize(Tons = round(sum(Tons), 1))

cons <- cons %>%
  group_by(Plant, Year, Die) %>%
  mutate(Roll.Count = n_distinct(Width),
         Roll.Index = row_number())


wb <- createWorkbook()
addWorksheet(wb, "EG", tabColour = "lightblue")
writeData(wb, "EG", cons)

headerStyle <- createStyle(fgFill = "#76933C", fontColour = "white", 
                           textDecoration = "bold")
thicklineStyle <- createStyle(borderStyle = "thick", 
                              borderColour = "#76933C", border = "Top")

conditionalFormatting(wb, "EG", cols = 1:ncol(cons), 
                      rows = 2:nrow(cons),
                      rule = "$L2 == 1", style = thicklineStyle)
# Hide the cell entry if it the same as the preceeding (and a different parent)
#conditionalFormatting(wb, "EG", 
#                      cols = c(4:12), rows = 2:nrow(cons), 
#                      rule = "AND(D2 == D1, $D2 == $D1)", style = createStyle(fontColour = "white"))
# Need this twice as I can't make it skip over the Plant Name column
#conditionalFormatting(wb, "EG", 
#                      cols = c(15:ncol(cons)), rows = 2:nrow(cons), 
#                      rule = "AND(O2 == O1, $D2 == $D1)", style = createStyle(fontColour = "white"))
#conditionalFormatting(wb, "EG", 
#                      cols = 9, rows = 2:nrow(cons),
#                      rule = "$I2 <= $D2/2", style = createStyle(bgFill = "orange"))
#addStyle(wb, "EG", nbrStyle, cols = c(8, 10, 14, 16), rows = 2:nrow(cons), 
#         gridExpand = TRUE)
addStyle(wb, "EG", headerStyle, cols = 1:ncol(cons),
         rows = 1)
#addStyle(wb, "EG", createStyle(wrapText = TRUE), 
#         cols = c(4,8:12, 14:17), rows = 1, stack = TRUE)
#addStyle(wb, "EG", createStyle(halign = "center"),
#         cols = c(6, 7), rows = 1:nrow(cons), gridExpand = TRUE, stack = TRUE)
setColWidths(wb, "EG", cols = 4:5, widths = "auto")
setColWidths(wb, "EG", cols = c(11:12), 
             hidden = c(T,T))
freezePane(wb, "EG", firstActiveRow = 2)

saveWorkbook(wb, file = "EG Roll Widths.xlsx", overwrite = TRUE)

rm(headerStyle, thicklineStyle, wb)

