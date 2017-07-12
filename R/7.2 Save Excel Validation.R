#------------------------------------------------------------------------------
#
# Build the Excel output files for valiations
#
# The QlikView Asset Utilization shows the data for the die and board consumed
# by plant for 2016 and shows the work center, but does not give the actual 
# grade
#
# The QlikView Fiber Analysis shows the material description of the
# board consumed 
#

# The model parameters and variables are set in 7.1

# First run the 7.1 Opt Solution Post Process.R script
if (!exists("opt_detail")) source("R/7.1 Opt Solution Post Process.R")

# Use rawDataAll from the Read Data.R script
# If we don't have it, read it from the csv file
if (!exists("rawDataAll")) {
  rawDataAll <- read_csv(file.path(proj_root, paste("Data/All Shipments Fixed ", 
                               mYear, ".csv", sep = "")))
}

# Summarize by plant
rawData_plant <- rawDataAll %>%
  group_by(Year, Mill, Plant, Facility.Name, Fac.Abrev, 
           Grade, CalDW, Width) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

if (projscenario == "Europe"){
  rawData_plant <- rawData_plant %>%
    filter(str_sub(Plant, 1, 5) == "PLTEU")
}

rawData_plant$Year <- as.character(rawData_plant$Year)

#rawData_plant <- rawData_plant %>%
#  separate(CalDW, c("Cal", "Dia", "Wind"), remove = FALSE)

rawData_plant <- rawData_plant %>%
  group_by(Year, Mill, Grade, CalDW, Width) %>%
  mutate(Prod.Plant.Count = n_distinct(Plant))

rawData_plant <- rawData_plant %>%
  rename(Plant.Name = Facility.Name,
         Plant.Abrev = Fac.Abrev)

# Calculate the percent of total tons by plant
rawData_plant <- rawData_plant %>%
  group_by(Year, Mill, CalDW, Width) %>%
  mutate(Plant.Pct = Tons/sum(Tons))

# Join the optimal results table with the raw data by plant
opt_detail_plants <- 
  left_join(opt_detail, rawData_plant,
            by = c("Fac" = "Mill",
                   "Grade" = "Grade",
                   "CalDW" = "CalDW",
                   "Prod.Width" = "Width"))

opt_detail_plants <- opt_detail_plants %>%
  mutate(Sim.Parent.Avg.OH.DOH = 365*Sim.Parent.Avg.OH/Parent.DMD) 

# Rearrange the columns
opt_detail_plants <- opt_detail_plants %>%
  select(Year, Plant, Plant, Mill = Fac, Grade, CalDW, Parent.Width, 
         Prod.Width, Prod.Plant.Count, Prod.Trim.Width, Prod.Trim.Tons, 
         Tons, Parent.Policy, everything())

# Need to round the board width down to match the data from QlikView
# ClikView now is out to 4 places. 
opt_detail_plants <- opt_detail_plants %>%
  mutate(Prod.Width3 = round(Prod.Width, 3),
         Prod.Gross.Tons = round(Prod.Gross.Tons, 2),
         Parent.DMD = round(Parent.DMD, 2),
         Prod.Trim.Tons = round(Prod.Trim.Tons, 2),
         Parent.Width = round(Parent.Width, 4),
         Prod.Trim.Width = round(Prod.Trim.Width, 4),
         Prod.Width = round(Prod.Width, 4))

# Count the number of plants that use the parent
opt_detail_plants <- opt_detail_plants %>%
  group_by(Mill, Grade, CalDW, Parent.Width) %>%
  mutate(Parent.Plant.Count = n_distinct(Plant)) %>%
  ungroup()

# Compute the trim loss for each product by the plant in which it runs
opt_detail_plants <- opt_detail_plants %>%
  mutate(Prod.Trim.Plant.Tons = round(Prod.Trim.Tons * Plant.Pct, 2))

# Create a temp field for the ranking of products
opt_detail_plants <- opt_detail_plants %>%
  ungroup() %>%
  mutate(temp = paste(Grade, Caliper, sprintf("%.4f", Prod.Width), 
                      Dia, Wind, Mill, remove = F)) %>%
  group_by(Year, Plant) %>%
  arrange(Year, Plant, temp) %>%
  mutate(`Board Index` = dense_rank(temp),
         Tons = round(Tons, 1),
         Prod.Plant.Count = as.integer(Prod.Plant.Count)) %>%
  ungroup()

# Use mm for Europe dimensions
if (projscenario == "Europe") {
  opt_detail_plants <- opt_detail_plants %>%
    mutate(Prod.Width = round(Prod.Width * 25.4, 0),
           Parent.Width = round(Parent.Width * 25.4, 0),
           Prod.Trim.Width = round(Prod.Trim.Width * 25.4, 0))
}

# Read the Die/Sheet info
opt_files <- file.path("G:", "000_Supply Chain", "_Optimization", 
                       "Model Optimization", "Projects", "2017 Projects",
                       "Roll Width Consolidation")

die_sheet <- read.xlsx(file.path(opt_files, 
                                 "Board Consumption - 2016 All Sites.xlsx"),
                       startRow = 3)

# QlikView Consumption by Die Report on the Consumption tab on the 
#   Fiber Analysis report.  See the SUSConsumption Bookmark
#fpinput <- "Data"
#die_roll <- read.xlsx(file.path(fpinput, "QV SUS Roll Consumption 2015-2017.xlsx"))

# Just keep the SUS rows
die_sheet <- die_sheet %>%
  filter(Board.Type.Group == "SUS")

die_sheet <- die_sheet %>%
  select(-Die) %>%
  mutate(Die = Die.Fix)

die_sheet$TON <- round(die_sheet$TON, 1)

# Need to get the full grade out of the Material.Description field for the "AK"
# records.  If OM or PK, just add "XX".  OTherwise use the Short field
die_sheet <- die_sheet %>%
  mutate(Grade = ifelse(
    Board.Type.Short %in% c("PK", "OM"),
    paste(Board.Type.Short, "XX", sep = ""),
    ifelse(
      Board.Type.Short == "AK",
      ifelse(
        str_sub(Material.Description, 1, 1) == "I",
        str_sub(Material.Description, 5, 8),   # Imperial
        str_sub(Material.Description, 6, 9)    # Metric
      ), Board.Type.Short)
  ))

die_sheet$Board.Caliper <- as.character(die_sheet$Board.Caliper)

die_sheet <- die_sheet %>%
  mutate(Caliper = ifelse(Board.Caliper == "0",
                          ifelse(str_sub(Material.Description, 1, 1) == "I",
                                 str_sub(Material.Description, 3, 4),
                                 str_sub(Material.Description, 4, 5)),
                          Board.Caliper))


die_sheet <- die_sheet %>%
  mutate(Width = 
           ifelse(Board.Width == "0",
                  ifelse(str_sub(Material.Description, 1, 1) == "I",
                         as.numeric(str_sub(Material.Description, 10, 17)),
                         as.numeric(str_sub(Material.Description, 11, 14))),
                  Board.Width))

die_sheet <- die_sheet %>%
  mutate(Dia = 
           ifelse(Board.Type.Group == "SUS", 
                  ifelse(str_sub(Material.Description, 1, 1) %in% c("I", "("),
                         str_sub(Material.Description, -7, -5),
                         str_sub(Material.Description, -10, -6)), ""))

die_sheet$Width <- as.numeric(die_sheet$Width)
die_sheet$Year <- as.character(die_sheet$Year)

die_sheet <- die_sheet %>%
  mutate(Grade = ifelse(Grade == "SUSFC02", "FC02", Grade))


die_sheet <- die_sheet %>%
  mutate(Wind = 
           str_sub(Material.Description, str_length(Material.Description), -1))

die_sheet <- die_sheet %>%
  mutate(Plant = paste("PLT0", Plant, sep = ""))

die_sheet <- die_sheet %>%
  group_by(Year, Plant, PlantName, Board.Type.Group, MRPC2.Text, Die, 
           `Sheet./.Web`, Die.Description, Width, Grade, Caliper,
           Dia, Wind) %>%
  summarize(TON = sum(TON)) %>%
  ungroup()

# Create separate tons columns for each diameter
die_sheet <- die_sheet %>%
  spread(Dia, TON)

# Join them together
opt_detail_plants_die <- left_join(opt_detail_plants, die_sheet,
                                   by = c("Year" = "Year",
                                          "Plant" = "Plant",
                                          "Grade" = "Grade",
                                          "Caliper" = "Caliper",
                                          "Wind" = "Wind",
                                          "Prod.Width3" = "Width"))

# Replace the NAs in the Die column with "."
opt_detail_plants_die$Die[is.na(opt_detail_plants_die$Die)] <- "."

# Get rid of Piscataway
opt_detail_plants_die <- opt_detail_plants_die %>%
  filter(Plant != "PLT00019")

# Add another index to identify multiple dies for the same board
opt_detail_plants_die <- opt_detail_plants_die %>%
  group_by(Year, Plant, `Board Index`) %>%
  mutate(Die.Index = dense_rank(Die)) %>%
  ungroup()

opt_detail_plants_die <- opt_detail_plants_die %>%
  arrange(Year, Plant, `Board Index`, Die.Index)

opt_detail_plants_die <- opt_detail_plants_die %>%
  mutate(Desired.Board = " ",
         OK = " ") %>%
  select(Year, 
         Plant,
         `Plant Name` = Plant.Name, 
         Plant.Abrev,
         `Board Index`,
         Grade, 
         Caliper, 
         `Actual Width` = Prod.Width, 
         Dia, Wind,
         Machine = Mill,
         Parent.Policy,
         `Desired Board Index` = Desired.Board,
         #`Total Board Tons` = Prod.Gross.Tons, 
         `Parent Width` = Parent.Width, 
         `Trim Width` = Prod.Trim.Width, 
         `Board Tons` = Tons,
         `Trim Tons` = Prod.Trim.Plant.Tons,
         `Board Plant Count` = Prod.Plant.Count,
         `Parent Plant Count` = Parent.Plant.Count,
         `Parent Total Tons` = Parent.DMD,
         `Enter "N" if not OK` = OK,
         `Press Format` = `Sheet./.Web`, 
         Die, 
         #`Die Tons` = TON,
         `Die Description` = Die.Description, 
         MRPC = MRPC2.Text,
         Die.Index,
         everything())

opt_detail_plants_die <- opt_detail_plants_die %>%
  select(-(CalDW:Board.Type.Group))

# Clear the data for duplicate rows where the Die.Index > 1
opt_detail_plants_die[opt_detail_plants_die$Die.Index > 1, 6:20] <- ""

# Make these numeric so we can format them as numers
opt_detail_plants_die <- opt_detail_plants_die %>%
  mutate(`Board Tons` =  as.numeric(`Board Tons`),
         `Trim Tons` = as.numeric(`Trim Tons`),
         `Trim Width` = as.numeric(`Trim Width`),
         `Parent Total Tons` = as.numeric(`Parent Total Tons`))

#copy.table(opt_detail_plants_die)

#opt_detail_plants_die <- filter(opt_detail_plants_die, Plant == "PLT00003")

g <- opt_detail_plants_die$Plant
# odpd is a shortcut for opt_detail_plants_die
odpd_sp <- split(opt_detail_plants_die, g) 

# Delete the empty diameter columns in each dataframe.  Note that this
# deletes all empty columns, so might mess up the column sequences
# for some situations.
# Here is what I got off the web: df <- df[,colSums(is.na(df))<nrow(df)]
# And I tested one list with the following:
#   odpd_sp[[1]] <- 
#       odpd_sp[[1]][, colSums(is.na(odpd_sp[[1]])) < nrow(odpd_sp[[1]])]
odpd_sp <- lapply(odpd_sp, function(x) 
  x[, colSums(is.na(x)) < nrow(x)])

# Function that goes with lapply to add worksheets for each plant
add_plant_sheet <- function(plant_df) {
  mysheet <- unique(paste(plant_df$Plant, plant_df$Plant.Abrev, sep = "-"))
  # Delete the Plant.Abrev column.  We need to subtract 1 from the
  # column number viariables calculated outside the function
  plant_df <- select_(plant_df, quote(-Plant.Abrev))
  mtabColor <- "white"
  if (str_sub(mysheet, 1, 8) %in% c("PLT00070", "PLT00068")) mtabColor <- "blue"
  addWorksheet(wb, mysheet, tabColour = mtabColor)
  writeData(wb, mysheet, plant_df, startRow = 3, withFilter = TRUE)
  setColWidths(wb, mysheet, cols = mcols, widths = "auto")
  nbr_rows <- nrow(plant_df)
  nbr_cols <- ncol(plant_df)
  # Alternate shading of rows
  conditionalFormatting(wb, mysheet, cols = 1:nbr_cols, 
                        rows = 1:(nbr_rows + 3),
                        rule = "ISODD($D1)", style = rowFillStyle)
  # Thick line between parents
  conditionalFormatting(wb, mysheet, cols = 1:nbr_cols, 
                        rows = 1:(nbr_rows + 3),
                        rule = "$Y1 == 1", style = thicklineStyle)
  conditionalFormatting(wb, mysheet, cols = trimWidthCol - 1, 
                        rows = 4:(nbr_rows + 3), rule = "$N4 > 0.0", 
                        style = trimWidthStyle)
  writeFormula(wb, mysheet, paste("SUM(O4:O",nbr_rows+3,")", sep = ""),
               startRow = 1, startCol = 15)
  writeFormula(wb, mysheet, paste("SUM(P4:P",nbr_rows+3,")", sep = ""),
               startRow = 1, startCol = 16)
  writeFormula(wb, mysheet, paste("P1/O1"), startRow = 1, startCol = 17)
  addStyle(wb, mysheet, nbrStyle, cols = c(15, 16, 19), 
           rows = c(1:(nbr_rows+3)), 
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb, mysheet, headerStyle, cols = 1:nbr_cols, rows = 3, stack = TRUE)
  addStyle(wb, mysheet, leftAlignStyle, cols = trimWidthCol, 
           rows = c(4:(nbr_rows+3)), stack = TRUE)
  addStyle(wb, mysheet, rightAlignStyle, cols = 15:19, 
           rows = c(4:(nbr_rows+3)), gridExpand = TRUE, stack = TRUE)
  addStyle(wb, mysheet, pctStyle, rows = 1, cols = 17, stack = TRUE)
  addStyle(wb, mysheet, createStyle(wrapText = TRUE), cols = 1:24, 
           rows = 3, stack = TRUE)
  addStyle(wb, mysheet, createStyle(fontColour = "yellow"), cols = c(12, 20),
           rows = 3, stack = TRUE)
  # the column where the use enters "N". fontColor doesn't work?
  addStyle(wb, mysheet, createStyle(fontColour = "red", halign = "center",
                                    textDecoration = "bold"), 
           rows = c(4:(nbr_rows + 3)), cols = 20, stack = TRUE)
  addStyle(wb, mysheet, createStyle(halign = "center"),
           rows = c(4:(nbr_rows + 3)), cols = 4, stack = TRUE)
  setColWidths(wb, mysheet, cols = die_idx_pos - 1, hidden = T)
  freezePane(wb, mysheet, firstActiveRow = 4)
}

#df_irv <- opt_detail_plants_die %>%
#  filter(`Plant Name` == "IRVINE")

wb <- createWorkbook()

mcols <- match(c("Plant Name", "Die Description", "MRPC"), 
               names(opt_detail_plants_die))
die_idx_pos <- match(c("Die.Index"), names(opt_detail_plants_die))
msumcols <- match(c("Plant Board Tons", "Trim Tons"), 
                  names(opt_detail_plants_die))
trimWidthCol <- match(c("Trim Width"), names(opt_detail_plants_die))

nbrStyle <- createStyle(numFmt = "#,##0.0", halign = "right")
pctStyle <- createStyle(numFmt = "0.0%", halign = "right")
rowFillStyle <- createStyle(bgFill = "#EBF1DE")
trimWidthStyle <- createStyle(fontColour = "red")
leftAlignStyle <- createStyle(halign = "left")
rightAlignStyle <- createStyle(halign = "right")
headerStyle <- createStyle(fgFill = "#76933C", fontColour = "white", 
                           textDecoration = "bold")
thicklineStyle <- createStyle(borderStyle = "thick", 
                              borderColour = "#76933C", border = "Top")

addWorksheet(wb, "Parent Rolls", tabColour = "lightblue")
odp_x <- opt_detail_plants %>%
  select(Year, Grade, Caliper, `Parent Width` = Parent.Width, Dia, Wind, Mill,
         `Parent Tons` = Parent.DMD, `Prod Width` = Prod.Width, 
         `Prod Tons` = Prod.Gross.Tons, `Trim Width` = Prod.Trim.Width, 
         `Trim Tons` = Prod.Trim.Tons, `Plant Name` = Plant.Name, 
         `Plant Tons` = Tons,
         `Parent VMI` = Parent.Bev.VMI, `Prod Ship Count` = Prod.DMD.Count,
         `Parent DOH` = Sim.Parent.Avg.OH.DOH, Parent.Policy,
         everything()) %>%
  arrange(Year, Grade, Caliper, `Parent Width`, Dia, Wind, `Parent Tons`, 
          `Prod Width`)

writeData(wb, "Parent Rolls", odp_x)

conditionalFormatting(wb, "Parent Rolls", cols = 1:ncol(odp_x), 
                      rows = 2:nrow(odp_x),
                      rule = "$D2 != $D1", style = thicklineStyle)
# Hide the cell entry if it the same as the preceeding (and a different parent)
conditionalFormatting(wb, "Parent Rolls", 
                      cols = c(4:12), rows = 2:nrow(odp_x), 
                      rule = "AND(D2 == D1, $D2 == $D1)", style = 
                        createStyle(fontColour = "white"))
# Need this twice as I can't make it skip over the Plant Name column
conditionalFormatting(wb, "Parent Rolls", 
                      cols = c(15:ncol(odp_x)), rows = 2:nrow(odp_x), 
                      rule = "AND(O2 == O1, $D2 == $D1)", style = 
                        createStyle(fontColour = "white"))
conditionalFormatting(wb, "Parent Rolls", 
                      cols = 9, rows = 2:nrow(odp_x),
                      rule = "$I2 <= $D2/2", 
                      style = createStyle(bgFill = "orange"))
addStyle(wb, "Parent Rolls", nbrStyle, cols = c(8, 10, 14, 16), 
         rows = 2:nrow(odp_x), 
         gridExpand = TRUE)
addStyle(wb, "Parent Rolls", headerStyle, cols = 1:ncol(odp_x),
         rows = 1)
addStyle(wb, "Parent Rolls", createStyle(wrapText = TRUE), 
         cols = c(4,8:12, 14:17), rows = 1, stack = TRUE)
addStyle(wb, "Parent Rolls", createStyle(halign = "center"),
         cols = c(6, 7), rows = 1:nrow(odp_x), gridExpand = TRUE, stack = TRUE)
setColWidths(wb, "Parent Rolls", cols = 13, widths = "auto")
freezePane(wb, "Parent Rolls", firstActiveRow = 2)

instr <- data_frame(Item = 1, 
                    Note = "Note that the column widths to the right are not by machine")
#instr <- rbind(instr, c(2, 
#                        "Diameters may show that are wider than what came off any specific machine"))

#addWorksheet(wb, "Notes")
#writeData(wb, "Notes", instr)


# Add the plants
#l_ply(odpd_sp, add_plant_sheet)
walk(odpd_sp, add_plant_sheet)

saveWorkbook(wb, file = paste0("ParentValidation", projscenario, mYear, ".xlsx"), overwrite = TRUE)

rm(instr)

#------------------------------------------------------------------------------
#
# Create a summary by roll spread by diameters 
#

rd_2016 <- read_csv(file.path(proj_root, paste0("Data/All Shipments Fixed ", 
                          "2016", ".csv")))

rd_2016 <- rd_2016 %>%
  separate(CalDW, c("Cal", "Dia", "Wind"), sep = "-", remove = FALSE)


if (projscenario == "Europe") {
  rd_2016 <- rd_2016 %>%
    filter(str_sub(Plant, 1, 5) == "PLTEU")
}

# Just keep the items in the optimal solution 
rd_opt <- 
  semi_join(rd_2016, opt_detail,
            by = c("Mill" = "Fac",
                   "Grade" = "Grade",
                   "Cal" = "Caliper",
                   "Width" = "Prod.Width",
                   "Wind" = "Wind"))

rd_opt %>% summarize(Tons = sum(Tons))

# Add counts for number of Diam, Winds & Plants
rd_opt <- rd_opt %>%
  group_by(Mill, Grade, Cal, Width) %>%
  mutate(Nbr.Dia = n_distinct(Dia),
         Nbr.Wind = n_distinct(Wind),
         Nbr.Plants = n_distinct(Plant)) %>%
  ungroup()

# Summarize Tons by Year, Diam & Wind
rd_opt <- rd_opt %>%
  group_by(Year, Mill, Grade, Cal, Width, Dia, Wind, Nbr.Dia, 
           Nbr.Wind, Nbr.Plants) %>%
  summarize(Tons = round(sum(Tons),0)) %>%
  mutate(Tons = format(Tons, big.mark = ",")) %>%
  ungroup()

# Spread by Year, Dia
rd_opt <- rd_opt %>%
  unite(YD, Year, Dia) %>%
  spread(YD, Tons, fill = "", convert = TRUE)

rd_opt <- rd_opt %>%
  arrange(Mill, Grade, Cal, Width)

rd_opt[is.na(rd_opt)] <- ""

# join to opt solution to get the parent and trim info
# The opt_detail comes from the 7.1 Opt Solution Post Process.R file
opt_detail_m <- opt_detail %>%
  filter(Fac %in% facility_list)

opt_detail_m <- opt_detail_m %>%
  select(Fac, Grade, Caliper, Wind, Prod.Width, Parent.Width, Prod.Trim.Width, 
         Prod.Trim.Tons, Parent.Policy) 

opt_list <- left_join(rd_opt, opt_detail_m,
                      by = c("Mill" = "Fac",
                             "Grade" = "Grade",
                             "Cal" = "Caliper",
                             "Width" = "Prod.Width",
                             "Wind" = "Wind"))

opt_list <- opt_list %>%
  select(Mill, Grade, Cal, Width, Wind, Parent.Width,
         Prod.Trim.Width, Prod.Trim.Tons, everything())


if (projscenario == "Europe") {
  opt_list <- opt_list %>%
    mutate(Width = round(Width * 25.4, 0),
           Parent.Width = round(Parent.Width * 25.4, 0),
           Prod.Trim.Width.Round = round(Prod.Trim.Width * 25.4, 0))
} else {
  opt_list <- opt_list %>%
    mutate(Prod.Trim.Tons = round(Prod.Trim.Tons, 1),
           Prod.Trim.Width.Round = round(Prod.Trim.Width*8, 0)/8)
}

# Generate plot of trim width
ggplot(data = opt_list, aes(x = Prod.Trim.Width.Round,
                            y = Prod.Trim.Tons)) +
  geom_histogram(stat = "identity", geom = "bar",
                 aes(fill = Parent.Policy)) +
  ggtitle("Roll Consolidation Trim Loss by With") +
  scale_x_continuous(name = "Trim Loss Width (in.)") +
  scale_y_continuous(name = "Total Trim Tons") +
  theme(legend.key.size = unit(12, "points"))

# Eliminate the "." in the field names for output purposes
opt_list_out <- opt_list
names(opt_list_out) <- gsub("\\.", " ", names(opt_list_out))
names(opt_list_out) <- gsub("_", " ", names(opt_list_out))

names(opt_list_out) <- gsub("2016 ", "", names(opt_list_out))

copy.table(opt_list_out)

nbr_rows <- nrow(opt_list_out)
nbr_cols <- ncol(opt_list_out)

wb <- createWorkbook()
addWorksheet(wb, "Roll Consolidations", tabColour = "lightgreen")
# Format the header row and write the data
addStyle(wb, "Roll Consolidations", createStyle(wrapText = TRUE,
                                                textDecoration = "bold",
                                                border = "Bottom",
                                                borderStyle = "thick"), 
         cols = 1:nbr_cols, rows = 3, stack = TRUE)
writeData(wb, "Roll Consolidations", opt_list_out, startRow = 3, 
          withFilter = TRUE)
addStyle(wb, "Roll Consolidations", createStyle(halign = "right",
                                                numFmt = "NUMBER"),
        rows = 1:nbr_rows + 3, cols = 12:nbr_cols-1, stack = TRUE,
        gridExpand = TRUE)

# Add the titles
writeData(wb, "Roll Consolidations", paste(mYear, projscenario,
                                           "Parent Roll Consolidations"),
          startCol = 1, startRow = 1)
addStyle(wb, "Roll Consolidations", createStyle(fontSize = 18),
         cols = 1, rows = 1, stack = TRUE)

writeData(wb, "Roll Consolidations", paste(mYear, "Diameters"),
          startCol = 12, startRow = 2)
addStyle(wb, "Roll Consolidations", createStyle(textDecoration = "bold"),
         cols = 1:nbr_cols, rows = 1:2, gridExpand = TRUE, stack = TRUE)

# Hide the Nbr Diam and Nbr Wind columns
setColWidths(wb, "Roll Consolidations", cols = c(9:10), 
             hidden = c(T,T))
saveWorkbook(wb, file = paste0(projscenario, mYear, 
                               "ParentSummary.xlsx"), overwrite = TRUE)