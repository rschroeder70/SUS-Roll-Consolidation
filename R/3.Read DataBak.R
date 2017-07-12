# Read consumption Data.
# File must be in the following format:
# Date, MillName, Plant, Grade, Caliper, Width, Diameter, Wind, Tons

print(paste("Reading Roll Shipment Data for ", mYear, sep = ""), quote = F)

# Read the plant codes
plant_codes <- read.xlsx(file.path(proj_root, 
                                   "Data", "GPICustomerPlantCodes.xlsx"))
plant_codes <- plant_codes %>%
  select(Cust, SAP.Code, `Int/Ext/whse`) %>% 
  filter(!is.na(Cust)) %>%
  distinct(Cust, SAP.Code, .keep_all = TRUE)
# Make sure no duplicate Cust codes
if (nrow(plant_codes %>% group_by(Cust) %>%
         mutate(count = n()) %>% 
         filter(count > 1)) > 0) {
  stop("Duplicates in the PlantCodes table")
}

facility_names <- read_csv(file.path(proj_root, "Data", "FacilityNames.csv"))
# Make sure not duplicate Facility.ID
if (nrow(facility_names %>% group_by(Facility.ID) %>%
         mutate(count = n()) %>% 
         filter(count > 1)) > 0) {
  stop("Duplicates in the PlantCodes table")
}

# Reading from Paul's big Excel file is slow so read from the saved
# csv files if the exists
#
fix_data <- FALSE
if (fix_data == TRUE) {
  # This will remove the fixed data file
  file.remove(file.path(fpinput, paste0("All Shipments Fixed ", 
                                    mYear, ".csv")))
}
rm (fix_data)

if(file.exists(file.path(proj_root, "Data", paste0("All Shipments Fixed ", 
                                         mYear, ".csv")))) {
  rawDataAll <- read_csv(file.path(proj_root, 
                                   "Data", paste("All Shipments Fixed ", 
                               mYear, ".csv", sep = "")))  
} else {
  # 
  # Check if we have a No SKU Fix file saved, if we do read that 
  # 
  if (file.exists(file.path(proj_root, "Data", 
                            paste0("All Shipments No SKU Fix ", mYear, ".csv")))) {
    rawDataAll <- read_csv(
      file.path(proj_root, "Data", 
                paste0("All Shipments No SKU Fix ", mYear, ".csv"))) 
    rawDataAll <- rawDataAll %>%
      mutate(Cal = as.character(Cal))
  } else {
    # We need to start from scratch and read the xlsx file    
    rawDataAll <- read.xlsx(file.path(fpinput,consdatafile), 
                            sheet = consdatasheet, 
                            colNames = TRUE,
                            detectDates = TRUE, check.names = TRUE)
    
    rawDataAll <- as_tibble(rawDataAll)
    rawDataAll <- rawDataAll %>%
      select(Ship.to, Sold.to.pt, Description, Batch, MFGPlant, 
             Date = Ac.GI.date, Cal = Cal.lb, Grade, Width = Calc.width, 
             Dia, Machine, Tons = Net.Ton, Grade.Type, I.M)
    
    rawDataAll <- rawDataAll %>% filter(Grade.Type == "SUS")
    rawDataAll <- rawDataAll %>% filter(!Grade %in% c("AKJP", "AK25"))
    rawDataAll <- rawDataAll %>% mutate(Year = mYear)
    rawDataAll <- rawDataAll %>% mutate(I.M = str_trim(I.M))
    
    # Create the Mill field.  If missing a MFGPlant,
    # fill in the MFGPlant data from the first two postitions 
    # of the Batch nbr
    rawDataAll <- rawDataAll %>%
      mutate(Mill = ifelse(is.na(MFGPlant), 
                           paste("00", str_sub(Batch, 1, 2), sep = ""), 
                           MFGPlant)) 
    rawDataAll <- rawDataAll %>% select(-MFGPlant, -Batch)
 
    rawDataAll <- rawDataAll %>%
      group_by(Ship.to, Sold.to.pt, Description,
               Date, Cal, Grade, Width, Dia, Machine,
               Grade.Type, I.M, Year, Mill) %>%
      summarize(Tons = sum(Tons)) %>%
      ungroup()
       
    rawDataAll <- rawDataAll %>%
      mutate(Wind = str_sub(Description, str_length(Description), -1))
    
    rawDataAll <- left_join(rawDataAll, plant_codes,
                            by = c("Ship.to" = "Cust"))
    
    # Exclude the R&D in West Monroe and mills
    rawDataAll <- rawDataAll %>%
      filter(!(SAP.Code %in% c("5538", "5539", "PLT00031", "PLT00033", "PLT00040")))
    
    # Exclude Verst and Piscataway
    rawDataAll <- rawDataAll %>%
      filter(!(SAP.Code %in% c("PLT00508", "PLT00019")))
    
    # If the ship to is a Europe port, then use the sold-to to get the actual
    # location.  Rename the SAP.Code field so it doesn't conflict with what we 
    # already have
    rawDataAll <-
      left_join(rawDataAll, rename(select(plant_codes, Cust, SAP.Code),
                                   SAP.Sold.To = SAP.Code),
                by = c("Sold.to.pt" = "Cust"))
    
    # Where are they different
    rawDataAll %>% filter(SAP.Code != SAP.Sold.To) %>%
      group_by(Ship.to, SAP.Code, Sold.to.pt, SAP.Sold.To) %>%
      summarize(Tons = sum(Tons))
    
    # Update the SAP.Code to the sold where they are not equal (and where the
    # SAP.Sold.To is not NA).  This fixes a
    # Portland issue, too.
    rawDataAll <- rawDataAll %>%
      mutate(SAP.Code = ifelse(!is.na(SAP.Sold.To) & (SAP.Code != SAP.Sold.To), 
                               SAP.Sold.To, SAP.Code))
    
    # Get rid of the Sold.To
    rawDataAll <- rawDataAll %>%
      select(-SAP.Sold.To)
    
    # If we don't have an SAP.Code (Plant), then it should be open market
    # or Japan and we don't want it
    rawDataAll <- rawDataAll %>%
      ungroup() %>%
      filter(!is.na(SAP.Code)) %>%
      rename(Plant = SAP.Code)
    
    # Check the total tons left
    rawDataAll %>% summarize(sum(Tons))
    
    # Figure out what to do with the missing machines (2015 data)
    rawDataAll <- rawDataAll %>%
      mutate(Machine = ifelse(
        is.na(Machine),
        ifelse(
          (Mill == "0031" | Mill == "31"),            # WM
          ifelse(Cal >= 24, "06", "07"),
          ifelse(Cal >= 18, "02", "01")  # Macon
        ),
        Machine
      ))
    
    # Build the Mill-Machine field
    rawDataAll <- rawDataAll %>% 
      mutate(Mill = ifelse((Mill == "0031" | Mill == "31"), "W", "M")) %>%
      mutate(Mill = paste(Mill, as.numeric(Machine), sep = "")) %>%
      select(-Machine)
    
    # How many SKUs are we starting with
    rawDataAll %>% 
      summarize(Prod.Count = n_distinct(Mill, Grade, Cal, Wind, Width, Dia),
                Plants = n_distinct(Plant))
    
    # Save as a csv at this point - prior to diameter consolidation
    write_csv(rawDataAll, 
              paste0("Data/All Shipments No SKU Fix ", mYear, ".csv"))
  }
  
  print("Proceeding to fix SKU data")
  
  tons_check <- TRUE
  if(tons_check) rawDataAll %>% filter(Width %in% c(62.875, 64.0625, 64.25) & Grade == "AKPG") %>%
    group_by(Mill, Cal, Width) %>%
    summarize(Tons = sum(Tons))
  
  # Note that the sequence of operations here is important as the 
  # validation file was created fairly early in the process and we 
  # need to be able to match that data to the rawDataAll data frame.
  
  # Consolidate Diameters.  These first two I did before sending out the
  # Validation file.  They need to run here and in this order.
  # Make anything less than 72 a 72
  rawDataAll <- rawDataAll %>%
    rename(Dia.Orig = Dia) 
  rawDataAll <- rawDataAll %>%
    mutate(Dia = ifelse(Dia.Orig < 72, 72, Dia.Orig))
  
  # Pick the maximum diameter that shipped from
  # each mill
  rawDataAll <- rawDataAll %>% 
    group_by(Year, Mill, Grade, Cal, Wind, Width, I.M) %>%
    mutate(Max.Diam = max(Dia)) %>%
    mutate(Dia = Max.Diam) %>%
    ungroup()
  
  # Look for low volume SKUS made at more than 1 mill
  mill_fix <- rawDataAll %>%
    select(Year, Mill, Grade, Cal, Dia, Width, Wind, Tons, Plant) %>%
    group_by(Year, Mill, Grade, Cal, Dia, Wind, Width) %>%
    summarize(Mill.Tons = sum(Tons),
              Plant.Count = n_distinct(Plant)) %>%
    ungroup() %>%
    group_by(Year, Grade, Cal, Wind, Width) %>%
    mutate(Nbr.Mills = n(),
           Total.Tons = sum(Mill.Tons)) %>%
    ungroup() %>%
    mutate(Mill.Pct = Mill.Tons/Total.Tons) %>% 
    filter(Nbr.Mills > 1 & Mill.Tons <= 5 & Mill.Pct < .05 & Total.Tons > 100) %>%
    rename(Old.Mill = Mill)
  
  # Join this back to the main data file to show all mills where these 
  # items were made
  mill_fix <- 
    left_join(mill_fix,
              summarize(group_by(
                select(rawDataAll, Year, Mill, Grade, Cal, Dia, Wind, Width, Tons), 
                Year, Mill, Grade, Cal, Dia, Wind, Width), Tons = sum(Tons)),
              by = c("Year", "Grade", "Cal", "Dia", "Wind", "Width")) %>%
    arrange(Year, Grade, Cal, Dia, Wind, Width, desc(Tons))
  
  # Pick the row with the max tons
  mill_fix <- mill_fix %>%
    group_by(Year, Grade, Cal, Dia, Wind, Width, Old.Mill) %>%
    slice(which.max(Tons)) %>%
    rename(New.Mill = Mill)
  
  rawDataAll <- 
    left_join(rawDataAll, 
              select(mill_fix, Year, Grade, Cal, Dia, Wind, 
                     Width, Old.Mill, New.Mill),
              by = c("Year", "Grade", "Cal", "Dia", "Wind", 
                     "Width", "Mill" = "Old.Mill")) %>%
    mutate(Mill = ifelse(is.na(New.Mill), Mill, New.Mill)) %>%
    select(-New.Mill)

  if(tons_check) rawDataAll %>% filter(Width %in% c(62.875, 64.0625, 64.25) & Grade == "AKPG") %>%
    group_by(Mill, Grade, Cal, Width) %>%
    summarize(Tons = sum(Tons))
     
  # For PLT00070, convert all AKPG 21 pt 62.875 to 24 pt 64.25 
  # 21 pt 64.0625 to 24 pt 62.875
  dt <- data.table(rawDataAll)
  dt[(Grade == "AKPG" & Plant == "PLT00070" & Cal == "21" & Width == 62.875), 
     `:=` (Cal = "24", Width = 64.25)]
  dt[(Grade == "AKPG" & Plant == "PLT00070" & Cal == "21" & Width == 64.0625), 
     `:=` (Cal = "24", Width = 62.875)]
  
  if(tons_check) dt %>% filter(Width %in% c(62.875, 64.0625, 64.25) & Grade == "AKPG") %>%
    group_by(Mill, Grade, Cal, Dia, Width) %>%
    summarize(Tons = sum(Tons))
  
  
  # The following per Pat Averitte
  dt[(Grade == "AKPG" & Cal == "24" & Width == 62.8125), 
     `:=` (Width = 62.875)]
  dt[(Grade == "AKPG" & Cal == "27" & Width == 47.0625), 
     `:=` (Width = 47.125)]
  
  # It looks like the 27 pt. AKPG 47.5 rolls converted to 24 lb in 
  # 2016.  Make them all 24, regardless OD
  dt[(Grade == "AKPG" & Cal == "27" & Width == 47.5),
     `:=` (Cal = "24")]

  # Pure Leaf change @ Perry - 3 up on the 56.25"
  dt[Grade == "AKPG" & Cal == "21" & Width == 41.75, 
     Width := 56.25]
  
  # Shiner @ Perry change
  dt[Grade == "AKPG" & Cal == "24" & Width == 37.5,
     Width := 39.375]
  
  # Change the wind on the SNEEK grades to "W"
  dt[Grade == "OMSN", Wind := "W"]
  
    # Make the 2015 AKPG M1 18-72-F 37.875 into 18-77-F to match 2016
  dt[(Grade == "AKOS" & Mill == "M1" & Cal == "18" & Dia == 72 & 
        Width == 37.875 & Year == "2015" & Wind == "F"), 
     `:=` (Dia = 77)]

  # Make the 38.125 wide AKPG 21 pt. that are 81 OD equal to 84
  dt[(Grade == "AKPG" & Mill == "W6" & Cal == "21" & Dia == 81 & 
        Width == 38.125 & Year == "2015" & Wind == "W"), 
     `:=` (Dia = 84)]
  
  # Others that I found just looking at the data....
  dt[(Grade == "AKPG" & Mill == "W6" & Cal == "21" & Dia == 84 & 
        Width == 42.5 & Year == "2015" & Wind == "W"), 
     `:=` (Dia = 77)]

  dt[(Grade == "AKPG" & Mill == "W6" & Cal == "21" & Dia == 77 & 
        Width == 56.25 & Year == "2015" & Wind == "W"), 
     `:=` (Dia = 84)]
  
  dt[(Grade == "AKPG" & Mill == "W6" & Cal == "18" & Dia == 77 & 
        Width == 45.25 & Year == "2015" & Wind == "W"), 
     `:=` (Dia = 81)]
  
  dt[(Grade == "AKOS" & Mill == "W7" & Cal == "24" & Dia == 72 & 
        Width == 39.5669 & Year == "2015" & Wind == "W"), 
     `:=` (Mill = "W6", Grade = "AKHS")]
  
  dt[(Grade == "AKPG" & Mill == "W6" & Cal == "24" & Dia == 81 & 
        Width == 31 & Year == "2015" & Wind == "W"), 
     `:=` (Dia = 84)]
  
  dt[(Grade == "AKOS" & Mill == "W6" & Cal == "24" & Dia == 72 & 
        Width == 42.875 & Year == "2015" & Wind == "W"), 
     `:=` (Dia = 77)]
 
  dt[(Grade == "AKPG" & Mill == "M2" & Cal == "18" & Dia == 72 & 
        Width == 45.25 & Year == "2015" & Wind == "W"), 
     `:=` (Dia = 81)]
  
  rawDataAll <- as_data_frame(dt)
  rm(dt)
  
  if(tons_check) rawDataAll %>% filter(Width %in% c(62.875, 64.0625, 64.25) & Grade == "AKPG") %>%
    group_by(Mill, Cal, Dia, Width) %>%
    summarize(Tons = sum(Tons))
  
  
#  browser()

  # Delete these items:
  rawDataAll <- rawDataAll %>%
    filter(!(Mill == "W6" & Cal == "24" & Grade == "AKPG" & Width == 31.75) &
    #         !(Mill == "W6" & Cal == "21" & Grade == "AKPG" & Width == 42.75) &
             !(Mill == "W6" & Cal == "24" & Grade == "AKOS" & Width == 37.875) &
             !(Mill == "W6" & Cal == "28" & Grade == "PKXX" & Width == 43.25))
  
  
  # Update to AKHS for the full year
  akhs_to_fix <- 
    tribble( ~Mill, ~Plant, ~Grade, ~Cal, ~Dia, ~Wind, ~Width, ~New.Grade, ~New.Mill,
             #----------------------------------------
             "W6",  "PLTEUMFR", "AKOS", "24", 72, "W",  43.5433, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"24", 72, "W",	39.5669, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	38.5433, "AKOS", "M2",
             "W6",	"PLTEUBUK",	"AKOS",	"26", 72, "W",  44.5669, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	44.2913, "AKHS", "W6",
             "M2",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	44.2913, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	51.0236, "AKHS", "W6",
             "M2",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	51.0236, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	37.0079, "AKHS", "W6",
             "M2",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	37.0079, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	40.5906, "AKHS", "W6",
             "M2",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	40.5906, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	37.9134, "AKHS", "W6",
             "M2",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	37.9134, "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	39.685,  "AKHS", "W6",
             "W6",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	39.685,  "AKHS", "W6",
             "M2",	"PLTEUBUK",	"AKOS",	"22", 72, "W",	48.7402, "AKHS", "W6",
             "M2",  "PLTEUBUK", "AKOS", "22", 72, "W",  48.7402, "AKHS", "W6")
  
  rawDataAll <- left_join(rawDataAll, akhs_to_fix,
                          by = c("Mill" = "Mill",
                                 "Plant" = "Plant",
                                 "Grade" = "Grade",
                                 "Cal" = "Cal",
                                 "Dia" = "Dia",
                                 "Wind" = "Wind",
                                 "Width" = "Width"))
  
  rawDataAll <- rawDataAll %>%
    mutate(Grade = ifelse(!is.na(New.Grade), New.Grade, Grade),
           Mill = ifelse(!is.na(New.Mill), New.Mill, Mill)) %>%
    select(-New.Grade, -New.Mill)
  
  rm(akhs_to_fix)

  # Fix substitutions, trials and obsoletes
  # This file was generated before we made the max diameter fix
  # Substitutions have a Desired.Board.Index > 0.  Items to delete
  # have a Desired.Board.Index = 0
  plant_list <- c("PLT00003-IRV", "PLT00006-STMTN", "PLT00008-CS",
                  "PLT00009-EGVW",  "PLT00015-WMC", 
                  "PLT00017-PAC",  "PLT00020-MAR", "PLT00022-SOL",
                  "PLT00023-VF", "PLT00041-LUM", "PLT00042-FS",
                  "PLT00043-CENT", "PLT00044-MIT", "PLT00047-KVL", 
                  "PLT00048-GORD", "PLT00051-CLT", "PLT00053-LAW",
                  "PLT00054-WAU", "PLT00055-ORO", "PLT00060-TUS",
                  "PLT00068-PER", "PLT00070-WMB", "PLT08001-POR",
                  "PLT08005-VAN", "PLT08007-STAU", "PLT08008-HAM",
                  "PLT08009-NEWT", "PLT08010-WAYNE", "PLT08102-MIS",
                  "PLT08103-WIN")
  
  # Use ldply to get and cobine data from each tab
  skus_to_fix <- 
    ldply(plant_list, 
          function(x) {
            read.xlsx(file.path("DataValidationResponses", 
              #"ParentValidationBase2016 - Pre-run Consolidation Options v3.xlsx"),
             "ParentValidation 2016 Base 03.29.17 RS.xlsx"),
                      sheet = x, startRow = 3)
          }
          , .id = NULL)
  
  # Just save the first row each roll index (Die.Index == 1)
  skus_to_fix <- skus_to_fix %>%
    filter(Die.Index == 1) %>%
    rename(Valid = `Enter."N".if.not.OK`) %>%
    select(Year:Valid) %>%
    mutate(Desired.Board.Index = as.numeric(Desired.Board.Index),
           Dia = as.numeric(Dia),
           Actual.Width = as.numeric(Actual.Width)) 
  
  # Make sure no duplicates
  skus_to_fix <- distinct(skus_to_fix)
  
  # Create a table for the SKUs to substitute
  # Exclude any where the Desired.Board.Index == 0 (items to delete)
  skus_to_sub <- skus_to_fix %>%
    filter(is.na(Desired.Board.Index) | Desired.Board.Index != 0)
    
  # self join with the desired index 
  skus_to_sub <- left_join(skus_to_sub, 
                    select(skus_to_sub, Plant, Board.Index, New.Grade = Grade,
                           New.Caliper = Caliper, New.Width = Actual.Width,
                           New.Dia = Dia, New.Wind = Wind, New.Machine = Machine),
            by = c("Plant" = "Plant", 
                   "Desired.Board.Index" = "Board.Index")) 
  
  # Just keep the substitutes
  skus_to_sub <- skus_to_sub %>%
    filter(!is.na(Desired.Board.Index))
  
  
  skus_to_sub <- skus_to_sub %>%
    select(Plant, Machine, Grade, Caliper, Dia, Wind, Actual.Width,
           New.Grade, New.Caliper, New.Width, New.Dia, New.Wind, 
           New.Machine)

  rawDataAll <- left_join(rawDataAll, skus_to_sub,
                          by = c("Plant" = "Plant",
                                 "Mill" = "Machine",
                                 "Grade" = "Grade",
                                 "Cal" = "Caliper",
                                 "Dia" = "Dia",
                                 "Wind" = "Wind",
                                 "Width" = "Actual.Width"))
  
  rawDataAll <- rawDataAll %>%
    mutate(Mill = ifelse(!is.na(New.Machine), New.Machine, Mill),
           Grade = ifelse(!is.na(New.Grade), New.Grade, Grade),
           Cal = ifelse(!is.na(New.Caliper), New.Caliper, Cal),
           Dia = ifelse(!is.na(New.Dia), New.Dia, Dia),
           Wind = ifelse(!is.na(New.Wind), New.Wind, Wind),
           Width = ifelse(!is.na(New.Width), New.Width, Width))
  
  rawDataAll <- rawDataAll %>%
    select(-starts_with("New"))
  
  rm(skus_to_sub)
      
  # Delete data where the SKU is obsolete or a trial
  skus_to_del <- skus_to_fix %>%
    filter(Desired.Board.Index == 0)
  
  rawDataAll <- 
    anti_join(rawDataAll, skus_to_del,
               by = c("Plant" = "Plant",
                      "Mill" = "Machine",
                      "Grade" = "Grade",
                      "Cal" = "Caliper",
                      "Dia" = "Dia",
                      "Wind" = "Wind",
                      "Width" = "Actual.Width"))
              
  rm(skus_to_del)
  rm(skus_to_fix)
  
  
  if(tons_check) rawDataAll %>% filter(Width %in% c(62.875, 64.0625, 64.25) & Grade == "AKPG") %>%
    group_by(Mill, Cal, Width) %>%
    summarize(Tons = sum(Tons))
  
  # Make Perry and West Monroe minimum 77
  rawDataAll <- rawDataAll %>%
    mutate(Dia = ifelse(Plant %in% c("PLT00070", "PLT00068") &
                          Dia < 77, 77, Dia))

  # Need to run this again to pick up the Perry and WM changes above
  # Pick the maximum diameter that shipped from
  # each mill
  rawDataAll <- rawDataAll %>% 
    group_by(Year, Mill, Grade, Cal, Wind, Width, I.M) %>%
    mutate(Max.Diam = max(Dia)) %>%
    mutate(Dia = Max.Diam) %>%
    ungroup()
  
  # I shouldn't need this as these changes are made by reading in the
  # validation files, but until they are consistent across all plants
  # we still need it
  # 2016 beverage mill/machine fixes
  bev_mill_fix <-
    tribble( ~ Mill, ~Grade, ~Cal, ~Width, ~New.Mill,
             #----------------------------------------
             "M2", 	"AKPG", 	"18", 	25.625, 	"W7",
             "M2", 	"AKPG", 	"18", 	31.125, 	"W7",
             "M2", 	"AKPG", 	"18", 	35.875, 	"W7",
             "M2", 	"AKPG", 	"18", 	36.1875, 	"W7",
             "M2", 	"AKPG", 	"18", 	37.75, 	"W7", 
             "M2", 	"AKPG", 	"18", 	42.75, 	"W7",  # FS has M2 only
             "M2", 	"AKPG", 	"18", 	45.25, 	"W7",
             "W7", 	"AKPG", 	"18", 	46.5, 	"M2",
             "M2", 	"AKOS", 	"21", 	32.625, 	"W6",
             "W7", 	"AKOS", 	"21", 	32.625, 	"W6",
             "M2", 	"AKPG", 	"21", 	38.125, 	"W6",
             #"W7", 	"AKOS", 	"21", 	42.75, 	"W6", # Fix @ FS in Val. File
             "W6",  "AKPG",   "21",   46.625,  "M2",
             "W7", 	"AKOS", 	"24", 	39.125, 	"W6",
             "M2",  "AKPG",   "24",   64.25,  "W6",
             "W6", 	"PKXX", 	"28", 	31.75, 	"W7",
             "W6", 	"PKXX", 	"28", 	37.875, 	"W7",
             "W6", 	"PKXX", 	"28", 	38.625, 	"W7",
             "W7", 	"PKXX", 	"28", 	43, 	"W6")
  
  rawDataAll <- left_join(rawDataAll, bev_mill_fix,
                          by = c("Mill", "Grade", "Cal", "Width"))
  
  rm(bev_mill_fix())
  
  rawDataAll <- rawDataAll %>%
    mutate(Mill = ifelse(!is.na(New.Mill), New.Mill, Mill)) %>%
    select(-New.Mill)
  
  # Do this one more time to consolidate diameters from the adjusted mills
  rawDataAll <- rawDataAll %>% 
    group_by(Year, Mill, Grade, Cal, Wind, Width, I.M) %>%
    mutate(Max.Diam = max(Dia)) %>%
    mutate(Dia = Max.Diam) %>%
    ungroup()

  
  if(tons_check) rawDataAll %>% filter(Width %in% c(62.875, 64.0625, 64.25) & Grade == "AKPG") %>%
    group_by(Mill, Cal, Width) %>%
    summarize(Tons = sum(Tons))
  
    
  rawDataAll <- rawDataAll %>%
    unite(CalDW, Cal, Dia, Wind, sep = "-")

  # Aggregate by date
  rawDataAll <- rawDataAll %>% 
    group_by(Year, Date, Mill, Plant, Grade, CalDW, Width, I.M) %>%
    summarize(Tons = sum(Tons)) %>%
    ungroup()
  
  # How many SKUs are we left with
  rawDataAll %>% 
    summarize(Prod.Count = n_distinct(Mill, Grade, CalDW, Width),
              Plants = n_distinct(Plant))
  
  rawDataAll <- left_join(rawDataAll, facility_names,
                          by = c("Plant" = "Facility.ID"))
  
  # Save as a csv for faster read-in
  write_csv(rawDataAll, 
            paste("Data/All Shipments Fixed ", mYear, ".csv", sep  = ""))
}

# Check the tons and # of SKUs
rawDataAll %>% group_by(Year) %>%
  unite(MGC, Mill, Grade, CalDW, Width) %>%
  summarize(Tons = round(sum(Tons), 0),
            SKU.Count = n_distinct(MGC))

# Determine the return freight from each plant.  Note this includes the $150
# opportunity cost.  This is read in the `User Inputs.R` script, but if not 
# read it in here
if (!exists("trim_freight")) {
  shippingData <- file.path(basepath, "Data", "ShipmentData.xlsx")
  trim_freight <- read.xlsx(shippingData, sheet = "Sheet1")
  trim_freight <- trim_freight %>%
    rename(Cost.Ton = `Cost/Ton`)
}
trim_freight_plant <- rawDataAll %>%
  group_by(Year, Mill, Plant, Grade, CalDW, Width, Facility.Name) %>%
  summarize(Tons = sum(Tons)) %>%
  left_join(trim_freight,
            by = c("Mill" = "Mill",
                   "Plant" = "Plant")) %>%
  ungroup()

# If we don't have a freight cost, make it $300
trim_freight_plant$Cost.Ton[is.na(trim_freight_plant$Cost.Ton)] <- 300

# Calculate the cost per ton for each plant based on the weighted volume 
# of each size roll shipping to each plant
trim_freight_plant <- trim_freight_plant %>%
  group_by(Year, Mill, Grade, CalDW, Width) %>%
  mutate(Freight.Cost = Tons * Cost.Ton) %>%
  summarize(Trim.Frt.Cost = sum(Tons * Cost.Ton)/sum(Tons))
trim_freight_plant$Width <- as.character(trim_freight_plant$Width)

# This next if block is where we restrict the data depending
# on the scenario
#
# The BevOnly scenario description might be restricted to just certiain 
# machines (i.e., `BevOnlyW6` so just
# look at the first part of the scenario description
if (str_sub(projscenario, 1, 7) == "BevOnly") {
  #  Pull in VMI list and match to the rolls (not the diameter)
  vmi_rolls <- read.xlsx("Data/VMI List.xlsx",
                         startRow = 2, cols = c(1:6))
  vmi_rolls[is.na(vmi_rolls)] <- ""
  vmi_rolls <- vmi_rolls %>%
    filter(Status != "D")
  vmi_rolls$Caliper <- as.character(vmi_rolls$Caliper)
  vmi_rolls <- vmi_rolls %>%
    select(Caliper, Width, Grade)
  vmi_rolls <- vmi_rolls %>%
    mutate(VMI = 1)
  
  rawDataAll <- rawDataAll %>%
    separate(CalDW, c("Caliper", "Dia", "Wind"), sep = "-",
             remove = FALSE)
  rawDataAll <- semi_join(rawDataAll, vmi_rolls,
                          by = c("Grade" = "Grade",
                                 "Caliper" = "Caliper",
                                 "Width" = "Width"))
  rm (vmi_rolls)
  rawDataAll <- rawDataAll %>%
    select(-Caliper, -Dia, -Wind)

} else if (projscenario == "Europe") {
  rawDataAll <- rawDataAll %>%
    filter(str_sub(Plant, 1, 5) == "PLTEU")
} else if (projscenario == "VMI98") {
  # Beverage VMI + Capri Sun
  #  Pull in VMI list and match to the rolls (not the diameter)
  vmi_rolls <- read.xlsx("Data/VMI List.xlsx",
                         startRow = 2, cols = c(1:6))
  vmi_rolls[is.na(vmi_rolls)] <- ""
  vmi_rolls <- vmi_rolls %>%
    filter(Status != "D")
  vmi_rolls$Caliper <- as.character(vmi_rolls$Caliper)
  vmi_rolls <- vmi_rolls %>%
    select(Caliper, Width, Grade)
  vmi_rolls <- vmi_rolls %>%
    mutate(VMI = 1)
  cpd_vmi <- 
    tribble( ~Caliper, ~Width, ~Grade, ~VMI,
             #----------------------------
             "30",      45.375, "FC02", 1)
             

  vmi_rolls <- bind_rows(vmi_rolls, cpd_vmi)
  
  rawDataAll <- rawDataAll %>%
    separate(CalDW, c("Caliper", "Dia", "Wind"), sep = "-",
             remove = FALSE)
  rawDataAll <- semi_join(rawDataAll, vmi_rolls,
                          by = c("Grade" = "Grade",
                                 "Caliper" = "Caliper",
                                 "Width" = "Width"))
  rawDataAll <- rawDataAll %>%
    select(-Caliper, -Dia, -Wind)
  rm(vmi_rolls, cpd_vmi)
  
} else if (projscenario == "Common1463") {
  # If we don't have the file in memory, read it in
  if (!exists("rd_common")) {
    rd_common <- read_csv("Data/RD Common 15 to 16 1463 SKUs.csv")
  }
  rawDataAll <- 
    semi_join(rawDataAll, rd_common,
              by = c("Mill" = "Mill",
                     "Grade" = "Grade",
                     "CalDW" = "CalDW",
                     "Width" = "Width"))
} else if (str_sub(projscenario, 1, 8) == "NoEurope") {
  rawDataAll <- rawDataAll %>%
    filter(str_sub(Plant, 1, 5) != "PLTEU")
} else if (projscenario == "CPDOnly95") {
  # Use rd_comb
  if (!exists("rd_comb")) source(file.path(proj_root, "R", 
                                           "YoY Data Comp.R"))
  rd_comb$Year <- as.numeric(rd_comb$Year)
  
  rawDataAll <- 
    semi_join(rawDataAll, 
              filter(rd_comb, CPD.Web | CPD.Sheet),
              by = c("Year", "Mill", "Grade", "CalDW", "Width"))
}

# These items are fixed and cannot be cut from a parent any larger 
# than Max.Width
# The logic in the 4a.Sub Parent Roll Optimization.R script uses this table
prod_fixed <- 
  tribble( ~Mill, ~Grade, ~CalDW, ~Width, ~Max.Width,
           #----------------------------
           "W6", "AKPG", "27-84-W", 44.375, 45.0, # Patricia Averitte
           #"W6", "AKPG", "27-84-W", 45.75,  # Patricia Averitte
           "W6", "AKPG", "24-84-W", 37.75, 37.75)  # John Zopf for trim
             
print("Finished Reading Roll Shipment Data", quote = FALSE)