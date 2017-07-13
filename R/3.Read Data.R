#------------------------------------------------------------------------------
#
# Read consumption or shipment Data.
# File results must be in the following format:
# Date, MillName, Plant, Grade, Caliper, Width, Diameter, Wind, Tons
#
# Steps:
# 1. Consolidate the batches by day
# 2. Filter out stuff we don't need
# 3. Make adjustements to the data
#
# This version saves all the original data in fiels with the .Orig suffix
#   After all modifications, these are saved with the .Orig fields to other
#   routines that read raw data can apply the same changes
#

print(paste("Reading Roll Shipment Data for ", mYear, sep = ""), quote = F)

facility_names <- read_csv(file.path(proj_root, "Data", "FacilityNames.csv"))
# Make sure not duplicate Facility.ID
if (nrow(facility_names %>% group_by(Facility.ID) %>%
         mutate(count = n()) %>% 
         filter(count > 1)) > 0) {
  stop("Duplicates in the PlantCodes table")
}

if(file.exists(file.path(proj_root, "Data", paste0("all-shipments-fixed-", 
                                                   mYear, ".csv")))) {
  rawDataAll <- read_csv(file.path(proj_root, 
                                   "Data", paste("all-shipments-fixed-", 
                                                 mYear, ".csv", sep = "")))  
} else {

  print("Proceeding to fix SKU data")
  rawDataAll <- read_csv(paste0("Data/All Shipments No SKU Fix-", mYear, ".csv"))
  rawDataAll$Cal <- as.character(rawDataAll$Cal)

  # Keep all the original info in the "Orig" suffix fields
  
  rawDataAll <- rawDataAll %>%
    rename(Mill.Orig = Mill,
           Grade.Orig = Grade,
           Cal.Orig = Cal,
           Dia.Orig = Dia,
           Wind.Orig = Wind,
           Width.Orig = Width)
  
  rawDataAll <- rawDataAll %>%
    mutate(Mill.Fix = Mill.Orig,
           Grade.Fix = Grade.Orig,
           Cal.Fix = Cal.Orig,
           Dia.Fix = Cal.Orig,
           Wind.Fix = Wind.Orig,
           Width.Fix = Width.Orig)
  
  tons_check <- TRUE
  
  if(tons_check) rawDataAll %>% filter(Width.Orig %in% 
                                         c(62.875, 64.0625, 64.25) & 
                                         Grade.Orig == "AKPG") %>%
    group_by(Mill.Orig, Cal.Orig, Width.Orig) %>%
    summarize(Tons = sum(Tons))
  
  # Note that the sequence of operations here is important as the 
  # validation file was created fairly early in the process and we 
  # need to be able to match that data to the rawDataAll data frame.
  
  # Consolidate Diameters.  These first two I did before sending out the
  # Validation file.  They need to run here and in this order.
  # Make anything less than 72 a 72
  rawDataAll <- rawDataAll %>%
    mutate(Dia.Fix = ifelse(Dia.Orig < 72, 72, Dia.Orig))
  
  # Pick the maximum diameter that shipped from
  # each mill
  rawDataAll <- rawDataAll %>% 
    group_by(Mill.Orig, Grade.Orig, Cal.Orig, Wind.Orig, 
             Width.Orig, I.M) %>%
    mutate(Max.Diam = max(Dia.Fix)) %>%
    mutate(Dia.Fix = Max.Diam) %>%
    ungroup()
  
  # Look for low volume SKUS made at more than 1 mill
  mill_fix <- rawDataAll %>%
    select(Mill.Fix, Grade.Fix, Cal.Fix, Dia.Fix, Width.Fix, 
           Wind.Fix, Tons, Plant) %>%
    group_by(Mill.Fix, Grade.Fix, Cal.Fix, Dia.Fix, Wind.Fix, 
             Width.Fix) %>%
    summarize(Mill.Tons = sum(Tons),
              Plant.Count = n_distinct(Plant)) %>%
    ungroup() %>%
    group_by(Grade.Fix, Cal.Fix, Wind.Fix, Width.Fix) %>%
    mutate(Nbr.Mills = n(),
           Total.Tons = sum(Mill.Tons)) %>%
    ungroup() %>%
    mutate(Mill.Pct = Mill.Tons/Total.Tons) %>% 
    filter(Nbr.Mills > 1 & Mill.Tons <= 5 & Mill.Pct < .05 & Total.Tons > 100) %>%
    rename(Old.Mill = Mill.Fix)
  
  # Join this back to the main data file to show all mills where these 
  # items were made
  mill_fix <- 
    left_join(mill_fix,
              summarize(group_by(
                select(rawDataAll, Mill.Fix, Grade.Fix, Cal.Fix, 
                       Dia.Fix, Wind.Fix, Width.Fix, Tons), 
                Mill.Fix, Grade.Fix, Cal.Fix, Dia.Fix, Wind.Fix, Width.Fix), 
                Tons = sum(Tons)),
              by = c("Grade.Fix", "Cal.Fix", "Dia.Fix", 
                     "Wind.Fix", "Width.Fix")) %>%
    arrange(Grade.Fix, Cal.Fix, Dia.Fix, Wind.Fix, Width.Fix, desc(Tons))
  
  # Pick the row with the max tons
  mill_fix <- mill_fix %>%
    group_by(Grade.Fix, Cal.Fix, Dia.Fix, Wind.Fix, Width.Fix, Old.Mill) %>%
    slice(which.max(Tons)) %>%
    rename(New.Mill = Mill.Fix)
  
  rawDataAll <- 
    left_join(rawDataAll, 
              select(mill_fix, Grade.Fix, Cal.Fix, Dia.Fix, Wind.Fix, 
                     Width.Fix, Old.Mill, New.Mill),
              by = c("Grade.Fix", "Cal.Fix", "Dia.Fix", "Wind.Fix", 
                     "Width.Fix", "Mill.Fix" = "Old.Mill")) %>%
    mutate(Mill.Fix = ifelse(is.na(New.Mill), Mill.Fix, New.Mill)) %>%
    select(-New.Mill)
  
  # Check where the mill is now different
  rawDataAll %>% filter(Mill.Orig != Mill.Fix)
  
  if(tons_check) rawDataAll %>% filter(Width.Fix %in% 
                                         c(62.875, 64.0625, 64.25) & 
                                         Grade.Fix == "AKPG") %>%
    group_by(Mill.Fix, Grade.Fix, Cal.Fix, Width.Fix) %>%
    summarize(Tons = sum(Tons))
  
  # For PLT00070, convert all AKPG 21 pt 62.875 to 24 pt 64.25 
  # 21 pt 64.0625 to 24 pt 62.875
  dt <- data.table(rawDataAll)
  dt[(Grade.Fix == "AKPG" & Plant == "PLT00070" & Cal.Fix == "21" & Width.Fix == 62.875), 
     `:=` (Cal.Fix = "24", Width.Fix = 64.25)]
  dt[(Grade.Fix == "AKPG" & Plant == "PLT00070" & Cal.Fix == "21" & Width.Fix == 64.0625), 
     `:=` (Cal.Fix = "24", Width.Fix = 62.875)]
  
  if(tons_check) dt %>% filter(Width.Fix %in% c(62.875, 64.0625, 64.25) & Grade.Fix == "AKPG") %>%
    group_by(Mill.Fix, Grade.Fix, Cal.Fix, Dia.Fix, Width.Fix) %>%
    summarize(Tons = sum(Tons))
  
  
  # The following per Pat Averitte
  dt[(Grade.Fix == "AKPG" & Cal.Fix == "24" & Width.Fix == 62.8125), 
     `:=` (Width.Fix = 62.875)]
  dt[(Grade.Fix == "AKPG" & Cal.Fix == "27" & Width.Fix == 47.0625), 
     `:=` (Width.Fix = 47.125)]
  
  # It looks like the 27 pt. AKPG 47.5 rolls converted to 24 lb in 
  # 2016.  Make them all 24, regardless OD
  dt[(Grade.Fix == "AKPG" & Cal.Fix == "27" & Width.Fix == 47.5),
     `:=` (Cal.Fix = "24")]
  
  # Pure Leaf change @ Perry - 3 up on the 56.25"
  dt[Grade.Fix == "AKPG" & Cal.Fix == "21" & Width.Fix == 41.75, 
     Width.Fix := 56.25]
  
  # Shiner @ Perry change
  dt[Grade.Fix == "AKPG" & Cal.Fix == "24" & Width.Fix == 37.5,
     Width.Fix := 39.375]
  
  # Change the wind on the SNEEK grades to "W"
  dt[Grade.Fix == "OMSN", Wind.Fix := "W"]
  
  # Make the 2015 AKPG M1 18-72-F 37.875 into 18-77-F to match 2016
  dt[(Grade.Fix == "AKOS" & Mill.Fix == "M1" & Cal.Fix == "18" & Dia.Fix == 72 & 
        Width.Fix == 37.875 & Year == "2015" & Wind.Fix == "F"), 
     `:=` (Dia.Fix = 77)]
  
  # Make the 38.125 wide AKPG 21 pt. that are 81 OD equal to 84
  dt[(Grade.Fix == "AKPG" & Mill.Fix == "W6" & Cal.Fix == "21" & Dia.Fix == 81 & 
        Width.Fix == 38.125 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Dia.Fix = 84)]
  
  # Others that I found just looking at the data....
  dt[(Grade.Fix == "AKPG" & Mill.Fix == "W6" & Cal.Fix == "21" & Dia.Fix == 84 & 
        Width.Fix == 42.5 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Dia.Fix = 77)]
  
  dt[(Grade.Fix == "AKPG" & Mill.Fix == "W6" & Cal.Fix == "21" & Dia.Fix == 77 & 
        Width.Fix == 56.25 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Dia.Fix = 84)]
  
  dt[(Grade.Fix == "AKPG" & Mill.Fix == "W6" & Cal.Fix == "18" & Dia.Fix == 77 & 
        Width.Fix == 45.25 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Dia.Fix = 81)]
  
  dt[(Grade.Fix == "AKOS" & Mill.Fix == "W7" & Cal.Fix == "24" & Dia.Fix == 72 & 
        Width.Fix == 39.5669 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Mill.Fix = "W6", Grade.Fix = "AKHS")]
  
  dt[(Grade.Fix == "AKPG" & Mill.Fix == "W6" & Cal.Fix == "24" & Dia.Fix == 81 & 
        Width.Fix == 31 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Dia.Fix = 84)]
  
  dt[(Grade.Fix == "AKOS" & Mill.Fix == "W6" & Cal.Fix == "24" & Dia.Fix == 72 & 
        Width.Fix == 42.875 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Dia.Fix = 77)]
  
  dt[(Grade.Fix == "AKPG" & Mill.Fix == "M2" & Cal.Fix == "18" & Dia.Fix == 72 & 
        Width.Fix == 45.25 & Year == "2015" & Wind.Fix == "W"), 
     `:=` (Dia.Fix = 81)]
  
  rawDataAll <- as_data_frame(dt)
  rm(dt)
  
  if(tons_check) rawDataAll %>% 
    filter(Width.Fix %in% c(62.875, 64.0625, 64.25) & Grade.Fix == "AKPG") %>%
    group_by(Mill.Fix, Cal.Fix, Dia.Fix, Width.Fix) %>%
    summarize(Tons = sum(Tons))
  
  
  #  browser()
  
  # Delete these items:
  rawDataAll <- rawDataAll %>%
    mutate(Grade.Fix = ifelse(Mill.Orig == "W6" & Cal.Orig == "24" & 
                                Grade.Orig == "AKPG" & Width.Orig == 31.75, 
                              "DELETE", Grade.Fix)) %>%
    #mutate(Grade.Fix = ifelse(Mill.Orig == "W6" & Cal.Orig == "21" & 
    #                            Grade.ORig == "AKPG" & Width.Orig == 42.75,
    #       "DELETE", Grade.Fix)) %>%
    mutate(Grade.Fix = ifelse(Mill.Orig == "W6" & Cal.Orig == "24" & 
                                Grade.Orig == "AKOS" & Width.Orig == 37.875,
                              "DELETE", Grade.Fix)) %>%
    mutate(Grade.Fix = ifelse(Mill.Orig == "W6" & Cal.Orig == "28" & 
                                Grade.Orig == "PKXX" & Width.Orig == 43.25,
                              "DELETE", Grade.Fix))
  
  
  # Update to AKHS for the full year
  akhs_to_fix <- 
    tribble( 
      ~Mill, ~Plant, ~Grade, ~Cal, ~Dia, ~Wind, ~Width, ~New.Grade, ~New.Mill,
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
                          by = c("Mill.Fix" = "Mill",
                                 "Plant" = "Plant",
                                 "Grade.Fix" = "Grade",
                                 "Cal.Fix" = "Cal",
                                 "Dia.Fix" = "Dia",
                                 "Wind.Fix" = "Wind",
                                 "Width.Fix" = "Width"))
  
  rawDataAll <- rawDataAll %>%
    mutate(Grade.Fix = ifelse(!is.na(New.Grade), New.Grade, Grade.Fix),
           Mill.Fix = ifelse(!is.na(New.Mill), New.Mill, Mill.Fix)) %>%
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
  # skus_to_fix <- 
  #   plyr::ldply(plant_list, 
  #         function(x) {
  #           read.xlsx(file.path("DataValidationResponses", 
  #                               #"ParentValidationBase2016 - Pre-run Consolidation Options v3.xlsx"),
  #                               "ParentValidation 2016 Base 03.29.17 RS.xlsx"),
  #                     sheet = x, startRow = 3, cols = 1:20)
  #         }
  #         , .id = NULL)

  skus_to_fix <- 
    map_df(plant_list,          
         function(x) {
           read.xlsx(file.path("DataValidationResponses", 
                               "ParentValidation 2016 Base 03.29.17 RS.xlsx"),
                     sheet = x, startRow = 3, cols = 1:25)
         })
  
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
                                 "Mill.Fix" = "Machine",
                                 "Grade.Fix" = "Grade",
                                 "Cal.Fix" = "Caliper",
                                 "Dia.Fix" = "Dia",
                                 "Wind.Fix" = "Wind",
                                 "Width.Fix" = "Actual.Width"))
  
  rawDataAll <- rawDataAll %>%
    mutate(Mill.Fix = ifelse(!is.na(New.Machine), New.Machine, Mill.Fix),
           Grade.Fix = ifelse(!is.na(New.Grade), New.Grade, Grade.Fix),
           Cal.Fix = ifelse(!is.na(New.Caliper), New.Caliper, Cal.Fix),
           Dia.Fix = ifelse(!is.na(New.Dia), New.Dia, Dia.Fix),
           Wind.Fix = ifelse(!is.na(New.Wind), New.Wind, Wind.Fix),
           Width.Fix = ifelse(!is.na(New.Width), New.Width, Width.Fix))
  
  rawDataAll <- rawDataAll %>%
    select(-starts_with("New"))
  
  rm(skus_to_sub)
  
  # Delete data where the SKU is obsolete or a trial
  skus_to_del <- skus_to_fix %>%
    filter(Desired.Board.Index == 0)
  
  skus_to_del <- skus_to_del %>%
    select(Plant, Machine, Grade, Caliper, Dia, Wind, Actual.Width) %>%
    mutate(New.Grade = "DELETE")
  
  rawDataAll <- 
    left_join(rawDataAll, skus_to_del,
              by = c("Plant" = "Plant",
                     "Mill.Fix" = "Machine",
                     "Grade.Fix" = "Grade",
                     "Cal.Fix" = "Caliper",
                     "Dia.Fix" = "Dia",
                     "Wind.Fix" = "Wind",
                     "Width.Fix" = "Actual.Width"))
  
  rawDataAll <- rawDataAll %>%
    mutate(Grade.Fix = ifelse(!is.na(New.Grade), New.Grade, Grade.Fix)) %>%
    select(-New.Grade)
  
  rm(skus_to_del)
  rm(skus_to_fix)
  
  
  if(tons_check) rawDataAll %>% filter(Width.Fix %in% c(62.875, 64.0625, 64.25) & Grade.Fix == "AKPG") %>%
    group_by(Mill.Fix, Cal.Fix, Width.Fix) %>%
    summarize(Tons = sum(Tons))
  
  
  # Make Perry and West Monroe minimum 77
  rawDataAll <- rawDataAll %>%
    mutate(Dia.Fix = ifelse(Plant %in% c("PLT00070", "PLT00068") &
                              Dia.Fix < 77, 77, Dia.Fix))
  
  # Need to run this again to pick up the Perry and WM changes above
  # Pick the maximum diameter that shipped from
  # each mill
  rawDataAll <- rawDataAll %>% 
    group_by(Mill.Fix, Grade.Fix, Cal.Fix, Wind.Fix, Width.Fix, I.M) %>%
    mutate(Max.Diam = max(Dia.Fix)) %>%
    mutate(Dia.Fix = Max.Diam) %>%
    ungroup()
  
  
  rawDataAll <- rawDataAll %>%
    mutate(Cal.Orig = as.character(Cal.Orig),
           Cal.Fix = as.character(Cal.Fix))
  
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
                          by = c("Mill.Fix" = "Mill", 
                                 "Grade.Fix" = "Grade", 
                                 "Cal.Fix" = "Cal", 
                                 "Width.Fix" = "Width"))
  
  rm(bev_mill_fix)
  
  rawDataAll <- rawDataAll %>%
    mutate(Mill.Fix = ifelse(!is.na(New.Mill), New.Mill, Mill.Fix)) %>%
    select(-New.Mill)
  
  # Do this one more time to consolidate diameters from the adjusted mills
  rawDataAll <- rawDataAll %>% 
    group_by(Mill.Fix, Grade.Fix, Cal.Fix, Wind.Fix, Width.Fix, I.M) %>%
    mutate(Max.Diam = max(Dia.Fix)) %>%
    mutate(Dia.Fix = Max.Diam) %>%
    ungroup()
  
  
  if(tons_check) rawDataAll %>% filter(Width.Fix %in% c(62.875, 64.0625, 64.25) & Grade.Fix == "AKPG") %>%
    group_by(Mill.Fix, Cal.Fix, Width.Fix) %>%
    summarize(Tons = sum(Tons))
  
  
  # What has changed
  rawDataAllFixes <- rawDataAll %>%
    filter(Mill.Orig != Mill.Fix | Grade.Orig != Grade.Fix | Cal.Orig != Cal.Fix |
             Dia.Orig != Dia.Fix | Wind.Orig != Wind.Fix)
  
  write_csv(rawDataAllFixes,
            paste0("Data/shipment-fixes-", mYear, ".csv", sep  = ""))
  
  rawDataAll <- rawDataAll %>%
    unite(CalDW, Cal.Fix, Dia.Fix, Wind.Fix, sep = "-")
  
  
  # Aggregate by date and just keep the fixed data
  rawDataAll <- rawDataAll %>% 
    group_by(Date, Mill.Fix, Plant, Grade.Fix, CalDW, Width.Fix, I.M) %>%
    summarize(Tons = sum(Tons)) %>%
    rename(Mill = Mill.Fix, Grade = Grade.Fix, Width = Width.Fix) %>%
    ungroup() %>%
    filter(Grade != "DELETE")
  
  # How many SKUs are we left with
  rawDataAll %>% 
    summarize(Prod.Count = n_distinct(Mill, Grade, CalDW, Width),
              Plants = n_distinct(Plant))
  
  rawDataAll <- left_join(rawDataAll, facility_names,
                          by = c("Plant" = "Facility.ID"))
  
  # Save as a csv for faster read-in
  write_csv(rawDataAll, 
            paste0("Data/all-shipments-fixed-", mYear, ".csv", sep  = ""))
}

start_date <- min(rawDataAll$Date)
end_date <- max(rawDataAll$Date)

# Check the tons and # of SKUs
rawDataAll %>% group_by(Year) %>%
  unite(MGC, Mill, Grade, CalDW, Width) %>%
  summarize(Tons = round(sum(Tons), 0),
            SKU.Count = n_distinct(MGC))

# Determine the return freight from each plant.  
# This is read in the `User Inputs.R` script, but if not 
# read it in here
if (!exists("trim_freight")) {
  shippingData <- file.path(basepath, "Data", "ShipmentData.xlsx")
  trim_freight <- read.xlsx(shippingData, sheet = "Sheet1")
  trim_freight <- trim_freight %>%
    rename(Cost.Ton = `Cost/Ton`)
}
trim_freight_plant <- rawDataAll %>%
  group_by(Mill, Plant, Grade, CalDW, Width, Facility.Name) %>%
  dplyr::summarize(Tons = sum(Tons)) %>%
  left_join(trim_freight,
            by = c("Mill" = "Mill",
                   "Plant" = "Plant")) %>%
  ungroup()

# If we don't have a freight cost, make it $150
trim_freight_plant$Cost.Ton[is.na(trim_freight_plant$Cost.Ton)] <- 150

# Calculate the cost per ton for each plant based on the weighted volume 
# of each size roll shipping to each plant
trim_freight_plant <- trim_freight_plant %>%
  group_by(Mill, Grade, CalDW, Width) %>%
  mutate(Freight.Cost = Tons * Cost.Ton) %>%
  dplyr::summarize(Trim.Frt.Cost = sum(Tons * Cost.Ton)/sum(Tons))
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

} else if (str_sub(projscenario, 1, 3) == "Eur") {
  rawDataAll <- rawDataAll %>%
    filter(str_sub(Plant, 1, 5) == "PLTEU")
} else if (str_sub(projscenario, 1, 3) == "VMI") {
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
} else if (str_sub(projscenario, 1, 7) == "CPDOnly") {
  # Use rd_comb
  if (!exists("rd_comb")) source(file.path(proj_root, "R", 
                                           "YoY Data Comp.R"))
  rd_comb$Year <- as.numeric(rd_comb$Year)
  
  rawDataAll <- 
    semi_join(rawDataAll, 
              filter(rd_comb, CPD.Web | CPD.Sheet),
              by = c("Year", "Mill", "Grade", "CalDW", "Width"))
  
  # Exclude the shipments going to Perry and West Monroe
  rawDataAll <- rawDataAll %>%
    filter(!(Plant %in% c("PLT00070", "PLT00068")))
}

# These items are fixed and cannot be cut from a parent any larger 
# than Max.Width
# The logic in the 4a.Sub Parent Roll Optimization.R script uses this table
prod_fixed <- 
  tribble( ~Mill, ~Grade, ~CalDW, ~Width, ~Max.Width,
           #----------------------------
           "W6", "AKPG", "27-84-W", 44.375, 45.0, # Patricia Averitte
           #"W6", "AKPG", "27-84-W", 45.75,  # Patricia Averitte
           "W6", "AKPG", "24-84-W", 37.75, 37.75,  # John Zopf for trim
           #"W6", "AKPG", "24-84-W", 42.875, 42.875,  # Poor trim?
           "W6", "AKPG", "21-84-W", 37.0,   37.0)

             
print("Finished Reading Roll Shipment Data", quote = FALSE)

