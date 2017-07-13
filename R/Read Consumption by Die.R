#------------------------------------------------------------------------------
#
# Sus Consumption by Die
#
# The file comes from QlikView Fiber Analysis Consumption by Die Report, 
# Bookmark = SUS Consumption with the Date set to Month Year.
#

cons <- read.xlsx(file.path(proj_root, "Data", 
                           "QV SUS Roll Consumption 2016-2017.xlsx"),
                 detectDates = T)

cons <- cons %>%
  mutate(Plant = paste0("PLT0", Plant),
         Board.Width = as.numeric(Board.Width)) %>%
  filter(Board.Width > 0)

# Use the floor_date in case the data starts after the 1st
cons <- cons %>%
  filter(MonthYear >= floor_date(start_date, "month") & MonthYear <= end_date)

cons <- cons %>%
  mutate(M.I = str_sub(Material.Description, 1, 1))

# Need to get the full grade out of the Material.Description field for the "AK"
# records.  If OM or PK, just add "XX".  OTherwise use the Short field
cons <- cons %>%
  mutate(Grade = ifelse(
    Board.Type.Short %in% c("PK", "OM"),
    paste(Board.Type.Short, "XX", sep = ""),
    ifelse(
      Board.Type.Short == "AK",
      ifelse(
        M.I == "I",
        str_sub(Material.Description, 5, 8),   # Imperial
        str_sub(Material.Description, 6, 9)    # Metric
      ), Board.Type.Short)
  ))
cons <- cons %>%
  mutate(Grade = ifelse(Grade == "SUSFC02", "FC02", Grade))

cons <- cons %>%
  mutate(Wind = str_sub(Material.Description, 
                        str_length(Material.Description), -1))
cons <- cons %>%
  mutate(Dia = ifelse(M.I == "I",
                      str_sub(Material.Description, 19, 20),
                      str_sub(Material.Description, 16, 19)))

cons <- cons %>%
  group_by(Plant, PlantName, 
           MRPC2.Text, Die, Die.Description,
           Grade, Board.Caliper, Dia, Wind, Board.Width, M.I) %>%
  dplyr::summarize(Tons = sum(TON)) %>%
  ungroup()

# VMI Consumption
vmi_rolls <- read.xlsx(file.path(proj_root, "Data", "VMI List.xlsx"),
                       startRow = 2, cols = c(1:6))
vmi_rolls[is.na(vmi_rolls)] <- ""
vmi_rolls$Caliper <- as.character(vmi_rolls$Caliper)
vmi_rolls <- vmi_rolls %>%
  select(Caliper, Width, Grade)
# Make sure no duplicates
if (any(duplicated(vmi_rolls))) stop("Duplicates in VMI List")
vmi_rolls <- vmi_rolls %>%
  mutate(Bev.VMI = T)

cons_vmi <- left_join(cons, vmi_rolls,
                        by = c("Grade", 
                               "Board.Caliper" = "Caliper", 
                               "Board.Width" = "Width"))

cons_vmi <- cons_vmi %>%
  mutate(Bev.VMI = ifelse(is.na(Bev.VMI), F, T))

cons_vmi <- cons_vmi %>%
  mutate(SUS.Type = ifelse(Bev.VMI, "VMI", 
                           ifelse(M.I == "M", "Export", "CPD")))

cons_vmi %>% group_by(SUS.Type) %>%
  summarize(Tons = sum(Tons))

# Aggregate over the
dies <- cons %>% 
  select(Grade, Board.Caliper, Dia, Wind, Board.Width, Die, 
         Die.Description) %>%
  group_by(Grade, Board.Caliper, Dia, Wind, Board.Width) %>%
  dplyr::mutate(Nbr.Dies = n(), 
         Die.List = paste0(Die, collapse = ", "),
         Die.Descr.List = paste0(Die.Description, collapse = ", ")) %>%
  select(-Die, -Die.Description) %>%
  ungroup() %>%
  distinct()

dies_plants <- cons %>% 
  select(Plant, Grade, Board.Caliper, Dia, Wind, Board.Width, Die, 
         Die.Description) %>%
  group_by(Plant, Grade, Board.Caliper, Dia, Wind, Board.Width) %>%
  dplyr::mutate(Nbr.Dies = n(), 
         Die.List = paste0(Die, collapse = ", "),
         Die.Descr.List = paste0(Die.Description, collapse = ", ")) %>%
  select(-Die, -Die.Description) %>%
  ungroup() %>%
  distinct()

plant_die_board <- cons %>% 
  group_by(Plant, PlantName, Die) %>%
  summarize(Tons = sum(Tons),
            Board.Count = n()) %>%
  ungroup()

plant_die_board %>% group_by(Plant, PlantName) %>%
  summarize(Boards.Per.Die = sum(Board.Count)/n_distinct(Die)) %>%
  arrange(desc(Boards.Per.Die))

# Clean up
rm(plant_die_board)

