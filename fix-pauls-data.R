###############################################################################
#
# Read in the data frile from Paul with the raw shipment data
#
# Fix a bunch of stuff and save as a csv with just the data
#  as All Shipments No SKU Fix 

pd_orig <- read.xlsx("Data/2017 first half Shipments for Dylan.xlsx", 
                "Shipment data", colNames = TRUE,
                detectDates = TRUE, check.names = TRUE)

plant_codes <- read.xlsx("Data/GPICustomerPlantCodes.xlsx")

plant_codes <- plant_codes %>%
  select(Cust, SAP.Code, Int.Ext.Whs = `Int/Ext/whse`) %>% 
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

pd <- as_tibble(pd_orig)
pd <- pd %>%
  select(Ship.to, Sold.to.pt, Description, Batch, MFGPlant, 
         Date = Ac.GI.date, Cal = Cal.lb, Grade, Width = Calc.width, 
         Dia, Machine, Tons = Net.Ton, Grade.Type, I.M)

pd <- pd %>%
  mutate(Cal = as.character(Cal))
pd <- pd %>% filter(Grade.Type == "SUS")
pd <- pd %>% filter(!Grade %in% c("AKJP", "AK25"))
pd <- pd %>% mutate(Year = year(Date))
pd <- pd %>% mutate(I.M = str_trim(I.M))

# Create the Mill field.  If missing a MFGPlant,
# fill in the MFGPlant data from the first two postitions 
# of the Batch nbr
pd <- pd %>%
  mutate(Mill = ifelse(is.na(MFGPlant), 
                       paste("00", str_sub(Batch, 1, 2), sep = ""), 
                       MFGPlant)) 
pd <- pd %>% select(-MFGPlant, -Batch)

# Summarize over batches once we are done with them
pd <- pd %>%
  group_by_(.dots = setdiff(names(pd), "Tons")) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

pd <- pd %>%
  mutate(Wind = str_sub(Description, str_length(Description), -1))

pd <- left_join(pd, plant_codes,
                        by = c("Ship.to" = "Cust"))

# Exclude the R&D in West Monroe and mills
pd <- pd %>%
  filter(!(SAP.Code %in% c("5538", "5539", "PLT00031", "PLT00033", "PLT00040")))

# Exclude Verst and Piscataway
pd <- pd %>%
  filter(!(SAP.Code %in% c("PLT00508", "PLT00019")))

# If the ship to is a Europe port, then use the sold-to to get the actual
