###############################################################################
#
# Read in the data frile from Paul with the raw shipment data
#
# Fix a bunch of stuff and save as a csv with just the data
#  as All Shipments No SKU Fix 

pd_orig <- read.xlsx("Data/2017.5 Shipments for Dylan.xlsx", 
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

pd <- as_tibble(pd_orig)
pd <- pd %>%
  select(Ship.to, Sold.to.pt, Description, Batch, MFGPlant, 
         Date = Ac.GI.date, Cal = Cal.lb, Grade, Width = Calc.width, 
         Dia, Machine, Tons = Net.Ton, Grade.Type, I.M)

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
# (uses new 0.7 syntax)
pd <- pd %>%
  group_by_at(vars(names(pd), -Tons)) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

pd <- pd %>%
  mutate(Wind = str_sub(Description, str_length(Description), -1))

pd <- left_join(pd, plant_codes,
                        by = c("Ship.to" = "Cust"))

# Exclude the R&D sold-to
pd <- pd %>% filter(Sold.to.pt != "5538")

# Exclude the R&D in West Monroe and mills
pd <- pd %>%
  filter(!(SAP.Code %in% c("5538", "5539", "PLT00031", "PLT00033", "PLT00040")))

# Exclude Verst and Piscataway
pd <- pd %>%
  filter(!(SAP.Code %in% c("PLT00508", "PLT00019")))

# If the ship to is a Europe port, then use the sold-to to get the actual
# location.  Rename the SAP.Code field so it doesn't conflict with what we 
# already have
pd <-
  left_join(pd, rename(select(plant_codes, Cust, SAP.Code),
                               SAP.Sold.To = SAP.Code),
            by = c("Sold.to.pt" = "Cust"))

# Where are they different
pd %>% filter(SAP.Code != SAP.Sold.To) %>%
  group_by(Ship.to, SAP.Code, Sold.to.pt, SAP.Sold.To) %>%
  summarize(Tons = sum(Tons))

# Update the SAP.Code to the sold where they are not equal (and where the
# SAP.Sold.To is not NA).  This fixes a Portland issue, too.
pd <- pd %>%
  mutate(SAP.Code = ifelse(!is.na(SAP.Sold.To) & (SAP.Code != SAP.Sold.To), 
                           SAP.Sold.To, SAP.Code))

# Get rid of the Sold.To
pd <- pd %>%
  select(-SAP.Sold.To)

# If we don't have an SAP.Code (Plant), then it should be open market
# or Japan and we don't want it
pd <- pd %>%
  ungroup() %>%
  filter(!is.na(SAP.Code)) %>%
  rename(Plant = SAP.Code)

# Check the total tons left
pd %>% summarize(sum(Tons))

# Figure out what to do with the missing machines (2015 data)
pd <- pd %>%
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
pd <- pd %>% 
  mutate(Mill = ifelse((Mill == "0031" | Mill == "31"), "W", "M")) %>%
  mutate(Mill = paste(Mill, as.numeric(Machine), sep = "")) %>%
  select(-Machine)

# Get rid of the following:
pd <- pd %>%
  select(-Ship.to, -Sold.to.pt, -Description, -Int.Ext.Whs, -Grade.Type)

# Summarize again over the fields that are left
pd <- pd %>%
  group_by_(.dots = setdiff(names(pd), "Tons")) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

# How many SKUs are we starting with
pd %>% 
  summarize(Prod.Count = n_distinct(Mill, Grade, Cal, Wind, Width, Dia),
            Plants = n_distinct(Plant))


pd16 <- read_csv(paste0("Data/All Shipments No SKU Fix ", "2016", ".csv"))
pd16$Cal <- as.character(pd16$Cal)

pd16 <- pd16 %>%
  filter(Date >= "2016-07-01")

pd <- bind_rows(pd16, pd)

# Save as a csv at this point - prior to diameter consolidation
write_csv(pd, 
          paste0("Data/All Shipments No SKU Fix-", "2017.5", ".csv"))

rm(pd, pd16, pd_orig)
