#------------------------------------------------------------------------------
#
# Analyze the raw demand for the SUS rolls shipping to the GPI plants.
# 
# Data provided by Paul Vorhauer for 2014, 2015 & 2016
# 

library(tidyverse)
library(lubridate)
library(stringr)
library(scales)
library(openxlsx)

#rawData <- read_csv("Data/2016 All_Consumption.csv")
#rawData <- rawData %>%
#  select(-X1)
#rawData$Date <- as.Date(rawData$Date)

# Change the caliper to put the actual caliper first
#rawData <- rawData %>%
#  mutate(Caliper = paste(str_sub(Caliper, 4, 5),
#                        str_sub(Caliper, 1, 2),
#                         str_sub(Caliper, 7, 7), sep = "-"))

# Make sure data is agregated by Date
#rawData <- rawData %>%
#  group_by(Date, Mill, Plant, Grade, Caliper, Width) %>%
#  summarize(Tons = sum(Tons)) %>% ungroup()

facility_names <- read_csv("Data/FacilityNames.csv")

plantCodes <- read.xlsx("Data/GPICustomerPlantCodes.xlsx")
plantCodes <- plantCodes %>%
  select(Cust, SAP.Code, `Int/Ext/whse`) %>% 
  filter(!is.na(Cust)) %>%
  distinct(Cust, SAP.Code, .keep_all = TRUE)
# Make sure no duplicate Cust codes
plantCodes %>% group_by(Cust) %>% 
  mutate(count = n()) %>% filter(count > 1)


# Read the 2014 dat set from Paul
rawData2014 <- read.xlsx("Data/2014 Shipments for Dylan.xlsx",
  sheet = "Shipment data",
  detectDates = TRUE,
  check.names = TRUE
)

# and removed all formatting to make the read statement work
#rawData2014 <- read_csv("Data/2014 Shipments for Dylan Reduced.csv")
rawData2014 <- rawData2014 %>%
  rename(Mill = MFGPlant, Cal = `Cal.lb`, Width = `Calc.width`,
         Grade.Type = `Grade.Type`, Date = `Ac.GI.date`)

rawData2014 <- rawData2014 %>%
  select(Ship.to, Description, Batch, Mill, Date, Cal,
         Grade, Width, Dia, Machine, Grade.Type)

# Just SUS
rawData2014 <- rawData2014 %>%
  filter(Grade.Type == "SUS")
# Exclude Japan
rawData2014 <- rawData2014 %>%
  filter(Grade != "AKJP")
rawData2014 <- rawData2014 %>%
  mutate(Mill = ifelse(is.na(Mill), 
                       paste("00", str_sub(Batch, 1, 2), sep = ""), 
                       Mill))
#rawData2014$Date <- as.Date(rawData2014$Date, "%m/%d/%Y") 
rawData2014 <- rawData2014 %>%
  mutate(Wind = str_sub(Description, str_length(Description), -1))

rawData2014 <- left_join(rawData2014, plantCodes,
                         by = c("Ship.to" = "Cust"))

# Exclude the R&D in West Monroe and mills
rawData2014 <- rawData2014 %>%
  filter(!(SAP.Code %in% c("5538", "5539", "PLT00031", "PLT00033", "PLT00040")))

# Exclude Verst
rawData2014 <- rawData2014 %>%
  filter(`Ship-to` != "PLT00508")

rawData2014 <- rawData2014 %>%
  group_by(Date, Mill, Grade, Cal, Wind, Dia, Width, Machine, SAP.Code) %>%
  summarize(Tons = sum(`Net Ton`)) %>%
  ungroup()

# If we don't have a SAP.Code (Plant), then it should be open market
# or Japan
rawData2014 <- rawData2014 %>%
  filter(!is.na(SAP.Code)) %>%
  rename(Plant = SAP.Code)

rawData2014 %>% summarize(sum(Tons))

# Figure out what to do with the missing machines
rawData2014 <- rawData2014 %>%
  mutate(Machine = ifelse(
    is.na(Machine),
    ifelse(
      Mill == "0031",                # WM
      ifelse(Cal >= 24, "06", "07"),
      ifelse(Cal >= 18, "02", "01")  # Macon
    ),
    Machine
  ))

# Build the Plant-Machine field
rawData2014 <- rawData2014 %>% 
  mutate(Mill = ifelse(Mill == "0031", "W", "M"),
         Year = "2015") %>%
  mutate(Mill = paste(Mill, as.numeric(Machine), sep = "")) %>%
  select(-Machine)

rawData2014 <- rawData2014 %>% 
  group_by(Year, Date, Mill, Grade, Cal, Wind, Dia, Width, Plant) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()


# Read the 2015 data set from Paul.  Deleted certain columns
# and removed all formatting to make the read statement work
rawData2015 <- read_csv("Data/2015 Shipments for Dylan Reduced.csv")
rawData2015 <- rawData2015 %>%
  rename(Mill = MFGPlant, Cal = `Cal/lb`, Width = `Calc width`,
         Grade.Type = `Grade Type`, Date = `Ac.GI date`)
# Just SUS
rawData2015 <- rawData2015 %>%
  filter(Grade.Type == "SUS")
# Exclude Japan
rawData2015 <- rawData2015 %>%
  filter(Grade != "AKJP")
rawData2015 <- rawData2015 %>%
  mutate(Mill = ifelse(is.na(Mill), 
                       paste("00", str_sub(Batch, 1, 2), sep = ""), 
                       Mill))
rawData2015$Date <- as.Date(rawData2015$Date, "%m/%d/%Y") 
rawData2015 <- rawData2015 %>%
  mutate(Wind = str_sub(Description, str_length(Description), -1))

rawData2015 <- left_join(rawData2015, plantCodes,
                        by = c("Ship-to" = "Cust"))

# Exclude the R&D in West Monroe and mills
rawData2015 <- rawData2015 %>%
  filter(!(SAP.Code %in% c("5538", "5539", "PLT00031", "PLT00033", "PLT00040")))

# Exclude Verst
rawData2015 <- rawData2015 %>%
  filter(`Ship-to` != "PLT00508")

rawData2015 <- rawData2015 %>%
  group_by(Date, Mill, Grade, Cal, Wind, Dia, Width, Machine, SAP.Code) %>%
  summarize(Tons = sum(`Net Ton`)) %>%
  ungroup()

# If we don't have a SAP.Code (Plant), then it should be open market
# or Japan
rawData2015 <- rawData2015 %>%
  filter(!is.na(SAP.Code)) %>%
  rename(Plant = SAP.Code)

rawData2015 %>% summarize(sum(Tons))

# Figure out what to do with the missing machines
rawData2015 <- rawData2015 %>%
  mutate(Machine = ifelse(
    is.na(Machine),
    ifelse(
      Mill == "0031",                # WM
      ifelse(Cal >= 24, "06", "07"),
      ifelse(Cal >= 18, "02", "01")  # Macon
    ),
    Machine
  ))

# Build the Plant-Machine field
rawData2015 <- rawData2015 %>% 
  mutate(Mill = ifelse(Mill == "0031", "W", "M"),
         Year = "2015") %>%
  mutate(Mill = paste(Mill, as.numeric(Machine), sep = "")) %>%
  select(-Machine)

rawData2015 <- rawData2015 %>% 
  group_by(Year, Date, Mill, Grade, Cal, Wind, Dia, Width, Plant) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

#-------- 2016 Data --------
rawData2016 <- read_csv("Data/2016 Shipments for Dylan Reduced.csv")

rawData2016 <- rawData2016 %>%
  rename(Mill = MFGPlant, Cal = `Cal/lb`, Width = `Calc width`,
         Grade.Type = `Grade Type`, Date = `Ac.GI date`)
rawData2016 <- rawData2016 %>%
  filter(Grade.Type == "SUS")
rawData2016 <- rawData2016 %>%
  filter(Grade != "AKJP")

rawData2016 <- rawData2016 %>%
  mutate(Mill = ifelse(is.na(Mill), 
                       paste("00", str_sub(Batch, 1, 2), sep = ""), 
                       Mill))
rawData2016$Date <- as.Date(rawData2016$Date, "%m/%d/%Y") 
rawData2016 <- rawData2016 %>%
  mutate(Wind = str_sub(Description, str_length(Description), -1))

rawData2016 <- left_join(rawData2016, plantCodes,
                         by = c("Ship-to" = "Cust"))

# Exclude the R&D in West Monroe and mills
rawData2016 <- rawData2016 %>%
  filter(!(SAP.Code %in% c("5538", "5539", "PLT00031", "PLT00033", "PLT00040")))

# Exclude Verst
rawData2016 <- rawData2016 %>%
  filter(`Ship-to` != "PLT00508")

rawData2016 <- rawData2016 %>% ungroup()

rawData2016 <- rawData2016 %>%
  group_by(Date, Mill, Grade, Cal, Wind, Dia, Width, Machine, SAP.Code) %>%
  summarize(Tons = sum(`Net Ton`)) %>%
  ungroup()

# If we don't have a SAP.Code (Plant), then it should be open market
# or Japan
rawData2016 <- rawData2016 %>%
  filter(!is.na(SAP.Code)) %>%
  rename(Plant = SAP.Code)

rawData2016 %>% summarize(sum(Tons))

# Figure out what to do with the missing machines
rawData2016 <- rawData2016 %>%
  mutate(Machine = ifelse(
    is.na(Machine),
    ifelse(
      (Mill == "0031" | Mill == "31"),            # WM
      ifelse(Cal >= 24, "06", "07"),
      ifelse(Cal >= 18, "02", "01")  # Macon
    ),
    Machine
  ))

# Build the Plant-Machine field
rawData2016 <- rawData2016 %>% 
  mutate(Mill = ifelse((Mill == "0031" | Mill == "31"), "W", "M"),
         Year = "2016") %>%
  mutate(Mill = paste(Mill, as.numeric(Machine), sep = "")) %>%
  select(-Machine)

rawData2016 <- rawData2016 %>% 
  group_by(Year, Date, Mill, Grade, Cal, Wind, Dia, Width, Plant) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

#--------

rawData <- bind_rows(rawData2015, rawData2016) 

# Check the tons by plant
temp <- rawData %>% 
  group_by(Year, Plant) %>%
  summarize(Tons = sum(Tons)) %>%
  spread(Year, Tons, fill = "") %>%
  group_by(Plant) %>%
  summarize(Tons2015 = sum(as.numeric(`2015`)),
            Tons2016 = sum(as.numeric(`2016`)))

#rawDataSplit <- rawData %>%
#  separate(Caliper, c("Cal", "Diam", "Wind")) %>%
#  mutate(Year = year(Date))  %>%
#  group_by(Year, Plant, Grade, Cal, Diam, Wind, Width) %>%
#  filter(Year == "2016") %>%
#  summarize(Tons = round(sum(Tons), 2), nbr_dmd = n()) %>%
#  group_by()

rawData <- left_join(rawData, facility_names,
                          by = c("Plant" = "Facility.ID"))
rawData <- rawData %>%
  select(Year, Date, Mill, Plant, Facility.Name, Grade, Cal, Dia, Wind, Width, Tons)

# Summarize the Tons and the number of orders (NBR.Dmd) by "SKU"
rawDataSKU <- rawData %>%
  group_by(Year, Mill, Grade, Cal, Dia, Wind, Width) %>%
  summarize(Tons = sum(Tons), Nbr.Dmd = n()) %>% 
  ungroup() %>%
  group_by(Year, Grade, Nbr.Dmd) %>%
  mutate(Nbr.Grade.SKUs = n()) %>%  # Nbr of `Nbr.Dmd`` SKUs by Grade
  ungroup()

ggplot(data = filter(rawDataSKU, Nbr.Dmd <= 10) %>%
         group_by(Year, Grade, Nbr.Dmd) %>% 
         summarize(Tons = sum(Tons), SKU.Count = mean(Nbr.Grade.SKUs)), 
       aes(x = Nbr.Dmd, y = Tons, fill = Grade)) +
  geom_bar(stat = "identity") +
  geom_text(aes(label = SKU.Count), position = position_stack(vjust = 0.5)) +
  scale_y_continuous(labels = comma) +
  facet_grid(Year ~ .)

# Count the number of different diameter rolls of the same width 
# shipping to the same facility
rawDataDiam <- rawData %>% group_by(Year, Plant, Facility.Name, Mill, Grade, Cal, Width, Dia) %>%
  summarize(Tons = sum(Tons), Nbr.Dmd = n()) %>% 
  mutate(nbr_diams = n_distinct(Dia)) %>%
  filter(nbr_diams > 1) %>% 
  gather(Var, val, Nbr.Dmd, Tons) %>%
  unite(Var1, Var, Dia) %>%
  spread(Var1, val, fill = "")

copy.table(rawDataDiam)

# Do the same for the wind direction
rawDataWind <- rawData %>% group_by(Year, Plant, Mill, Grade, Cal, Width, Wind) %>%
  summarize(Tons = sum(Tons), Nbr.Dmd = n()) %>% 
  mutate(nbr_winds = n_distinct(Wind), Nbr_plants = n_distinct(Plant)) %>%
  filter(nbr_winds > 1) %>% 
  gather(Var, val, Nbr.Dmd, Tons) %>%
  unite(Var1, Var, Wind) %>%
  spread(Var1, val, fill = "") %>%
  arrange(Grade, Cal, Width)

copy.table(rawDataSplitWind)

# Consolidate Diameters
# First make anything less than 72 a 72
rawData <- rawData %>%
  mutate(Dia = ifelse(Dia < 72, 77, Dia))
# Now for any plant, round up to the max roll diameter
rawData <- rawData %>% 
  group_by(Year, Mill, Grade, Cal, Wind, Width) %>%
  mutate(Max.Diam = max(Dia)) %>%
  mutate(Dia = Max.Diam)

rawDataSKU <- rawData %>%
  group_by(Year, Mill, Grade, Cal, Dia, Wind, Width) %>%
  summarize(Tons = sum(Tons), Nbr.Dmd = n()) %>% 
  ungroup()

rawData <- rawData %>% select(-Max.Diam) %>%
  ungroup()

write_csv(rawData, "Data/All Shipments Fixed.csv")

copy.table(rawDataSKU)
