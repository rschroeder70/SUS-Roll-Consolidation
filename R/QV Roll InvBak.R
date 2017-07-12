#------------------------------------------------------------------------------
#
# Look at actual roll inventory history
#
# Data is from the SUS Rollstock Bookmark in the QlikView Inventory
# cube under the Inventory Details tab/Mohthly Trend Details worksheet.  
#
# The data is already in long format
# 
# In order to match up with model results, we need to make the same
# adjustments to the SKUs as we do in the Read Data.R script
# 
#
library(plyr)
library(tidyverse)
library(scales)
library(openxlsx)
library(stringr)
library(data.table)

# This projects root directory
proj_root <- file.path("C:", "Users", "SchroedR", "Documents",
                       "Projects", "SUS Roll Consolidation") 

# sr_inv - sus roll inventory
sr_inv <- read.xlsx(file.path(proj_root,"Data",
                    "QV SUS Roll Inv Mnth Trend Det 2015 2016.xlsx"),
                           detectDates = T)

# Just 2016 data
sr_inv <- sr_inv %>%
  dplyr::filter(between(Date, as.Date("2016-01-01"), as.Date("2016-12-31")))

sr_inv %>% dplyr::filter(Date == as.Date("2016-12-31")) %>%
  summarize(Inv.Tons = sum(Qty.in.TON))

# Ignore material numbers that don't start with a "1"
sr_inv <- sr_inv %>%
  dplyr::filter(str_sub(Material, 1, 1) == "1")

# PM2 is external board and PM9 is Upgrage board
# Note that external on the CRB side is Panther
sr_inv <- sr_inv %>%
  dplyr::filter(!str_sub(Material.Group.Sel, 1, 3) %in% c("PM2", "PM9"))

sr_inv <- sr_inv %>% dplyr::filter(!is.na(Material))

sr_inv %>% dplyr::filter(Date == as.Date("2016-12-31")) %>%
  summarize(Tons = sum(Qty.in.TON))

# Aggregate over the Die
sr_inv <- sr_inv %>%
  group_by(Date, Plt, Plant.Name, Material.Group.Sel, Board.Group,
           Material, Material.Description) %>%
  summarize(Inv.Tons = sum(Qty.in.TON), Inv.Cost = sum(Cost.Val)) %>%
  ungroup()

# Fix the plant codes
sr_inv <- sr_inv %>%
  mutate(Plant = paste0("PLT", str_pad(Plt, 5, side = "left", pad = "0"))) %>%
  select(-Plt)

sr_inv %>% dplyr::filter(Date == as.Date("2016-12-31")) %>%
  #group_by(Date, Plant) %>%
  summarize(Inv.Tons = sum(Inv.Tons))

sr_inv <- sr_inv %>%
  rename(Material.Descr = Material.Description)

# Extract the PMP components from the description
sr_inv <- sr_inv %>%
  mutate(M.I = str_sub(Material.Descr, 1, 1))

sr_inv <- sr_inv %>% 
  mutate(Width = ifelse(M.I == "I",
                        str_sub(Material.Descr, 10, 17),
                        str_sub(Material.Descr, 11, 14)))
sr_inv$Width <- as.numeric(sr_inv$Width)

sr_inv <- sr_inv %>%
  mutate(Caliper = ifelse(M.I == "I",
                          str_sub(Material.Descr, 3, 4),
                          str_sub(Material.Descr, 4, 5)))
sr_inv <- sr_inv %>%
  mutate(Grade = ifelse(M.I == "I",
                        str_sub(Material.Descr, 5, 8),
                        str_sub(Material.Descr, 6, 9)))

sr_inv <- sr_inv %>%
  mutate(Wind = str_sub(Material.Descr, 
                        str_length(Material.Descr), -1))

sr_inv <- sr_inv %>%
  mutate(Dia = ifelse(M.I == "I",
                      str_sub(Material.Descr, 19, 20),
                      str_sub(Material.Descr, 16, 19)))

sr_inv <- sr_inv %>%
  dplyr::filter(Grade %in% c("PKXX", "AKOS", "AKPG", "OMXX", "FC02", 
                      "OMSN", "AKHS", "PKHS"))

# Exclude Verst
sr_inv <- sr_inv %>%
  dplyr::filter(Plant != "PLT00508")

# Aggregate over the diameters & Plants
sr_inv <- sr_inv %>%
  group_by(Date, Grade, Caliper, Width, Wind, M.I) %>%
  summarize(Inv.Tons = sum(Inv.Tons),
            Inv.Cost = sum(Inv.Cost)) %>%
  ungroup()

# Calculate the cost per ton
sr_inv %>% group_by(Grade) %>%
  summarize(CPT = sum(Inv.Cost)/sum(Inv.Tons))

# Identify VMI inventory
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

sr_inv_vmi <- left_join(sr_inv, vmi_rolls,
                        by = c("Grade", "Caliper", "Width"))

sr_inv_vmi <- sr_inv_vmi %>%
  mutate(Bev.VMI = ifelse(is.na(Bev.VMI), F, T))

sr_inv_vmi <- sr_inv_vmi %>%
  mutate(SUS.Type = ifelse(Bev.VMI, "VMI", 
         ifelse(M.I == "M", "Export", "CPD")))

# Compute average for year
sr_inv_vmi %>%
  group_by(SUS.Type, Date) %>%
  summarize(Month.End.Tons = sum(Inv.Tons)) %>%
  summarize(Avg.Tons = mean(Month.End.Tons))

rm(sr_inv_vmi, vmi_rolls)

# Make the same fixes as when we read in the raw data.  
dt <- data.table(sr_inv)
dt[(Grade == "AKPG" & Caliper == "21" & Width == 62.875), 
   `:=` (Caliper = "24", Width = 64.25)]
dt[(Grade == "AKPG" & Caliper == "21" & Width == 64.0625), 
   `:=` (Caliper = "24", Width = 62.875)]

# It looks like the 27 pt. AKPG 47.5 rolls converted to 24 lb in 
# 2016.  Make them all 24, regardless OD
dt[(Grade == "AKPG" & Caliper == "27" & Width == 47.5),
   `:=` (Caliper = "24")]

# Pure Leaf change - 3 up on the 56.25"
dt[Grade == "AKPG" & Caliper == "21" & Width == 41.75, 
   Width := 56.25]

# Shiner change
dt[Grade == "AKPG" & Caliper == "24" & Width == 37.5,
   Width := 39.375]

# Change the wind on the SNEEK grades to "W"
dt[Grade == "OMSN", Wind := "W"]

sr_inv <- as_data_frame(dt)
rm(dt)

# Update to AKHS for the full year
akhs_to_fix <- 
  tribble( ~Mill, ~Plant, ~Grade, ~Cal, ~Dia, ~Wind, ~Width, ~New.Grade, ~New.Mill,
           #-----------------------------------------------------------------
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

# We only have Grade, Caliper, Width and Wind to work with (no Diameters)
akhs_to_fix <- akhs_to_fix %>%
  select(Grade, Cal, Wind, Width, New.Grade) %>%
  distinct()

sr_inv <- left_join(sr_inv, akhs_to_fix,
                        by = c("Grade" = "Grade",
                               "Caliper" = "Cal",
                               "Wind" = "Wind",
                               "Width" = "Width"))

sr_inv <- sr_inv %>%
  mutate(Grade = ifelse(!is.na(New.Grade), New.Grade, Grade)) %>%
  select(-New.Grade)

# Fix substitutions, trials and obsoletes
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
skus_to_sub <- 
  left_join(skus_to_sub, 
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

# Different from Read Data - only keep changes in Grade, Caliper, 
# Width or Wind - No Diameter, Plant or Mill
skus_to_sub <- skus_to_sub %>%
  select(Grade, New.Grade, Caliper, New.Caliper, Actual.Width, New.Width, 
         Wind, New.Wind ) %>% 
  distinct()

sr_inv <- left_join(sr_inv, skus_to_sub,
                        by = c("Grade" = "Grade",
                               "Caliper" = "Caliper",
                               "Wind" = "Wind",
                               "Width" = "Actual.Width"))

# No Mill or Dia
sr_inv <- sr_inv %>% 
  ungroup() %>%
  mutate(Grade = ifelse(!is.na(New.Grade), New.Grade, Grade),
         Caliper = ifelse(!is.na(New.Caliper), New.Caliper, Caliper),
         Wind = ifelse(!is.na(New.Wind), New.Wind, Wind),
         Width = ifelse(!is.na(New.Width), New.Width, Width))

sr_inv <- sr_inv %>%
  select(-starts_with("New"))

rm(skus_to_sub)

# Delete data where the SKU is obsolete or a trial
skus_to_del <- skus_to_fix %>%
  filter(Desired.Board.Index == 0)

skus_to_del <- skus_to_del %>%
  select(Grade, Caliper, Actual.Width, Wind) %>%
  distinct()

sr_inv <- 
  anti_join(sr_inv, skus_to_del,
            by = c("Grade" = "Grade",
                   "Caliper" = "Caliper",
                   "Wind" = "Wind",
                   "Width" = "Actual.Width"))

rm(skus_to_del)
rm(skus_to_fix)

# Convert metric to imperial
sr_inv <- sr_inv %>%
  mutate(Width = ifelse(M.I == "M",
         round(Width/25.4, 4), Width))

sr_inv %>% dplyr::filter(Date == as.Date("2016-12-31")) %>%
  summarize(Inv.Tons = sum(Inv.Tons))

# Make sure we have the opt_detail dataframe available. If not, run the
# 7.1 Opt Solution Post Process.R script to create it
if (!exists("opt_detail")) {
  source ("R/7.1 Opt Solution Post Process.R")
}

# The QV inventory file does not include mill of manufacture, and the same size
# roll can have different Prod.Types, so as we aggregate ove the Mill & Dia
# we need to pick the most likely product type
opt_detail_pt <- opt_detail %>%
  group_by(Grade, Caliper, Prod.Width, Wind) %>%
  summarize(Prod.Type = first(Prod.Type, order_by = desc(Prod.DMD.Tons)),
            Prod.Policy = first(Parent.Policy, order_by = desc(Prod.DMD.Tons)),
            Prod.DMD.Tons = sum(Prod.DMD.Tons)) %>%
  ungroup()

# get the Prod.Type for each roll
sr_inv_pt <- left_join(sr_inv, opt_detail_pt,
                       by = c("Grade" = "Grade",
                              "Caliper" = "Caliper",
                              "Width" = "Prod.Width",
                              "Wind" = "Wind")) 

sr_inv_pt <- sr_inv_pt %>%
  rename(Annual.Prod.DMD.Tons = Prod.DMD.Tons)

# Filter out the NAs.  These are Asia and Open Market items
sr_inv_pt <- sr_inv_pt %>%
  dplyr::filter(!is.na(Prod.Policy))


# December inventory check
sr_inv_pt %>% dplyr::filter(Date == as.Date("2016-12-31")) %>%
  summarize(Dec.Inv.Tons = sum(Inv.Tons))

ggplot(data = sr_inv_pt, aes(x = Date, y = Inv.Tons)) +
  geom_bar(position = "stack", stat = "identity",
           aes(fill = Prod.Type)) +
  scale_y_continuous(name = "Inventory Tons", labels = comma) +
  scale_x_date(name = "Month End") + 
  ggtitle("QlikView Month End SUS Inventory")
  
#copy.table(sr_inv_pt)

# Get the total for each month 
sr_inv_ttl <- sr_inv_pt %>%
  group_by(Date, Grade, Caliper, Width, Wind, Prod.Policy, Prod.Type) %>%
  summarize(Inv.Tons = sum(Inv.Tons),
            Annual.Prod.DMD.Tons = max(Annual.Prod.DMD.Tons))

ggplot(data = sr_inv_ttl, aes(x = Date, y = Inv.Tons)) +
  geom_bar(position = "stack", stat = "identity",
           aes(fill = Prod.Type)) +
  facet_wrap(~ Prod.Policy)

# Compute the mean for each SKU
sr_inv_mean <- sr_inv_ttl %>%
  group_by(Grade, Caliper, Width, Wind, Prod.Policy, Prod.Type) %>%
  summarize(FY.DMD = mean(Annual.Prod.DMD.Tons),
            Avg.OH = mean(Inv.Tons))
  
# Compute the full year average
# First aggregate by Type & Policy for each Month
sr_inv_me <- sr_inv_ttl %>%
  group_by(Date, Prod.Policy, Prod.Type) %>%
  summarize(On.Hand = sum(Inv.Tons),
            Annual.DMD = sum(Annual.Prod.DMD.Tons)) %>%
  mutate(Scenario = "Q.V. Month End") %>%
  ungroup() 

# Full year average
sr_inv_fy <- sr_inv_me %>%
  group_by(Prod.Policy, Prod.Type) %>%
  summarize(Avg.Inv.Tons = mean(On.Hand),
            Annual.DMD = mean(Annual.DMD)) %>%
  arrange(desc(Prod.Policy)) #%>%
  #mutate(Prod.Type = gsub("\\.", " ", Prod.Type))

copy.table(sr_inv_fy)
