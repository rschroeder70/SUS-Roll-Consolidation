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
#library(plyr)
library(tidyverse)
library(scales)
library(openxlsx)
library(stringr)
library(data.table)

# This projects root directory
if (!exists(proj_root)) proj_root <- 
    file.path("C:", "Users", "SchroedR", "Documents",
                       "Projects", "SUS Roll Consolidation") 

# sr_inv - sus roll inventory
sr_inv <- read.xlsx(file.path(proj_root,"Data",
                    "QV SUS Roll Inv Mnth Trend Det 2016 2017.xlsx"),
                           detectDates = T)

sr_inv$Date <- as.Date(sr_inv$Date, "%m/%d/%Y")

# Just 12 months of data
sr_inv <- sr_inv %>%
  filter(between(Date, as.Date("2016-07-01"), as.Date("2017-06-30")))

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

# Read the fixes file
fixes <- read_csv(file.path(proj_root,"Data",
                            "shipment-fixes-2016.csv"))
fixes <- fixes %>%
  select(Grade.Orig, Cal.Orig, Width.Orig, Wind.Orig,
         Grade.Fix, Cal.Fix, Width.Fix, Wind.Fix) %>%
  distinct(Grade.Orig, Cal.Orig, Width.Orig, Wind.Orig,
           .keep_all = TRUE) %>%
  mutate(Cal.Fix = as.character(Cal.Fix),
         Cal.Orig = as.character(Cal.Orig))

sr_inv <- left_join(sr_inv, fixes,
                    by = c("Grade" = "Grade.Orig",
                           "Caliper" = "Cal.Orig",
                           "Width" = "Width.Orig",
                           "Wind" = "Wind.Orig"))

sr_inv <- sr_inv %>%
  mutate(Grade = ifelse(!is.na(Grade.Fix), Grade.Fix, Grade),
         Caliper = ifelse(!is.na(Cal.Fix), Cal.Fix, Caliper),
         Width = ifelse(!is.na(Width.Fix), Width.Fix, Width),
         Wind = ifelse(!is.na(Wind.Fix), Wind.Fix, Wind))

sr_inv <- sr_inv %>%
  select(-Grade.Fix, -Cal.Fix, -Width.Fix, -Wind.Fix)

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
# roll can have different Prod.Types, so as we aggregate over the Mill & Dia
# we need to pick the most likely product type
opt_detail_pt <- opt_detail %>%
  group_by(Grade, Caliper, Prod.Width, Wind) %>%
  summarize(Prod.Type = first(Prod.Type, order_by = desc(Prod.DMD.Tons)),
            Prod.Policy = first(Parent.Policy, order_by = desc(Prod.DMD.Tons)),
            Prod.DMD.Tons = sum(Prod.DMD.Tons)) %>%
  ungroup()

# get the Prod.Type for each roll in the solution
sr_inv_pt <- left_join(sr_inv, opt_detail_pt,
                       by = c("Grade" = "Grade",
                              "Caliper" = "Caliper",
                              "Width" = "Prod.Width",
                              "Wind" = "Wind")) 

sr_inv_pt <- sr_inv_pt %>%
  rename(Annual.Prod.DMD.Tons = Prod.DMD.Tons)

# Filter out the NAs.  These are Asia and Open Market items
# and any items not in the current solution
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

#Beverage Items to eliminate
elim <- 
  tribble( 
    ~Caliper, ~Grade, ~Width.e, ~Width.p,
    #-----------------------
    "18", "AKPG", 29.5, 31.125,
    "18", "AKPG", 30.125, 31.125,
    "18", "AKPG", 44.125, 45.25,
    "21", "AKPG", 45.5, 46,
    "24", "AKPG", 36.75, 37.75,
    "24", "AKPG", 41.875, 42.25,
    "24", "AKPG", 42.125, 42.25,
    "24", "AKPG", 42.875, 43.125,
    "24", "AKPG", 44.875, 46.625,
    "24", "AKPG", 47.5, 47.625,
    "24", "AKPG", 52.875, 54.5,
    "27", "AKPG", 40.875, 42.375,
    "27", "AKPG", 45.75, 47.125,
    "27", "AKPG", 46.625, 47.125,
    "28", "PKXX", 37.875, 38.625)

sr.inv <- sr_inv %>%
  group_by(Date, Grade, Caliper, Width) %>%
  summarize(Tons = sum(Inv.Tons))

sr.inv.elim <- 
  semi_join(sr.inv, elim,
            by = c("Grade", "Caliper", "Width" = "Width.e"))

# Average for all Items
sr.inv.elim %>% group_by(Date) %>%
  summarize(Tons = sum(Tons)) %>%
  summarize(Tons = mean(Tons))
