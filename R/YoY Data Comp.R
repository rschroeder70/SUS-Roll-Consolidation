#
#library(plyr)
library(tidyverse)
library(lubridate)
library(stringr)
library(openxlsx)
library(scales)
library(RColorBrewer)
library(gridExtra)

source(file.path(proj_root, "R", "FuncReadResults.R"))
basepath <- file.path("C:","Users", "SchroedR", "Documents",
                      "Projects", "RollAssortmentOptimization")

# Load the various functions
source(file.path(basepath, "R", "0.Functions.R"))

#------------------------------------------------------------------------------
#
# Look at the raw data for both years and identify the common items
# We also create the rd_comb data frame that is used elsewhere
#

rd_2015 <- read_csv(file.path(proj_root, "Data",
                              "All Shipments Fixed 2015.csv"))
rd_2016 <- read_csv(file.path(proj_root, "Data",
                             "All Shipments Fixed 2016.csv"))
# Summarize the tons over the dates.  Do this for the individual files
# and then combine as we may want to use this annual data for other
# purposes
rd_2015 <- rd_2015 %>%
  group_by(Year, Mill, Grade, CalDW, Width, Plant, Fac.Abrev) %>%
  dplyr::summarize(Tons = sum(Tons)) %>%
  ungroup()
rd_2016 <- rd_2016 %>%
  group_by(Year, Mill, Grade, CalDW, Width, Plant, Fac.Abrev) %>%
  dplyr::summarize(Tons = sum(Tons)) %>%
  ungroup()
rd_2015$Year <- as.character(rd_2015$Year)
rd_2016$Year <- as.character(rd_2016$Year)
rd_2015 <- rd_2015 %>%
  separate(CalDW, c("Caliper", "Dia", "Wind"), remove = FALSE)
rd_2016 <- rd_2016 %>%
  separate(CalDW, c("Caliper", "Dia", "Wind"), remove = FALSE)

# Combine the raw data into a single df
rd_comb <- bind_rows(rd_2015, rd_2016)

# How much for the new grades in 2016
rd_comb %>% dplyr::filter(Year == "2016" & 
                     Grade %in% c("AKHS", "PKHS", "OMSN")) %>%
  group_by(Grade) %>%
  dplyr::summarize(Tons = sum(Tons), Count = n())

#------------------------------------------------------------------------------
#
# Determine if a roll is Bev VMI, Eruope, CPD VMI (Capri Sun), CPD Web or
#   CPD Sheet
#

# Read the Beverage VMI List
vmi_rolls <- read.xlsx(file.path(proj_root, "Data", "VMI List.xlsx"),
                       startRow = 2, cols = c(1:6))
vmi_rolls[is.na(vmi_rolls)] <- ""
vmi_rolls <- vmi_rolls %>%
  filter(Status != "D")
vmi_rolls$Caliper <- as.character(vmi_rolls$Caliper)
vmi_rolls <- vmi_rolls %>%
  select(Caliper, Width, Grade)
# Make sure no duplicates
if (any(duplicated(vmi_rolls))) stop("Duplicates in VMI List")
vmi_rolls <- vmi_rolls %>%
  mutate(Bev.VMI = T)

rd_comb <- left_join(rd_comb, vmi_rolls,
                     by = c("Grade" = "Grade",
                            "Caliper" = "Caliper",
                            "Width" = "Width")) %>%
  mutate(Bev.VMI = ifelse(is.na(Bev.VMI), F, T)) %>%
  arrange(Grade, Bev.VMI)
rm(vmi_rolls)

# Identify the Eruope Items
rd_eur <- rd_comb %>%
  filter(str_sub(Plant, 1, 5) == "PLTEU") %>%
  distinct(Mill, Grade, CalDW, Width) %>%
  mutate(Eur = T)

rd_comb <- left_join(rd_comb, rd_eur,
                     by = c("Mill" = "Mill",
                            "Grade" = "Grade",
                            "CalDW" = "CalDW",
                            "Width" = "Width"))
rd_comb <- rd_comb %>%
  mutate(Eur = ifelse(is.na(Eur), F, T))
rm(rd_eur)

# ID Web Rolls by facility 
facility_names <- read_csv(file.path(proj_root, "Data", "FacilityNames.csv"))

# make sure no duplicates
facility_names %>%
  group_by(Facility.ID) %>%
  filter(n() > 1)

cpd_web_plants <- facility_names %>%
  filter(Fac.Type == "CPD Web") %>%
  select(Facility.ID)
rd_cpd_web <- rd_comb %>%
  filter(Plant %in% as_vector(cpd_web_plants)) %>%
  distinct(Mill, Grade, CalDW, Width) %>%
  mutate(CPD.Web = T)

rd_comb <- left_join(rd_comb, rd_cpd_web,
                     by = c("Mill" = "Mill",
                            "Grade" = "Grade",
                            "CalDW" = "CalDW",
                            "Width" = "Width"))
rd_comb <- rd_comb %>%
  mutate(CPD.Web = ifelse(is.na(CPD.Web), F, T))
rm(cpd_web_plants)

# Read Web press consumption for Atlanta
cpd_web_cons <- read.xlsx(file.path(proj_root, "Data", 
                                    "QV SUS WEb Consumption 2015 2016.xlsx"),
                          detectDates = TRUE,
                          cols = 1:14)

cpd_web_cons <- cpd_web_cons %>%
  filter(Plant.Name == "Atlanta") %>%
  mutate(Year = as.character(year(Date))) %>%
  distinct(Year, Plant, Board.Caliper, Board.Width) %>%
  mutate(Plant = paste0("PLT0", Plant),
         CPD.Web.Cons = T,
         Board.Width = as.numeric(Board.Width))

# If a web roll was used anywhere, call it a web roll
rd_comb <- left_join(rd_comb, cpd_web_cons,
                     by = c("Year" = "Year",
                            "Caliper" = "Board.Caliper",
                            "Width" = "Board.Width"))
rd_comb <- rd_comb %>%
  mutate(CPD.Web = ifelse(!is.na(CPD.Web.Cons), T, CPD.Web)) %>%
  select(-CPD.Web.Cons)

rm(cpd_web_cons)

# CPD VMI (Capri Sun)
rd_comb <- rd_comb %>%
  mutate(CPD.VMI = 
           ifelse(Grade == "FC02" & Caliper == "30" & Width == 45.375,
                  TRUE, FALSE)) %>%
  mutate(CPD.Web = ifelse(CPD.VMI, F, CPD.Web))

# Define the Other category
rd_comb <- rd_comb %>%
  mutate(CPD.Sheet = ifelse(Bev.VMI | CPD.Web | Eur | CPD.VMI, F, T))

# Now that we have the types identified, we can get rid of the 
# Plant fields
rd_comb <- rd_comb %>%
  group_by(Year, Mill, Grade, CalDW, Caliper, Dia, Wind, Width,
           Bev.VMI, CPD.VMI, Eur, CPD.Web, CPD.Sheet) %>%
  dplyr::summarize(Tons = sum(Tons), Plant.Count = n()) %>%
  ungroup()

# Create a new field "Prod.Type" that identifies the type.  Give
# preference to Bev.VMI, then CPD.Web, the Europe
rd_comb <- rd_comb %>%
  mutate(Prod.Type = case_when(.$Bev.VMI ~ "Bev.VMI",
                               .$CPD.VMI ~ "CPD.VMI",
                               .$Eur ~ "Europe",
                               .$CPD.Web ~ "CPD.Web",
                               TRUE ~ "CPD.Sheet"))
# Sort
rd_comb$Prod.Type <- 
  factor(rd_comb$Prod.Type, 
         levels = c("Bev.VMI", "CPD.VMI", "CPD.Web", "CPD.Sheet", "Europe"))

#  mutate(Prod.Type = ifelse(Bev.VMI, "Bev.VMI",
#                            ifelse(CPD.Web, "CPD.Web",
#                                   ifelse(Eur, "Europe", "CPD.Sheet"))))

# Plot the 2016 Volume
rd_comb_16_plot <- rd_comb %>%
  filter(Year == "2016")

ggplot(data = rd_comb_16_plot, 
       aes(x = Width, y = Tons)) +
  geom_bar(stat = "identity", width = 0.25, position = "stack",
           aes(fill = Prod.Type), alpha = 0.6) +
  scale_y_continuous(labels = comma) +
  ggtitle("2016 SUS Volume by Width")

rm(rd_comb_16_plot)

# Get the items common to both years
rd_common <- 
  dplyr::intersect(select(filter(rd_comb, Year == "2015"),
                          Mill, Grade, CalDW, Width),
                   select(filter(rd_comb, Year == "2016"),
                          Mill, Grade, CalDW, Width))

rd_common$Common.SKU <- TRUE

# Save this in case we want to use the common SKUs starting from scratch
write_csv(rd_common, file.path(proj_root, "Data", 
                               "RD Common 15 to 16 1463 SKUs.csv"))
write_csv(rd_cpd_web, file.path(proj_root, "Data", "CPD Web SKUs.csv"))

# Ignore the Mill
rd_common_xm <- 
  dplyr::intersect(select(filter(rd_comb, Year == "2015"),
                          Grade, CalDW, Width),
                   select(filter(rd_comb, Year == "2016"),
                          Grade, CalDW, Width))
rd_common_xm$Common.SKUXM <- TRUE

# Ignore the Diameter
rd_common_xd <- 
  dplyr::intersect(select(filter(rd_comb, Year == "2015"),
                          Mill, Grade, Caliper, Wind, Width),
                   select(filter(rd_comb, Year == "2016"),
                          Mill, Grade, Caliper, Wind, Width))
rd_common_xd$Common.SKUXD <- TRUE

# Identify the common items
rd_comb <- 
  left_join(rd_comb, rd_common,
            by = c("Mill" = "Mill",
                   "Grade" = "Grade",
                   "CalDW" = "CalDW",
                   "Width" = "Width")) 
rd_comb <- rd_comb %>%
  mutate(Common.SKU = ifelse(is.na(Common.SKU), FALSE, TRUE))

rd_comb <- 
  left_join(rd_comb, rd_common_xm,
            by = c("Grade" = "Grade",
                   "CalDW" = "CalDW",
                   "Width" = "Width"))
rd_comb <- rd_comb %>%
  mutate(Common.SKUXM = ifelse(is.na(Common.SKUXM), FALSE, TRUE))

rd_comb <-
  left_join(rd_comb, rd_common_xd,
            by = c("Mill" = "Mill",
                   "Grade" = "Grade",
                   "Caliper" = "Caliper",
                   "Wind" = "Wind",
                   "Width" = "Width"))
rd_comb <- rd_comb %>%
  mutate(Common.SKUXD = ifelse(is.na(Common.SKUXD), FALSE, TRUE))

# Add the Die count info.  Need to run the Read Consumption by Die.R script
# if we don't already have the data frames in memoru
if (!exists("dies")) source(file.path(proj_root, "R", 
                                      "Read Consumption by Die.R"))
rd_comb <- 
  left_join(rd_comb, dies,
            by = c("Grade" = "Grade",
                   "Caliper" = "Board.Caliper",
                   "Dia" = "Dia",
                   "Wind" = "Wind",
                   "Width" = "Board.Width"))

# Summary of the raw data SKUs
copy.table(rd_comb %>% 
             group_by(Year, Prod.Type, Common.SKU, Common.SKUXM, 
                      Common.SKUXD) %>%
             dplyr::summarize(SKUs = n(), 
                       Tons = sum(Tons),
                       Dies = sum(Nbr.Dies, na.rm = T)))

# Top x no match in each year
copy.table(rd_comb %>% 
             group_by(Year) %>%
             filter(is.na(Common.SKU)) %>%
             top_n(50, Tons) %>%
             arrange(Year, desc(Tons)))

# Summary of matches by Bev, Eur & CPD.Web
copy.table(rd_comb %>%
             group_by(Year, Common.SKU, Common.SKUXM, Common.SKUXD, 
                      Bev.VMI, Eur, CPD.Web) %>%
             dplyr::summarize(SKUs = n(), Tons = sum(Tons)))

# Duplicate Summary Results
rd_comb %>% group_by(Year, Mill, Grade, CalDW, Width) %>%
  dplyr::summarize(Tons = sum(Tons), Count = n_distinct(Year, Mill, 
                                                 Grade, CalDW, Width)) %>%
  filter(Count > 1)

# Generate a plot of the products
ggplot(data = rd_comb, aes(x = Caliper, y = Width)) +
  geom_jitter(aes(color = Year, shape = Year), size = 2, alpha = 0.5,
              height = 0, width = .1) +
  facet_wrap(~ Grade, scales = "free") +
  scale_color_brewer(palette = "Set1") +
  ggtitle("Comparison of Products (Across all Machines) 2015 vs. 2016")

# Identify the items with no match
rd_16_not_15 <- dplyr::setdiff(select(filter(rd_comb, Year == "2016"),
                                      Mill, Grade, Caliper, CalDW, Width),
                               select(filter(rd_comb, Year == "2015"),
                                      Mill, Grade, Caliper, CalDW, Width))

# Match this to the original 2016 data so we can see locations
rd_16_not_15 <- semi_join(rd_2016, rd_16_not_15,
                          by = c("Mill" = "Mill",
                                 "Grade" = "Grade",
                                 "CalDW" = "CalDW",
                                 "Width" = "Width"))

# And pull in the categories from the rd_comb df
rd_16_not_15 <- left_join(rd_16_not_15, 
                          select(filter(rd_comb, Year == "2016"),
                                 Mill, Grade, CalDW, Width, Bev.VMI, 
                                 Eur, CPD.Web),
                          by = c("Mill" = "Mill",
                                 "Grade" = "Grade",
                                 "CalDW" = "CalDW",
                                 "Width" = "Width"))


# Pull in the die information by plant
rd_2015_dies <- left_join(rd_2015, dies_plants,
                          by = c("Plant" = "Plant",
                                 "Grade" = "Grade",
                                 "Caliper" = "Board.Caliper",
                                 "Dia" = "Dia",
                                 "Wind" = "Wind",
                                 "Width" = "Board.Width"))

rd_2016_dies <- left_join(rd_2016, dies_plants,
                          by = c("Plant" = "Plant",
                                 "Grade" = "Grade",
                                 "Caliper" = "Board.Caliper",
                                 "Dia" = "Dia",
                                 "Wind" = "Wind",
                                 "Width" = "Board.Width"))

rd_16_not_15_die <- left_join(rd_16_not_15, rd_2016_dies,
                              by = c("Mill" = "Mill",
                                     "Plant" = "Plant",
                                     "Grade" = "Grade",
                                     "CalDW" = "CalDW",
                                     "Width" = "Width"))

# Summarize by plant
new_sku_plot_df <- rd_16_not_15 %>% 
  group_by(Year, Plant, Fac.Abrev) %>%
  dplyr::summarize(Tons = sum(Tons), SKUs = n_distinct(Mill, Grade, CalDW, Width)) %>%
  arrange(desc(Tons)) %>% ungroup()

ggplot(data = new_sku_plot_df, aes(x = Fac.Abrev, y = Tons)) +
  geom_bar(stat = "identity", fill = "seagreen") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5)) +
  scale_y_continuous(labels = comma) +
  geom_text(aes(label = SKUs), vjust = "inward") +
  ggtitle("2016 New SKU Tons and Counts by Plant")


# Summarize by the VMI, Eur & CPD.Web categories
rd_16_not_15 %>% group_by(Year, Mill, Grade, CalDW, Width, Bev.VMI, Eur, CPD.Web) %>%
  dplyr::summarize(Tons = sum(Tons), SKUs = n_distinct(Year, Mill, Grade, CalDW, Width)) %>%
  group_by(Bev.VMI, Eur, CPD.Web) %>%
  dplyr::summarize(Tons = sum(Tons), SKUs = n())

copy.table(rd_16_not_15_die %>% filter(CPD.Web | Bev.VMI))

# Clean up
rm(rd_2015, rd_2016, rd_16_not_15, new_sku_plot_df,
   rd_2015_dies, rd_2016_dies, rd_16_not_15_die, rd_common_xd,
   rd_common_xm)

