#------------------------------------------------------------------------------
# 
# Look at the solution costs

# mYear <- "2016"
# mYear <- "2015"
# projscenario <- "Base"
# projscenario <- "BevOnly"
# fpoutput <- file.path("Scenarios", mYear, projscenario)

# Get the Optimization Solutions Summary.csv files
sol_sum <- read_results("sol_summary", mYear, projscenario)
sol_sum <- sol_sum %>%
  filter(Fac %in% facility_list)


sol_sum$Year <- mYear

#min_sol <- sol_sum %>% 
#  group_by(Fac) %>%
#  slice(which.min(Total.Cost)) %>%
#  select(Fac, NbrParents, Total.Cost) %>%
#  rename(Min.Nbr.Parents = NbrParents, Min.Total.Cost = Total.Cost)

#sol_sum <- left_join(sol_sum, min_sol,
#                     by = c("Fac" = "Fac"))

# plot the solutions by first putting the data in long format
sol_sum_plot <- sol_sum %>% 
  select(Fac, NbrParents, Total.Cost, Trim.Cost, Opty.Cost, Handling.Cost, 
         Storage.Cost,
         Inv.Carrying.Cost, Year, Min.Nbr.Parents, Min.Total.Cost) %>%
  gather(key = `Cost Component`, value = Cost, 
         Total.Cost:Inv.Carrying.Cost)

ggplot(sol_sum_plot, 
       aes(x = NbrParents, color = `Cost Component`)) +
  geom_line(aes(y = Cost), size = 1) +
  scale_y_continuous(labels = unit_format(unit = "M", scale = 1e-6)) +
  geom_text(aes(label = as.character(Min.Nbr.Parents),
                x = Min.Nbr.Parents, y = Min.Total.Cost/2), color = "black") +
  geom_segment(aes(x = Min.Nbr.Parents, xend = Min.Nbr.Parents, 
                   y = 0, yend = Min.Total.Cost),
               size = 0.4, color = "black") +
  scale_color_discrete(breaks = c("Total.Cost", "Handling.Cost",
                                  "Inv.Carrying.Cost", "Opty.Cost",
                                  "Storage.Cost", "Trim.Cost")) +
  facet_grid(~ Fac, scales = "free")
  
# Get the minimum for each Mill/Mach
sol_sum <- sol_sum %>%
  group_by(Fac) %>%
  mutate(Rank = rank(Total.Cost))

sol_sum <- sol_sum %>%
  filter(Rank == min(Rank) | Rank == max(Rank))

sol_sum_diff <- sol_sum %>%
  gather(Key, Val, Total.Cost:SS.Tons) %>%
  filter(Key %in% c("Total.Cost", "Exp.OH.Tons")) %>%
  mutate(Val = ifelse(Rank != 1, -Val, Val)) %>%
  group_by(Fac) %>%
  mutate(Nbr = n_distinct(solutionNbr)) %>%
  filter(Nbr > 1) %>%
  ungroup() %>%
  group_by(Key) %>%
  summarize(Val = sum(Val))


#------------------------------------------------------------------------------
#
# Look at actual roll inventory history
#
# Data is from the Bev Roll Inventory Bookmark in the QlikView Batch Inventory
# cube under the Inv Trend Details tab.  Make sure you select Mfg Plan, 
# Material Group, Material and Material Description for the layout.
#
library(tidyverse)
library(scales)
library(openxlsx)
library(stringr)

# First load the optimal solution details
# mYear = "2016"
# mYear = "2015"
# projscenario <- "Base"
# projscenario <- "BevOnly"

sim_parents <- read_results("sim_parents", mYear, projscenario)

# If W6 only
sim_parents <- sim_parents %>% filter(Fac %in% facility_list)

sus_inv <- read.xlsx("Data/Sus Roll Inventory 2016.xlsx", 
                     check.names = TRUE, detectDates = TRUE)

#sus_inv <- read.xlsx("Data/Sus Roll Inventory 2015 2016.xlsx", 
#                          check.names = TRUE, detectDates = TRUE)                 

# Just keep the Tons
# This will get rid of Batches and Cost
#sus_inv <- sus_inv[as.vector(!(sus_inv[1,] == "# Batches" | 
#  sus_inv[1,] == "Cost Val"))]
sus_inv <- sus_inv[as.vector(!(sus_inv[1,] == "# Batches"))]

sus_inv <- sus_inv %>% gather(Range, Tons, 5:ncol(sus_inv))
names(sus_inv)[1:4] <- sus_inv[1,1:4]
sus_inv <- sus_inv %>% filter(`Mfg Plant` != "Mfg Plant") # summary rows

sus_inv <- sus_inv %>%
  select(-`Material Group`) %>%
  rename(Mfg.Plant = `Mfg Plant`, Material.Descr = `Material Description`,
         Value = Tons)
sus_inv <- sus_inv %>% filter(Value != "-")
sus_inv <- sus_inv %>% filter(Material.Descr != "Deleted material")

# Use the length of the Range field to to assign a value type
sus_inv <- sus_inv %>%
  mutate(Val.Type = ifelse(str_length(Range) == 13, "Cost", "Tons"))

sus_inv <- sus_inv %>% filter(!is.na(Material))
sus_inv$Value <- as.numeric(gsub(",", "", sus_inv$Value))
sus_inv$Range <- as.Date(str_sub(sus_inv$Range, 2, 11), "%Y.%m.%d")

sus_inv <- sus_inv %>% 
  mutate(Width = ifelse(str_sub(Material.Descr, 1, 1) == "I",
                        str_sub(Material.Descr, 10, 17),
                        str_sub(Material.Descr, 11, 14)))
sus_inv$Width <- as.numeric(sus_inv$Width)

sus_inv <- sus_inv %>%
  mutate(Caliper = ifelse(str_sub(Material.Descr, 1, 1) == "I",
                          str_sub(Material.Descr, 3, 4),
                          str_sub(Material.Descr, 4, 5)))
sus_inv <- sus_inv %>%
  mutate(Grade = ifelse(str_sub(Material.Descr, 1, 1) == "I",
                        str_sub(Material.Descr, 5, 8),
                        str_sub(Material.Descr, 6, 9)))

sus_inv <- sus_inv %>%
  mutate(Wind = str_sub(Material.Descr, 
                        str_length(Material.Descr), -1))

sus_inv <- sus_inv %>%
  mutate(Dia = ifelse(str_sub(Material.Descr, 1, 1) == "I",
                      str_sub(Material.Descr, 19, 20),
                      str_sub(Material.Descr, 16, 19)))

# Fix the plant names
sus_inv <- sus_inv %>%
  mutate(Mill = ifelse(str_detect(Mfg.Plant, "0031"), "W", "M"))

# Calculate the cost per ton
sus_inv %>% group_by(Val.Type) %>%
  summarize(Total = sum(Value))

# Get rid of the cost info
sus_inv <- sus_inv %>% filter(Val.Type == "Tons")
sus_inv <- sus_inv %>%
  rename(Tons = Value)

# The inventory doesn't have the machine, so we need to strip off 
# the machine from the Fac field in the detail file
sim_parents <- sim_parents %>%
  mutate(Mill = str_sub(Fac, 1, 1))
sim_parents <- sim_parents %>%
  mutate(Wind = str_sub(CalDW, 7, 7))

# The sim results for the individual products corresponds to the
# rows where nbr_rolls = 1.  Also restrict the dat to W6 items
sim_prod <- sim_parents %>%
  filter(nbr_rolls == 1)

# We need to add back the items we modified on the data read
dtf <- data_frame(Mill = c("W", "W", "W"), 
                  Grade = c("AKPG", "AKPG", "AKPG"),
                  Caliper = c("21", "21", "27"),
                  Parent_Width = c(62.875, 64.0625, 47.5),
                  Wind = c("W","W", "W"))

sim_prod <- bind_rows(sim_prod, dtf)                  

# restrict the inventory file to just the items of interest
sus_inv <- 
  semi_join(sus_inv, sim_prod,
            by = c("Mill" = "Mill",
                   "Grade" = "Grade",
                   "Caliper" = "Caliper",
                   "Width" = "Parent_Width",
                   "Wind" = "Wind"))


ggplot(sus_inv, aes(x = Range, y = Tons)) +
  geom_bar(position = "stack", stat = "identity",
           aes(color = Material)) +
  scale_color_discrete(guide = FALSE) +
  scale_y_continuous(labels = comma) +
  scale_x_date(name = "Month End") +
  ggtitle("QlikView Month End Inventory for W6 SKUs")

# Get the total for each month and use this in the cumulative simulation
# plots in 9.Aggregate Sim Plots.R
sus_inv_ttl <- sus_inv %>%
  group_by(Range) %>%
  summarize(On.Hand = sum(Tons)) %>%
  mutate(Scenario = "Q.V. Month End") %>%
  rename(Date = Range)

sus_inv_fy <- sus_inv %>%
  group_by(Mill, Grade, Caliper, Width, Wind) %>%
  summarize(Avg.OH = mean(Tons))

# Join to get QlikView and Sim comparison
inv_fy <- 
  left_join(sus_inv_fy,
            select(sim_prod, Mill, Grade, Caliper, Parent_Width, Wind,
                   cycle_dbr,
                   Sim.Parent.Avg.OH, Sim.Parent.Qty.Fill.Rate),
                   by = c("Mill" = "Mill",
                              "Grade" = "Grade",
                              "Caliper" = "Caliper",
                              "Width" = "Parent_Width",
                              "Wind" = "Wind"))


# Compute the average OH for the full 12 months
sus_inv_fy <- sus_inv %>%
  group_by(Mill, Grade, Caliper, Dia, Wind, Width) %>%
  summarize(Avg.OH.Tons = sum(Tons)/12)

# Find the rolls that are in the solution.  Ignore the diameter

sus_inv_fy %>% 
  ungroup() %>%
  summarize(Avg.OH.Tons = sum(Avg.OH.Tons))


#------------------------------------------------------------------------------
#
# Analyze the W6 scenarios
#
# Start with the raw data, without fixing the diameter info because
# we want to know what diameters we actually used
#
#projscenario <- "BevOnly"
rd_2016 <- read_csv(paste("Data/All Shipments Fixed Orig Dia ", 
                          "2016", ".csv", sep = ""))

vmi_rolls <- read.xlsx("Data/VMI List.xlsx",
                       startRow = 2, cols = c(1:6))
vmi_rolls[is.na(vmi_rolls)] <- ""
vmi_rolls <- vmi_rolls %>%
  filter(Status != "D")
vmi_rolls$Caliper <- as.character(vmi_rolls$Caliper)
vmi_rolls <- vmi_rolls %>%
  select(-Status)

#Identify just the VMI rolls, all Diameters
rd_vmi <- 
  semi_join(rd_2016, vmi_list,
            by = c("Grade" = "Grade",
                   "Cal" = "Caliper",
                   "Width" = "Width"))

rd_2016 <- rd_2016 %>%
  mutate(Cal = as.character(Cal)) 

rd_opt <- 
  semi_join(rd_2016, opt_detail,
            by = c("Mill" = "Fac",
                   "Grade" = "Grade",
                   "Cal" = "Caliper",
                   "Width" = "Prod.Width",
                   "Wind" = "Wind"))

#Just items made on W6 in 2016.  Don't include Diameter
rd_vmi <- 
  semi_join(rd_vmi,
            filter(rd_vmi, Year == 2016 &
                     Mill %in% facility_list),
            by = c("Grade" = "Grade",
                   "Cal" = "Cal",
                   "Width" = "Width",
                   "Wind" = "Wind"))
rd_opt <- 
  semi_join(rd_opt,
            filter(rd_opt, Year == "2016" &
                     Mill %in% facility_list),
            by = c("Mill" = "Mill",
                   "Grade" = "Grade",
                   "Cal" = "Cal",
                   "Width" = "Width",
                   "Wind" = "Wind"))

# Add counts for number of Diam, Winds & Plants
rd_vmi <- rd_vmi %>%
  group_by(Grade, Cal, Width) %>%
  mutate(Nbr.Dia = n_distinct(Dia),
         Nbr.Wind = n_distinct(Wind),
         Nbr.Plants = n_distinct(Plant)) %>%
  ungroup()

rd_opt <- rd_opt %>%
  group_by(Mill, Grade, Cal, Width) %>%
  mutate(Nbr.Dia = n_distinct(Dia),
         Nbr.Wind = n_distinct(Wind),
         Nbr.Plants = n_distinct(Plant)) %>%
  ungroup()

# Summarize Tons by Year, Diam & Wind
rd_vmi <- rd_vmi %>%
  group_by(Year, Grade, Cal, Width, Dia, Wind, Nbr.Dia, 
           Nbr.Wind, Nbr.Plants) %>%
  summarize(Tons = round(sum(Tons),0)) %>%
  mutate(Tons = format(Tons, big.mark = ",")) %>%
  ungroup()

rd_opt <- rd_opt %>%
  group_by(Year, Mill, Grade, Cal, Width, Dia, Wind, Nbr.Dia, 
           Nbr.Wind, Nbr.Plants) %>%
  summarize(Tons = round(sum(Tons),0)) %>%
  mutate(Tons = format(Tons, big.mark = ",")) %>%
  ungroup()


# Spread by Year, Dia
rd_vmi <- rd_vmi %>%
  unite(YDW, Year, Dia) %>%
  spread(YDW, Tons, fill = "")

rd_opt <- rd_opt %>%
  unite(YDW, Year, Dia) %>%
  spread(YDW, Tons, fill = "")

vmi_list <- inner_join(vmi_list, rd_vmi,
                      by = c("Grade" = "Grade",
                             "Caliper" = "Cal", 
                             "Width" = "Width"))

vmi_list <- vmi_list %>%
  arrange(Grade, Caliper, Width)

rd_opt <- rd_opt %>%
  arrange(Mill, Grade, Caliper, Width)

vmi_list[is.na(vmi_list)] <- ""

rd_opt[is.na(rd_opt)] <- ""

# join to opt solution
# The opt_detail comes from the 7.1 Opt Solution Post Process.R file
opt_detail_m <- opt_detail %>%
  filter(Fac %in% facility_list)


# This for the VMI
opt_detail_v <- opt_detail_m %>%
  select(Fac, Grade, Caliper, Prod.Width, Parent.Width, Prod.Trim.Width, 
         Prod.Trim.Tons) %>%
  mutate(Caliper = as.numeric(Caliper))

opt_detail_m <- opt_detail_m %>%
  select(Fac, Grade, Caliper, Wind, Prod.Width, Parent.Width, Prod.Trim.Width, 
         Prod.Trim.Tons) 

vmi_list <- left_join(vmi_list, opt_detail_v,
                      by = c("Grade" = "Grade",
                             "Caliper" = "Caliper",
                             "Width" = "Prod.Width"))

opt_list <- left_join(rd_opt, opt_detail_m,
                      by = c("Mill" = "Fac",
                             "Grade" = "Grade",
                             "Cal" = "Caliper",
                             "Width" = "Prod.Width",
                             "Wind" = "Wind"))


vmi_list <- vmi_list %>%
  select(Fac, Grade, Caliper, Width, Wind, Purpose, Die, Parent.Width, 
         Prod.Trim.Width, Prod.Trim.Tons, everything())
opt_list <- opt_list %>%
  select(Mill, Grade, Cal, Width, Wind, Parent.Width,
         Prod.Trim.Width, Prod.Trim.Tons, everything())

vmi_list <- vmi_list %>%
  mutate(Prod.Trim.Tons = round(Prod.Trim.Tons, 1))
opt_list <- opt_list %>%
  mutate(Prod.Trim.Tons = round(Prod.Trim.Tons, 1))

vmi_elim <- vmi_list %>%
  filter(Prod.Trim.Width > 0) %>%
  mutate(Caliper = as.character(Caliper))

# Match the vmi_elim up with the QlikView On-Hand in the sus_inv table from
# the previous section in this script.
sus_inv_elim <- semi_join(sus_inv, vmi_elim,
                          by = c("Grade" = "Grade",
                                 "Caliper" = "Caliper",
                                 "Width" = "Width",
                                 "Wind" = "Wind"))

ggplot(sus_inv_elim, aes(x = Range, y = Tons)) +
  geom_bar(stat = "identity", aes(fill = paste(Grade, Caliper, Width)),
           color = "black") +
  #scale_fill_discrete(guide = FALSE) +
  #scale_fill_discrete(name = "Grade Cal Width") +
  scale_x_date(name = "") +
  scale_fill_discrete(guide = guide_legend(title = "Grade Caliper Width", 
                                           keyheight = .8,
        ncol = 1))
  #theme(legend.key.size = unit(.035, "npc"))


# Eliminate the "." in the field names for output purposes
vmi_list_out <- vmi_list
names(vmi_list_out) <- gsub("\\.", " ", names(vmi_list_out))
names(vmi_list_out) <- gsub("_", " ", names(vmi_list_out))

opt_list_out <- opt_list
names(opt_list_out) <- gsub("\\.", " ", names(opt_list_out))
names(opt_list_out) <- gsub("_", " ", names(opt_list_out))

copy.table(vmi_list_out)
copy.table(opt_list_out)

#------------------------------------------------------------------------------
#
# Create a list of the detail rolls and their parents, spread by actual 
# diameters shipped

rd_2016 <- read_csv(paste("Data/All Shipments Fixed Orig Dia ", 
                          "2016", ".csv", sep = ""))

rd_2016 <- rd_2016 %>%
  mutate(Cal = as.character(Cal)) 

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
  spread(YD, Tons, fill = "")

rd_opt <- rd_opt %>%
  arrange(Mill, Grade, Cal, Width)

rd_opt[is.na(rd_opt)] <- ""

# join to opt solution to get the parent and trim info
# The opt_detail comes from the 7.1 Opt Solution Post Process.R file
opt_detail_m <- opt_detail %>%
  filter(Fac %in% facility_list)

opt_detail_m <- opt_detail_m %>%
  select(Fac, Grade, Caliper, Wind, Prod.Width, Parent.Width, Prod.Trim.Width, 
         Prod.Trim.Tons) 

opt_list <- left_join(rd_opt, opt_detail_m,
                      by = c("Mill" = "Fac",
                             "Grade" = "Grade",
                             "Cal" = "Caliper",
                             "Width" = "Prod.Width",
                             "Wind" = "Wind"))

opt_list <- opt_list %>%
  select(Mill, Grade, Cal, Width, Wind, Parent.Width,
         Prod.Trim.Width, Prod.Trim.Tons, everything())

opt_list <- opt_list %>%
  mutate(Prod.Trim.Tons = round(Prod.Trim.Tons, 1))

if (projscenario == "Europe") {
  opt_list <- opt_list %>%
    mutate(Width = round(Width * 25.4, 0),
           Parent.Width = round(Parent.Width * 25.4, 0),
           Prod.Trim.Width = round(Prod.Trim.Width * 25.4, 0))
}

# Eliminate the "." in the field names for output purposes
opt_list_out <- opt_list
names(opt_list_out) <- gsub("\\.", " ", names(opt_list_out))
names(opt_list_out) <- gsub("_", " ", names(opt_list_out))

copy.table(opt_list_out)

# Generate plot of trim width
ggplot(data = opt_list, aes(x = Prod.Trim.Width,
                            y = Prod.Trim.Tons)) +
  geom_histogram(stat = "identity", geom = "bar",
                 aes(fill = Grade)) +
  ggtitle("Europe Roll Consolidation Trim Loss") +
  scale_x_continuous(name = "Trim Loss (mm)") +
  scale_y_continuous(name = "Total Trim Tons")
  

