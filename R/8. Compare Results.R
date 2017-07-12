#------------------------------------------------------------------------------
#
# Deep dive into 2016 data
#
library(plyr)
library(tidyverse)
library(stringr)
library(openxlsx)
library(scales)
library(RColorBrewer)
library(gridExtra)

# Read in the 2016 product data
rd_2016 <- read_csv(paste("Data/All Shipments Fixed ", 
                         as.numeric(mYear), ".csv", sep = ""))

rd_2016 <- rd_2016 %>%
  group_by(Year, Mill, Grade, CalDW, Width, Plant, Fac.Abrev) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()
rd_2016$Year <- as.character(rd_2016$Year)
rd_2016 <- rd_2016 %>%
  separate(CalDW, c("Caliper", "Dia", "Wind"), remove = FALSE)

# Read the VMI List
vmi_rolls <- read.xlsx("Data/VMI List.xlsx",
                       startRow = 2, cols = c(1:6))
vmi_rolls[is.na(vmi_rolls)] <- ""
vmi_rolls <- vmi_rolls %>%
  filter(Status != "D")
vmi_rolls$Caliper <- as.character(vmi_rolls$Caliper)
vmi_rolls <- vmi_rolls %>%
  select(Caliper, WIdth, Grade)
vmi_rolls <- vmi_rolls %>%
  mutate(Bev.VMI = 1)

rd_2016 <- left_join(rd_2016, vmi_rolls,
                        by = c("Grade" = "Grade",
                               "Caliper" = "Caliper",
                               "Width" = "Width")) %>%
  mutate(Bev.VMI = ifelse(is.na(Bev.VMI), F, T)) %>%
  arrange(Grade, Bev.VMI)

rd_2016 <- rd_2016 %>%
  mutate(Prod.Type = ifelse(Bev.VMI, "Bev.VMI",
                            ifelse(I.M == "M", "Europe", 
                                   "CPD")))

mt <- rd_2016 %>%
  group_by(Prod.Type) %>%
  summarize(Tons = sum(Tons),
            Prods = n_distinct(Mill, Grade, CalDW, Width)) %>%
  mutate(Tons = round(Tons,0))
grid.table(format(mt, big.mark = ",", nsmall = 0, justify = "right"), 
           rows = NULL)

# Note that virtually ALL the Europe tons ship to PLT00EUR

ggplot(data = rd_2016, aes(x = reorder(Grade, -Tons, sum), y = Tons)) +
  geom_bar(aes(fill = Prod.Type), stat = "identity") +
  scale_fill_brewer(palette = "Set1", name = "Div") +
  scale_y_continuous(labels = comma) +
  xlab("") +
  #theme_grey(base_size = 9) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1.0),
        text = element_text(size = 9),
        legend.key.size = unit(1, "strheight", "1"))
  
# Look at Caliper detial for AKOS
ggplot(data = unite(filter(rd_2016, Grade == "AKOS"), 
                    Grade, CalDW, Mill),
       aes(x = Grade, y = Tons)) +
  geom_bar(aes(fill = Prod.Type), stat = "identity") +
  scale_fill_brewer(palette = "Set1", name = "Div") +
  scale_y_continuous(labels = comma) +
  xlab("") +
  #theme_grey(base_size = 9) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1.0),
        text = element_text(size = 9),
        legend.key.size = unit(1, "strheight", "1"))


my_tf_colors <- c("darkseagreen", "blue")
names(my_tf_colors) <- levels(factor(rd_2016$Bev.VMI))
ggplot(data = summarize(group_by(rd_2016, Mill, Grade, CalDW, Width, 
                                 Prod.Type), Tons = sum(Tons)),
       aes(x = reorder(paste0(Mill, Grade, CalDW, Width, Prod.Type), -Tons, sum), 
           y = Tons)) +
  geom_bar(aes(fill = factor(Prod.Type)), stat = "identity", width = 3) +
  scale_y_log10(labels = comma) +
  scale_fill_brewer(palette = "Set1") +
  xlab("Products") +
  scale_x_discrete(breaks = NULL) +
  ggtitle("Total Volume = 1,111,340 Tons and 2,418 Products")

ggplot(data = filter(rd_2016, Bev.VMI == T),
       aes(x = reorder(paste0(Mill, Grade, CalDW, Width), -Tons), y = Tons)) +
  geom_bar(aes(fill = Bev.VMI), stat = "identity", width = 1) +
  scale_y_log10(labels = comma) +
  xlab("VMI Products") +
  #scale_x_discrete(breaks = NULL) +
  ggtitle("Total Beverage VMI volume = 484,450 Tons (50%) and 126 `Products`")



#------------------------------------------------------------------------------
#
# Compare Original MOdel (Dylan's) with Revised Model
# Data: 2016
# Policy: Normal distribution for all
# 
library(plyr)
library(tidyverse)
library(stringr)
# Defint the file paths and list of file names to read in
fp_orig <- file.path("Scenarios", "2016", "Dylan")
fp_rev <- file.path("Scenarios", "2016", "Base")

fl_orig <- list.files(fp_orig, ".SelectionAssignments.")
fl_orig <- data_frame(files = file.path(fp_orig, fl_orig))
orig_fac_list <- data_frame(facs = str_sub(fl_orig_opt_assign_list, 1, 2))
fl_orig_df <- bind_cols(fl_orig, orig_fac_list)

opt_orig <- adply(fl_orig_df, 1, function(x) {
  #print(x[2])
  t <- read_csv(file = as.character(x[1]))
  mutate(t, Fac = as.character(x[2]))
  })

opt_orig <- opt_orig %>%
  rename(Nbr.Parents = X..Parents) %>%
  select(-files, -facs, -X1) %>%
  mutate(CalDW = paste(str_sub(Caliper, 4, 5), str_sub(Caliper, 1, 2),
                        str_sub(Caliper, 7, 7), sep = "-"))

opt_orig <- opt_orig %>%
  group_by(Fac, Grade, CalDW, Parent) %>%
  mutate(Parent.Nbr.Subs = n()) %>%
  ungroup() %>%
  rename(Prod.Width = Width, Parent.Width = Parent) %>%
  arrange(Fac, Grade, CalDW, Parent.Width, Prod.Width) %>%
  select(Fac, Grade, CalDW, Parent.Width, Prod.Width, everything())

fl_rev <- list.files(fp_rev, ".Sim Results.")
fl_rev <- data.frame(files = file.path(fp_rev, fl_rev))
opt_rev <- adply(fl_rev, 1, function(x) {
  print(as.character(x[[1]]))
  read_csv(file = as.character(x[[1]]))
})

names(opt_rev)[2] <- "Fac"

opt_rev <- opt_rev %>%
  select(Fac, Grade, CalDW, Sol.Num.Parents, Sol.Trim.Tons, Parent.Nbr.Subs,
         Parent.Width, Prod.Width) %>%
  ungroup() %>%
  arrange(Fac, Grade, CalDW, Parent.Width, Prod.Width)

# Products in Orig not in Rev
no_prod <- setdiff(distinct(opt_orig, Fac, Grade, CalDW, Prod.Width),
        distinct(opt_rev, Fac, Grade, CalDW, Prod.Width))

#------------------------------------------------------------------------------
#
# Compare 2016 vs 2015

library(plyr)
library(tidyverse)
library(readr)
library(openxlsx)
library(stringr)
library(scales)
library(RColorBrewer)
basepath <- file.path("C:","Users", "SchroedR", "Documents",
                      "Projects", "RollAssortmentOptimization")
# Load the various functions
source(file.path(basepath, "R", "0.Functions.R"))


# If it doesn't exist: 
#comp15v16 <- "Base"
comp15v16 <- "Common1463"

det2015 <- load_opt_detail("opt_detail", "2015", comp15v16)
det2016 <- load_opt_detail("opt_detail", "2016", comp15v16)

det2015$Year <- "2015"
det2016$Year <- "2016"

detail <- bind_rows(det2015, det2016)
detail <- detail %>%
  separate(CalDW, c("Cal", "Dia", "Wind"), remove = F)

# get the Division - use the raw data and the VMI files
# First read the VMI List
vmi_rolls <- read.xlsx("Data/VMI List.xlsx",
                       startRow = 2, cols = c(1:3))
vmi_rolls$Caliper <- as.character(vmi_rolls$Caliper)
vmi_rolls <- vmi_rolls %>%
  mutate(Bev.VMI = TRUE)
detail <- 
  left_join(detail, vmi_rolls,
            by = c("Grade" = "Grade",
                   "Caliper" = "Caliper",
                   "Prod.Width" = "Width"))
rm(vmi_rolls)

# Forget trying to identify just the Europe rolls for now
#rd_comb <- rd_comb %>% 
#  filter(Plant == "PLT00EUR") %>%
#  distinct(Year, Grade, CalDW, Width, Plant) %>%
#  mutate(Year = as.character(Year))
#detail <- 
#  left_join(detail, rd_comb,
#            by = c("Year" = "Year",
#                   "Grade" = "Grade",
#                   "CalDW" = "CalDW",
#                   "Parent.Width" = "Width"))
#rm(rd_comb)  

detail <- detail %>%
  mutate(Div = ifelse(Bev.VMI, "Bev.VMI", NA)) 
#detail <- detail %>%
#  mutate(Div = ifelse((is.na(Div) & Plant == "PLT00EUR"), "EUR", Div))
#detail <- detail %>%
#  mutate(Div = ifelse(is.na(Div), "CPD", Div))

# Get the Optimization Solutions Summary.csv files
sol_sum2015 <- load_opt_detail("sol_summary", "2015", comp15v16)
sol_sum2016 <- load_opt_detail("sol_summary", "2016", comp15v16)
sol_sum2015$Year <- "2015"
sol_sum2016$Year <- "2016"
sol_sum <- bind_rows(sol_sum2015, sol_sum2016)


# get the data in long format for easier plotting
# First compute the minimum and join it back to the sol_sum df
min_sol <- sol_sum %>% 
  group_by(Year, Fac) %>%
  slice(which.min(Total.Cost)) %>%
  select(Fac, NbrParents, Total.Cost) %>%
  rename(Min.Nbr.Parents = NbrParents, Min.Total.Cost = Total.Cost)
sol_sum <- left_join(sol_sum, min_sol,
                     by = c("Fac" = "Fac",
                            "Year" = "Year"))
sol_sum_plot <- sol_sum %>% 
  select(Fac, NbrParents, Total.Cost, Trim.Cost, Handling.Cost, Storage.Cost,
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
                                  "Inv.Carrying.Cost", 
                                  "Storage.Cost", "Trim.Cost")) +
  facet_grid(Fac ~ Year, scales = "free")


# Tons summary: not sure why slightly different from what I get in the
# 7.Summarize.R script?
detail %>%
  #group_by(Year, Parent.Type) %>%
  group_by(Year) %>%
  summarize(Nbr.Prods = n(),
            Prod.Gross.Tons = sum(Prod.Gross.Tons),
            Prod.DMD.Tons = sum(Prod.DMD.Tons),
            Nbr.Parents = n_distinct(Fac, Grade, CalDW, Parent.Width)) %>%
  ungroup()

# Generate a plot of the products
ggplot(data = detail, aes(x = Caliper, y = Prod.Width)) +
  geom_jitter(aes(color = Year, shape = Year), size = 2, alpha = 0.5,
              height = 0, width = .1) +
  facet_wrap(~ Grade, scales = "free") +
  scale_color_brewer(palette = "Set1") +
  ggtitle("Comparison of Products (Across all Machines)")

# How many products are the same from year to year
# I think we are better off using the raw data files above
products <- detail %>%
  select(Year, Fac, Grade, CalDW, Prod.Width, Prod.DMD.Tons)

products %>%
  group_by(Year) %>%
  summarize(Count = n(),
            Prod.DMD.Tons = sum(Prod.DMD.Tons))

prod_common <- dplyr::intersect(select(filter(products, Year == "2015"), 
                                       Fac, Grade, CalDW, Prod.Width),
                                select(filter(products, Year == "2016"),
                                       Fac, Grade, CalDW, Prod.Width))
prod_common  %>%
  summarize(Count = n())

# Use this to generate a 2015 and 2016 run with common items


# Use this to compare the products from year to year
# First look at where the same item ran on different machines
# pyf is the product by year & facility combination
pyf <- products %>% 
  mutate(Tons = round(Prod.DMD.Tons, 2)) %>% 
  select(-Prod.DMD.Tons) %>%
  unite(YF, Year, Fac) %>%
  spread(YF, Tons, fill = -1)

# Categorize the items
pyf <- pyf %>%
  mutate(
    Comp.Type =
      ifelse (
        pyf$`2015_M1` * pyf$`2016_M1` > 0 &
          pyf$`2015_M2` * pyf$`2016_M2` > 0 &
          pyf$`2015_W6` * pyf$`2016_W6` > 0 &
          pyf$`2015_W7` * pyf$`2016_W7` > 0,
        "Same",
        ifelse((
          pyf$`2015_M1` > 0 |
            pyf$`2015_M2` > 0 | pyf$`2015_W6` > 0 |
            pyf$`2015_W7` > 0
        ) &
          (
            pyf$`2016_M1` > 0 |
              pyf$`2016_M2` > 0 | pyf$`2016_W6` > 0 |
              pyf$`2016_W7` > 0
          ),
        "Ch Mach",
        ifelse((
          pyf$`2015_M1` > 0 |
            pyf$`2015_M2` > 0 | pyf$`2015_W6` > 0 |
            pyf$`2015_W7` > 0
        ) &
          !(
            pyf$`2016_M1` > 0 |
              pyf$`2016_M2` > 0 | pyf$`2016_W6` > 0 |
              pyf$`2016_W7` > 0
          ),
        "Disc",
        ifelse(!(
          pyf$`2015_M1` > 0 |
            pyf$`2015_M2` > 0 | pyf$`2015_W6` > 0 |
            pyf$`2015_W7` > 0
        ) &
          (
            pyf$`2016_M1` > 0 |
              pyf$`2016_M2` > 0 | pyf$`2016_W6` > 0 |
              pyf$`2016_W7` > 0
          ),
        "New", NA)
        )
        )
      )
  )


copy.table(pyf)
# Save this to the `Products 2015 to 2016.xlsx`` file

# Now look at where the diameter might have changed
# pyd is the product by year & facility combination
pyd <- products %>% 
  mutate(Tons = round(Prod.DMD.Tons, 2)) %>% 
  separate(CalDW, c("Cal", "Dia", "Wind")) %>%
  select(-Prod.DMD.Tons) %>%
  unite(YD, Year, Dia) %>%
  spread(YD, Tons, fill = -1)

pyd <- pyd %>%
  mutate(
    Comp.Type =
      ifelse (
        pyd$`2015_72` * pyd$`2016_72` > 0 &
          pyd$`2015_77` * pyd$`2016_77` > 0 &
          pyd$`2015_81` * pyd$`2016_81` > 0 &
          pyd$`2015_84` * pyd$`2016_84` > 0,
        "Same",
        ifelse((
          pyd$`2015_72` > 0 |
            pyd$`2015_77` > 0 | pyd$`2015_81` > 0 |
            pyd$`2015_84` > 0
        ) &
          (
            pyd$`2016_72` > 0 |
              pyd$`2016_77` > 0 | pyd$`2016_81` > 0 |
              pyd$`2016_84` > 0
          ),
        "Ch Diam",
        ifelse((
          pyd$`2015_72` > 0 |
            pyd$`2015_77` > 0 | pyd$`2015_81` > 0 |
            pyd$`2015_84` > 0
        ) &
          !(
            pyd$`2016_72` > 0 |
              pyd$`2016_77` > 0 | pyd$`2016_81` > 0 |
              pyd$`2016_84` > 0
          ),
        "Disc",
        ifelse(!(
          pyd$`2015_72` > 0 |
            pyd$`2015_77` > 0 | pyd$`2015_81` > 0 |
            pyd$`2015_84` > 0
        ) &
          (
            pyd$`2016_72` > 0 |
              pyd$`2016_77` > 0 | pyd$`2016_81` > 0 |
              pyd$`2016_84` > 0
          ),
        "New", NA)
        )
        )
      )
  )


copy.table(pyd)
# Save this to the `Products 2015 to 2016.xlsx`` file

# Summarize the parent rolls
parents <- detail %>% 
  group_by(Year, Fac, Grade, CalDW, Parent.Width, Parent.DMD,
           Parent.Nbr.Subs, Parent.DMD.Count, 
           #Sim.Parent.dclass, 
           Sim.Parent.Avg.OH, Sim.Parent.Qty.Fill.Rate, 
           Parent.Distrib) %>%
  summarize(Trim.Loss = sum(Prod.Trim.Tons)) %>%
  mutate(Sim.Parent.Avg.OH.DOH = 365*Sim.Parent.Avg.OH/Parent.DMD) %>%
  ungroup()

parents <- parents %>%
  separate(CalDW, c("Caliper", "Dia", "Wind"), sep = "-", remove = FALSE)


ggplot(data = parents, aes(x = Caliper, y = Parent.Width)) +
  geom_jitter(aes(color = Year, shape = Year), size = 2, alpha = 0.5,
  height = 0, width = .2)+
  facet_wrap(~ Grade, scales = "free") +
  scale_color_brewer(palette = "Set1")

max_DOH <- 55
min_FR <- .85

parents <- parents %>%
  mutate(Parent.Type = 
           ifelse(Parent.Distrib == "MTO",
                              "Exclude - Slow", 
                              ifelse(Sim.Parent.Avg.OH.DOH > max_DOH | 
                                        Sim.Parent.Qty.Fill.Rate < min_FR, 
                                      "MTO", "MTS")))


# Summary of total solution
parents %>%
  group_by(Year) %>%
  summarize(
    Sub.Rolls = sum(Parent.Nbr.Subs),
    Parent.Rolls = n(),
    Annual.Tons = sum(Parent.DMD),
    Avg.OH = sum(Sim.Parent.Avg.OH),
    Total.DOH = 365 * sum(Sim.Parent.Avg.OH) / sum(Parent.DMD),
    Trim.Tons = sum(Trim.Loss),
    Fill.Rate = sum(Parent.DMD * Sim.Parent.Qty.Fill.Rate) /
      sum(Parent.DMD)
  )


# Summary with groupings
parents %>%
  group_by(Year, Parent.Type) %>%
  summarize(
    Sub.Rolls = sum(Parent.Nbr.Subs),
    Nbr.Rolls = n(),
    Annual.Tons = sum(Parent.DMD),
    Avg.OH = sum(Sim.Parent.Avg.OH),
    Total.DOH = 365 * sum(Sim.Parent.Avg.OH) / sum(Parent.DMD),
    Trim.Tons = sum(Trim.Loss),
    Fill.Rate = sum(Parent.DMD * Sim.Parent.Qty.Fill.Rate) /
      sum(Parent.DMD)
  ) 

parents <- parents %>% 
  ungroup() %>%
  group_by(Fac, Grade, CalDW, Parent.Width) %>%
  mutate(Flag = n())
  

common <- dplyr::intersect(select(filter(parents, Year == "2015"), 
                                  Fac, Grade, CalDW, Parent.Width),
                           select(filter(parents, Year == "2016"),
                                  Fac, Grade, CalDW, Parent.Width))

common_mts <- dplyr::intersect(select(filter(parents, Year == "2015" & 
                                        Parent.Type == "MTS"),
                               Fac, Grade, CalDW, Parent.Width, Parent.Type),
                        select(filter(parents, Year == "2016" &
                                        Parent.Type == "MTS"),
                               Fac, Grade, CalDW, Parent.Width, Parent.Type))

common <- inner_join(common, parents)
common_mts <- inner_join(common_mts, parents)

common %>% group_by(Year) %>%
  summarize(Parents = n(), 
  Sub.Rolls = sum(Parent.Nbr.Subs),
  Nbr.Rolls = n(),
  Annual.Tons = sum(Parent.DMD),
  Avg.OH = sum(Sim.Parent.Avg.OH),
  Total.DOH = 365 * sum(Sim.Parent.Avg.OH) / sum(Parent.DMD),
  Trim.Tons = sum(Trim.Loss),
  Fill.Rate = sum(Parent.DMD * Sim.Parent.Qty.Fill.Rate) /
    sum(Parent.DMD)
) 


common_mts %>% group_by(Year) %>%
  summarize(Parents = n(), tons = sum(Parent.DMD))

#------------------------------------------------------------------------------
#
# Take the optimal set of parents from 2015 and determine how they would have
# covered the 2016 demand
#
library(plyr)
library(tidyverse)
library(stringr)
library(scales)

basepath <- file.path("C:","Users", "SchroedR", "Documents",
                      "Projects", "RollAssortmentOptimization")

# Load the various functions
source(file.path(basepath, "R", "0.Functions.R"))
source(file.path(basepath, "R", "0.FuncDemandClassification.R"))

# Combine the files with details from the optimal run for each facility

# mYear = "2015"
# projscenario <- "Base"
# projscenario <- "Common1463"
# fpoutput <- file.path("Scenarios", mYear, projscenario)

# Use the function from 7.Summarize Results to pull in the results from 2015
opt_det_comp <- load_opt_detail("opt_detail", "2015", projscenario)

# Make sure no duplicates
opt_det_comp %>%
  group_by(Fac, Grade, CalDW, Prod.Width) %>%
  summarize(mcount = n()) %>%
  filter(mcount > 1)

# Reduce the number of columns 
opt_det_comp <- opt_det_comp %>%
  select(Fac, Grade, CalDW, Parent.Width, 
         Parent.Nbr.Subs, Parent.DMD,
         Prod.Width, Prod.Trim.Width, Prod.Trim.Fact,
         Prod.Trim.Tons, Prod.DMD.Tons,
         Sim.Parent.Avg.OH, Sim.Parent.Qty.Fill.Rate,
         Parent.Distrib)

opt_det_comp %>%
  summarize(Nbr.Prods = n(),
            Prod.DMD.Tons = sum(Prod.DMD.Tons),
            Nbr.Parents = n_distinct(Fac, Grade, CalDW, Parent.Width))

#copy.table(opt_det_comp)

parents2015 <- opt_det_comp %>%
  select(Fac, Grade, CalDW, Parent.Width, Parent.DMD, Parent.Distrib,
         Sim.Parent.Avg.OH, Sim.Parent.Qty.Fill.Rate) %>%
  distinct() %>%
  mutate(Roll.Type = "Parent15",
         Year = "2015",
         Parent.Width = as.character(Parent.Width),
         Sim.Parent.Avg.OH.DOH = 365*Sim.Parent.Avg.OH/Parent.DMD) %>%
  rename(Width = Parent.Width,
         Tons = Parent.DMD,
         Mill = Fac)

max_DOH <- 55
min_FR <- .85

parents2015 <- parents2015 %>%
  mutate(Parent.Type = 
           ifelse(Parent.Distrib == "MTO",
                  "Exclude - Slow", 
                  ifelse(Sim.Parent.Avg.OH.DOH > max_DOH | 
                           Sim.Parent.Qty.Fill.Rate < min_FR, 
                         "MTO", "MTS")))

#copy.table(parents2015)

# Read in the 2016 product data
rd2016 <- read_csv(paste("Data/All Shipments Fixed ", 
                             "2016", ".csv", sep = ""))

if (projscenario == "Common1463") {
  rd2016 <- 
    semi_join(rd2016, rd_common,
              by = c("Mill" = "Mill",
                     "Grade" = "Grade",
                     "CalDW" = "CalDW",
                     "Width" = "Width"))
}

# Need the caliper field to join with the VMI list
rd2016 <- rd2016 %>%
  mutate(Caliper = str_sub(CalDW, 1, 2),
         Width = as.character(Width))

# Identify Bev.VMI items
# Read the VMI List
vmi_rolls <- read.xlsx("Data/VMI List.xlsx",
                       startRow = 2, cols = c(1:3))

vmi_rolls <- vmi_rolls %>%
  mutate(Caliper = as.character(Caliper),
         Width = as.character(Width),
         Bev.VMI = 1)

rd2016 <- left_join(rd2016, vmi_rolls,
                    by = c("Grade" = "Grade",
                           "Caliper" = "Caliper",
                           "Width" = "Width")) %>%
  mutate(Bev.VMI = ifelse(is.na(Bev.VMI), F, T)) %>%
  arrange(Grade, Bev.VMI)


# Run the data through the demand classifier
dmd_aggreg_period <- "weekly"
# Parameters for the demand classification
dmd_parameters <- c(nbr_dmd = 5, # Extremely Slow nbr demands
                    dmd_mean = 5, # Extremely Small demand mean
                    outlier = 10, 
                    intermit = 1.32, # Intermittency demand interval
                    variability = 4, # non-zero dmd std dev
                    dispersion = 0.49) # CV^2, non inter:smooth, 

# Combine the mill the grade, and summarize over the plants
rd2016_stats <- rd2016 %>% 
  unite(Grade, Mill, Grade) %>%
  group_by(Date, Grade, CalDW, Width) %>% # sum over the Plant
  summarize(Tons = sum(Tons)) %>%
  separate(CalDW, c("Caliper", "Diam", "Wind"), sep = "-") %>%
  ungroup()

start_date <- min(rd2016_stats$Date)
end_date <- max(rd2016_stats$Date)
num_days <- as.numeric(end_date - start_date)

rd2016_stats <- demand_class(rd2016_stats, 7, start_date, end_date,
                            dmd_aggreg_period, dmd_parameters)

rd2016_stats <- rd2016_stats %>%
  unite(CalDW, Caliper, Diam, Wind, sep = "-") %>%
  separate(Grade, c("Mill", "Grade"))

rd2016_stats <- rd2016_stats %>%
  select(Mill, Grade, CalDW, Width, nz_dmd_count, dclass, distrib)

rd2016 <- left_join(rd2016, rd2016_stats,
                    by = c("Mill" = "Mill",
                           "Grade" = "Grade",
                           "CalDW" = "CalDW",
                           "Width" = "Width"))

prods2016 <- rd2016 %>%
  mutate(Div = ifelse(Facility.Name == "GPI EUROPE", "EUR", 
                      ifelse(Bev.VMI, "Bev.VMI", NA))) %>%
  group_by(Year, Mill, Grade, CalDW, Div, Width, dclass, distrib) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup() %>%
  mutate(Roll.Type = "Prod16",
         Year = as.character(Year))

#copy.table(prods2016)

# Combine the 2015 Parents with the 2016 product demands
combined <- bind_rows(parents2015, prods2016) %>%
  select(Mill, Grade, CalDW, Roll.Type, Parent.Type, Year, Width, Tons, Div, 
         dclass, distrib) %>%
  arrange(Mill, Grade, CalDW, Width, desc(Year)) %>%
  mutate(Trim = NA,
         Trim.Fact = NA,
         Gross.Tons = NA,
         Trim.Loss = NA,
         Width = as.numeric(Width))

parent_eval <- function(x) {
  for (r in 1:nrow(x)) {
    if (x$Roll.Type[r] == "Prod16") {
      nrow <- match("Parent15", x$Roll.Type[r:nrow(x)], nomatch = -1)
      if (nrow > 0) {
        x$Trim[r] <- x$Width[r + nrow - 1] - x$Width[r]
        x$Trim.Fact[r] <- 
          1 + (x$Width[r + nrow - 1] %% x$Width[r])/x$Width[r]
        x$Gross.Tons[r] <- x$Trim.Fact[r] * x$Tons[r]
        x$Trim.Loss[r] <- x$Gross.Tons[r] - x$Tons[r]
      } else (x$Trim[r] <- -1)
    }
  }
  return(x)
}

combinedALL <- ddply(combined, c("Mill", "Grade", "CalDW"), parent_eval)
combinedMTS <- ddply(filter(combined, 
                            !(Roll.Type == "Parent15" & Parent.Type != "MTS")), 
                            c("Mill", "Grade", "CalDW"), parent_eval)

combinedALL %>% group_by(Div) %>%
  summarize(Tons = sum(Trim.Loss, na.rm = TRUE))
combinedMTS %>% group_by(Div) %>%
  summarize(Tons = sum(Trim.Loss, na.rm = TRUE))

combinedALL[is.na(combinedALL)] <- ""
combinedMTS[is.na(combinedMTS)] <- ""

copy.table(combinedALL)
#copy.table(combinedMTS)

combinedALL %>%
  filter(Roll.Type == "Prod16") %>%
  #filter(distrib != "MTO") %>%
  mutate(Trim = as.numeric(Trim),
         Trim.Category = ifelse(Trim == -1, "Uncov",
                                ifelse(Trim <= 3.0, "<= 3", "> 3"))) %>%
  group_by(Trim.Category) %>%
  #group_by(Div, Trim.Category) %>%
  #group_by(Div) %>%
  summarize(Tons2016 = sum(Tons),
            Trim.Loss = sum(as.numeric(Trim.Loss), na.rm = T))

#combinedMTS %>%
#  filter(Roll.Type == "Prod16") %>%
#  #filter(distrib != "MTO") %>%
#  mutate(Trim = as.numeric(Trim),
#         Trim.Category = ifelse(Trim == -1, "Uncov",
#                                ifelse(Trim <= 3.0, "<= 3", "> 3"))) %>%
#  group_by(Trim.Category) %>%
#  summarize(Tons2016 = sum(Tons),
#            Trim.Loss = sum(as.numeric(Trim.Loss), na.rm = T))


#------------------------------------------------------------------------------
#
# How much trim if we standardize widths to the half inch?
#
library(plyr)
library(tidyverse)
library(stringr)
library(scales)
library(openxlsx)

basepath <- file.path("C:","Users", "SchroedR", "Documents",
                      "Projects", "RollAssortmentOptimization")

# Load the various functions
source(file.path(basepath, "R", "0.Functions.R"))

if (!exists("rawDataAll")) {
  source("R/3.Read Data.R")
}

rd <- rawDataAll %>%
  group_by(Mill, Grade, CalDW, Width) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

rd <- rd %>%
  mutate(New.Width = ceiling(Width * 2)/2)

rd <- rd %>%
  mutate(New.Tons = Tons * New.Width / Width)

rd %>% summarize(SKU.Count = n_distinct(Mill, Grade, CalDW, New.Width),
                 Trim = sum(New.Tons - Tons))
