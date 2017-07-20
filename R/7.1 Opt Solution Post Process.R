#------------------------------------------------------------------------------
#
# Read in the Optimal Solution details, add various product details
# and create the optimal parent summary file.
# 
#library(plyr)
library(tidyverse)
library(scales)
library(stringr)
library(openxlsx)

basepath <- file.path("C:","Users", "SchroedR", "Documents",
                      "Projects", "RollAssortmentOptimization")
proj_root <- file.path("C:", "Users", "SchroedR", "Documents",
                       "Projects", "SUS Roll Consolidation") 

Sys.setenv("R_ZIPCMD" = 
             "C:/RBuildtools/3.4/bin/zip.exe") # for openxlsx writing

# Load the various functions
source(file.path(basepath, "R", "0.Functions.R"))
source(file.path(proj_root, "R", "FuncReadResults.R"))

source(file.path(proj_root, "R", "2.User Inputs.R"))

# Remove loaded packages
#lapply(paste('package:',names(sessionInfo()$otherPkgs),sep=""),
#   detach,character.only=TRUE,unload=TRUE)

# Read the optimal solution details for all machines
opt_detail <- read_results("opt_detail", mYear, projscenario)

# Filter for just the machines of interest.
opt_detail <- opt_detail %>% filter(Fac %in% facility_list)

# Create an index that ranks each product under it's parent
opt_detail <- opt_detail %>%
  group_by(Fac, Grade, CalDW, Parent.Width) %>%
  mutate(Parent.Prod.IDX = dense_rank(Prod.Width)) %>%
  ungroup()

# Plot the total volume by facility/machine (and Grade)
fac_total <- opt_detail %>%  # Use for bar total values
  group_by(Fac) %>%
  summarize(Fac.Tons = sum(Prod.DMD.Tons)) 

ggplot(opt_detail, aes(x = Fac, y = Prod.DMD.Tons)) +
  geom_bar(stat = "identity", aes(fill = Grade)) +
  ggtitle("2016 Beverage VMI Shipments by Machine") +
  scale_y_continuous(name = "Tons", labels = comma,
                     limits = c(0, 400000)) +
  scale_x_discrete(name = "Mill/Machine") +
  geom_text(data = fac_total, vjust = -.5,
            aes(x = Fac, y = Fac.Tons, 
                label = comma(round(Fac.Tons, 0))), position = "stack")
rm(fac_total) 

# First check total tons
opt_detail %>% 
  summarize(Tons = sum(Prod.DMD.Tons), 
            SKUs = n_distinct(Fac, Grade, CalDW, Prod.Width))


# Pull in the Prod.Type from the rawDataAll data frame
if (!exists("rawDataAll")) source(file.path(proj_root, "R", "3.Read Data.R"))

# Make sure we have the same number of products in rd_comb for mYear as we 
# do in opt_detail

opt_detail <- 
  left_join(opt_detail,
            distinct(rawDataAll,
                   Mill, Grade, CalDW, Width, Prod.Type),
            by = c("Fac" = "Mill",
                   "Grade" = "Grade",
                   "CalDW" = "CalDW",
                   "Prod.Width" = "Width"))

opt_sum <- opt_detail %>%
  summarize(nbr.prods = n(),
            nbr.parents = n_distinct(Fac, Grade, CalDW, Parent.Width),
            Tons = sum(Prod.DMD.Tons))
n_parents <- opt_detail %>%
  summarize(nbr.parents = n_distinct(Fac, Grade, CalDW, Parent.Width))

ggplot(mutate(opt_detail, GCDW = paste(Grade, CalDW))
       , aes(x = Prod.Width, y = Prod.DMD.Tons,
             fill = Grade)) +
  geom_bar(stat = "identity", width = .25, position = "stack",
           alpha = 0.7, color = "black") +
  ggtitle(paste(mYear, toString(facility_list), opt_sum$nbr.prods, 
                "SKUs", format(round(opt_sum$Tons, 0), big.mark = ",", 
                               scientific = F), "Tons and ",opt_sum$nbr.prods,
                "Products")) +
  scale_y_continuous(name = "Tons", labels = comma)

# Create a data frame for plotting in long format 
# the Bar.Width variable controlls the width of the bar in the plot
opt_detail_plot <- opt_detail %>%
  gather(Type, Tons, Prod.Gross.Tons, Parent.DMD) %>%
  mutate(Width = ifelse(Type == "Prod.Gross.Tons", Prod.Width, Parent.Width),
         Bar.Width = ifelse(Type == "Prod.Gross.Tons", .1, .2)) %>%
  select(Fac, Grade, CalDW, Type, Width, Bar.Width, Tons) %>%
  distinct()

# Check that this worked
opt_detail_plot %>%
  group_by(Type) %>%
  summarize(Tons = sum(Tons))

# one plot showing all the consolidations
ggplot(opt_detail_plot, 
       aes(x = Width, y = Tons, fill = Type, 
           width = Bar.Width)) +
  geom_bar(stat = "identity", 
           position = "identity",
           alpha = 0.5) +
  scale_fill_manual(values = c("red", "blue")) +
  scale_y_continuous(labels = comma) +
  coord_flip() +
  geom_text(data = filter(opt_detail_plot, Type == "Parent.DMD"),
            aes(label = Width), size = 1.5,
            hjust = "inward") +
  facet_grid(Grade ~ CalDW, scales = "free_x") +
  ggtitle(paste0(mYear, " ", projscenario, "  Consolidations")) +
  theme(axis.text.x = element_text(angle = 45,
                                   hjust = 1.0),
        text = element_text(size = 7),
        legend.key.size = unit(1, "strheight", "1"))

# Create a list of plots split by facility and grade
opt_plot_list <- opt_detail_plot %>%
  unite(Fac.Grade, c(Fac, Grade), remove = FALSE) %>%
  split(.$Fac.Grade) %>%
    map(function(x1)
  ggplot(x1, aes(x = Width, y = Tons, fill = Type, 
                 width = Bar.Width)) +
    geom_bar(stat = "identity", #width = 0.2,
             #position = position_jitter(height = 0),
             position = "identity",
             alpha = 0.5) +
    scale_fill_manual(values = c("red", "blue")) +
    scale_y_continuous(labels = comma) +
    coord_flip() +
    geom_text(data = filter(x1, Type == "Parent.DMD"),
              aes(label = Width), size = 3,
              hjust = "inward") +
    facet_grid( ~ CalDW, scales = "free_x") +
    ggtitle("Grade") +
    theme(axis.text.x = element_text(angle = 45,
                                     hjust = 1.0)))

# Fix the titles
opt_plot_list <- lapply(opt_plot_list, function(x) {
  x$labels$title <- paste0(mYear, ", ", projscenario, ": ", 
                           x$data$Fac, " - ", x$data$Grade);
  return(x)
})

# Save these to a PDF file
parent_plot_file <- file.path(proj_root, "Scenarios", mYear, projscenario, 
                              paste0("parent_plots.pdf"))
if (file.exists(parent_plot_file)) file.remove(parent_plot_file)
pdf(file = parent_plot_file, width = 10, paper = "USr", pointsize = 12,
    onefile = TRUE)
invisible(lapply(opt_plot_list, base::print))
dev.off()

opt_plot_list[5]

# Number of Parents
opt_detail %>%
  group_by(Grade) %>%
  summarize(Nbr.Products = n(),
            Nbr.Parents = n_distinct(Fac, Grade, CalDW, Parent.Width))
opt_detail %>%
  group_by(Fac) %>%
  summarize(Nbr.Products = n(),
            Nbr.Parents = n_distinct(Fac, Grade, CalDW, Parent.Width))

# assign a Parent.Type based on the Prod.Type of the anchor (largest) size
opt_detail <- opt_detail %>%
  arrange(Fac, Grade, CalDW, Parent.Width, desc(Prod.Width)) %>%
  ungroup()

opt_detail <- opt_detail %>%
  group_by(Fac, Grade, CalDW, Parent.Width) %>%
  mutate(Parent.Type = first(Prod.Type)) %>%
  ungroup()

opt_detail <- opt_detail %>%
  mutate(Sim.Parent.Avg.OH.DOH = 365*Sim.Parent.Avg.OH/Parent.DMD) 
  

# compute the statistics for the baseline (all parents case)
sim_parents <- read_results("sim_parents", mYear, projscenario)

# Filter for just the facilities of interest
sim_parents <- sim_parents %>% filter(Fac %in% facility_list)

sim_base <- sim_parents %>%
  filter(nbr_rolls == 1) 

sim_base %>% 
  summarize(Rolls = n(),
            Tons = sum(Sim.Parent.DMD),
            Avg.OH = sum(Sim.Parent.Avg.OH),
            Total.DOH = 365*sum(Sim.Parent.Avg.OH)/sum(Sim.Parent.DMD),
            Trim.Tons = 0,
            Fill.Rate = sum(Sim.Parent.DMD * 
                              Sim.Parent.Qty.Fill.Rate, na.rm = T)/
              sum(Sim.Parent.DMD))

# Summarize the parent rolls
opt_parents <- opt_detail %>% 
  group_by(Fac, Grade, CalDW, Caliper, Dia, Wind, Parent.Width, Parent.DMD,
           Parent.Nbr.Subs, Parent.DMD.Count, 
           Parent.dclass,
           Sim.Parent.Avg.OH, Sim.Parent.Qty.Fill.Rate, 
           Sim.Parent.Avg.OH.DOH,
           Parent.Distrib, Parent.Type, Parent.OTL, Parent.DBR) %>%
  summarize(Parent.Trim.Loss = sum(Prod.Trim.Tons)) %>%
  ungroup()

opt_parents <- opt_parents %>%
  mutate(Parent.dclass.group = 
           ifelse(str_detect(Parent.dclass, "Extreme"), "Slow",
                  ifelse(str_detect(Parent.dclass, "^Intermittent"),
                         "Intermittent", "Non-Intermittent")))

# Any duplicate products by  grade, Caliper & Widths made on multiple
# machines at the same mill
# Do these have to be made on different machines?
copy.table(opt_detail %>%
             ungroup() %>%
             group_by(Grade, Caliper, Prod.Width) %>%
             mutate(Parent.Dups = n()))


# Total Summary
opt_parents %>% 
  ungroup() %>%
  summarize(Nbr.Rolls = n(),
            Annual.Tons = sum(Parent.DMD),
            Avg.OH = sum(Sim.Parent.Avg.OH),
            Total.DOH = 365*sum(Sim.Parent.Avg.OH)/sum(Parent.DMD),
            Trim.Tons = sum(Parent.Trim.Loss),
            Fill.Rate = sum(Parent.DMD * Sim.Parent.Qty.Fill.Rate)/
              sum(Parent.DMD))

# Get the solution costs
sol_sum <- read_results("sol_summary", mYear, projscenario)
# Get the minimum for each Mill/Mach
sol_sum <- sol_sum %>%
  group_by(Fac) %>%
  mutate(Rank = rank(Total.Cost)) %>%
  mutate(RankMinMax = ifelse(Rank == min(Rank), "Min",
                             ifelse(Rank == max(Rank), "Max", NA)))

sol_sum <- sol_sum %>%
  filter(!is.na(RankMinMax))

sol_sum %>% select(Fac, NbrParents, Total.Cost)
copy.table(sol_sum %>% select(Fac, NbrParents, Total.Cost))

# Summarize by parent distribution
opt_parents %>% 
  ungroup() %>%
  group_by(Parent.Distrib) %>%
  summarize(Nbr.Rolls = n(),
            Annual.Tons = sum(Parent.DMD),
            Avg.OH = sum(Sim.Parent.Avg.OH),
            Total.DOH = 365*sum(Sim.Parent.Avg.OH)/sum(Parent.DMD),
            Trim.Tons = sum(Parent.Trim.Loss),
            Fill.Rate = sum(Parent.DMD * Sim.Parent.Qty.Fill.Rate)/
              sum(Parent.DMD))

# Plot Tons by Service Level
ggplot(data = opt_parents, aes(x = Sim.Parent.Qty.Fill.Rate,
                               y = Parent.DMD,
                               color = Parent.dclass)) +
  geom_bar(stat = "identity", width = .01, position = "identity",
           fill = "grey") +
  scale_x_reverse(labels = percent)

# Which parents are in or out (MTS vs. MTO)

# Create the MTO vs MTS Inventory Policy
# Identify Slow Demand ( < 5)
opt_parents <- opt_parents %>%
  mutate(Parent.DMD.Freq = ifelse(Parent.DMD.Count <= 5, "Slow", 
                                  ifelse(Parent.DMD.Count <= 30, 
                                         "Medium", "Fast")))

opt_parents %>% 
  group_by(Parent.DMD.Freq) %>%
  summarize(Sub.Rolls = sum(Parent.Nbr.Subs),
            Nbr.Rolls = n(),
            Annual.Tons = sum(Parent.DMD),
            Avg.OH = sum(Sim.Parent.Avg.OH),
            Total.DOH = 365*sum(Sim.Parent.Avg.OH)/sum(Parent.DMD),
            Trim.Tons = sum(Parent.Trim.Loss),
            Fill.Rate = sum(Parent.DMD * Sim.Parent.Qty.Fill.Rate)/
              sum(Parent.DMD),
            Dmd.Count = sum(Parent.DMD.Count))

# Create the Parent.Policy field
opt_parents <- opt_parents %>%
  mutate(Parent.Policy = 
           ifelse((Parent.Type %in% c("Bev.VMI", "CPD.VMI",  
                    "CPD.Web") & Sim.Parent.Avg.OH.DOH < 45) | 
                    Sim.Parent.Avg.OH.DOH < 30, 
                  "MTS", "MTO"))


if (projscenario == "Europe") {
  opt_parents <- opt_parents %>%
    mutate(Parent.Policy = 
             ifelse(Sim.Parent.Avg.OH.DOH < 60 |                              
                      str_sub(Parent.dclass, 1, 5) == "Non-I", 
                    "MTS", "MTO"))
}

# Summarize by Policy
opt_parents %>% 
  group_by(Parent.Policy) %>%
  summarize(Sub.Rolls = sum(Parent.Nbr.Subs),
            Nbr.Rolls = n(),
            Annual.Tons = sum(Parent.DMD),
            Avg.OH = sum(Sim.Parent.Avg.OH),
            Total.DOH = 365*sum(Sim.Parent.Avg.OH)/sum(Parent.DMD),
            Trim.Tons = sum(Parent.Trim.Loss),
            Fill.Rate = sum(Parent.DMD * Sim.Parent.Qty.Fill.Rate)/
              sum(Parent.DMD))

# Histogram of the number of shipments for each parent roll
ggplot(data = opt_parents, 
       aes(x = Parent.DMD.Count)) +
  geom_histogram(binwidth = 1, aes(fill = Parent.Policy), color = "black") +
  ggtitle("Parent Roll Shipment Counts") +
  scale_y_continuous(name = "Number of Parent Rolls") +
  scale_x_continuous(name = "Number of Shipments") +
  theme(legend.key.size = unit(12, "points"))

# Add the Parent.Policy back to the opt_detail so we can tally the
# the volumes by Prod.Type
opt_detail <- opt_detail %>%
  select(-starts_with("Parent.Policy"))
opt_detail <- left_join(opt_detail, 
                        select(opt_parents, Fac, Grade, CalDW, 
                               Parent.Width, Parent.Policy),
                        by = c("Fac" = "Fac",
                               "Grade" = "Grade",
                               "CalDW" = "CalDW",
                               "Parent.Width" = "Parent.Width"))

# Summarize tons by Prod.Type
# Careful here with the OH and DOH since we are filtering
# in the sum function
opt_detail %>% group_by(Parent.Policy, Prod.Type) %>%
  summarize(Tons = round(sum(Prod.Gross.Tons), 0), 
            SKUs = n(),
            Parents = n_distinct(Fac, Grade, CalDW, Parent.Width),
            Avg.OH = round(sum(ifelse(Parent.Prod.IDX == 1, 
                                      Sim.Parent.Avg.OH, 0)), 1),
            Avg.DOH = round(365*sum(ifelse(Parent.Prod.IDX == 1, 
                                           Sim.Parent.Avg.OH, 0))/
                              sum(ifelse(Parent.Prod.IDX == 1, 
                                         Parent.DMD, 0)), 1),
            Trim = round(sum(Prod.Trim.Tons), 0)) %>%
  arrange(desc(Parent.Policy), desc(Tons))


copy.table(opt_detail %>% group_by(Parent.Policy, Prod.Type) %>%
             summarize(Tons = round(sum(Prod.Gross.Tons), 0), 
                       SKUs = n(),
                       Parents = n_distinct(Fac, Grade, CalDW, Parent.Width),
                       Avg.OH = round(sum(ifelse(Parent.Prod.IDX == 1, 
                                                 Sim.Parent.Avg.OH, 0)), 1),
                       Avg.DOH = round(365*sum(ifelse(Parent.Prod.IDX == 1, 
                                                      Sim.Parent.Avg.OH, 0))/
                                         sum(ifelse(Parent.Prod.IDX == 1, 
                                                    Parent.DMD, 0)), 1),
                       Trim = round(sum(Prod.Trim.Tons), 0)) %>%
             arrange(desc(Parent.Policy), desc(Tons))
)


copy.table(opt_detail %>% group_by(Parent.Policy, Prod.Type) %>%
             summarize(Tons = sum(Prod.Gross.Tons), 
                       SKUs = n()))

# Parent Demand Count vs. DOH
ggplot(data = opt_parents, aes(x = Parent.DMD.Count,
                               y = Sim.Parent.Avg.OH.DOH,
                               color = Sim.Parent.Qty.Fill.Rate)) +
  geom_point(aes(shape = Parent.Policy), size = 3, alpha = .5) +
  scale_shape_manual(values = c(4, 5, 19)) +
  scale_colour_gradient(low="red", high = "blue")  +
  scale_x_log10(labels = comma) +
  ggtitle("Average Inventory Levels vs. Number of Orders") +
  xlab("Number of Orders") +
  ylab("Average Days On Hand (Simulation Generated)") +
  labs(color = "Fill Rate")


# Tons vs. DOH
ggplot(data = opt_parents,
       aes(x = Sim.Parent.Avg.OH.DOH,
           y = Parent.DMD,
           color = Sim.Parent.Qty.Fill.Rate)) +
  geom_point(size = 3, alpha = .3) +
  scale_shape_manual(values = c(4, 5, 19)) +
  scale_colour_gradient(name = "Fill Rate\n", low="red", high = "blue",
                        labels = percent) +
  scale_y_log10(labels = comma) +
  scale_x_continuous(limits = c(0,200)) +
  xlab("Average Days on Hand") +
  ylab("Parent Tons (Log Scale)") +
  ggtitle("Parent Roll Annual Volume vs. DOH") +
  labs(color = "Fill Rate") + 
  geom_rect(xmin = 30, xmax = Inf, ymin = -Inf, ymax = Inf,
            fill = "lightgreen", alpha = .006,
            inherit.aes = FALSE) +
  geom_text(aes(x = 150, y = 10000), label = "DOH > 30")


# Quadrant Criteria
cr_FR <- .95
cr_days <- 30

opt_detail <- opt_detail %>%
  mutate(FR = ifelse(round(Sim.Parent.Qty.Fill.Rate, digits = 2) >= cr_FR, 
                     paste0("FR >= ", 100*cr_FR,"%"), 
                     paste0("FR < ", 100*cr_FR,"%")),
         Days = ifelse(Sim.Parent.Avg.OH.DOH <= cr_days, 
                       paste0("Days <= ", cr_days),
                       paste0("Days > ",cr_days)))

rm(cr_days, cr_FR)

opt_parents %>%
  filter(Sim.Parent.Avg.OH.DOH <= 30) %>%
  summarize(Tons = sum(Parent.DMD),
            Parents = n(),
            Avg.OH = sum(Sim.Parent.Avg.OH))

save(opt_detail, file = file.path(proj_root, "Scenarios",
                                  mYear, projscenario, "opt_detail.RData"))

