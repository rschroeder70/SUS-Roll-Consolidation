#------------------------------------------------------------------------------
#
# Main module for the SUS Board Roll Consolidation routines
# 
# --------------------------------
# Load libraries

# Base Roll Consolidation R script file directory for files common
# to all projects
basepath <- file.path("C:","Users", "SchroedR", "Documents",
                      "Projects", "RollAssortmentOptimization")

# This projects root directory
proj_root <- file.path("C:", "Users", "SchroedR", "Documents",
                       "Projects", "SUS Roll Consolidation") 

# Load Packages & Functions
source(file.path(basepath, "R", "0.LoadPkgsAndFunctions.R"))

source("R/2.User Inputs.R")
print(projscenario)
print(service_level)
print(cost_of_funds)
print(trim_opty_cost)
print(prod_cost)
source("R/3.Read Data.R")
#ifac <- "W6"  # for testing
for (ifac in facility_list) {
  # Just get this mill/plant for the current iteration.  
  if (projtype == "Mill") {
    rawData <- rawDataAll %>%
      filter(Mill == ifac)
  } else {
    rawData <- rawDataAll %>%
      filter(Plant == ifac)
  }
  #rawData <- filter(rawData, Grade == "AKPG" & CalDW == "21-84-W") 
  #rawData <- filter(rawData, Width == 37)
  source(file.path(basepath, "R", "3.5.Demand Profiler.R"))  
  source(file.path(basepath, "R", "4.Parent Roll Selection.R"))
  source(file.path(basepath, "R", "5.Simulate Parents.R"))
  source(file.path(basepath, "R", "6.Save Results.R"))
}

source("R/7.1 Opt Solution Post Process.R")
source("R/7.2 Save Excel Validation.R")

# The full output table is in solutionRollDetailsSim
# Filter this to just the optimal solution
solutionRollParentsOpt <- solutionRollDetailsSim %>%
  ungroup() %>%
  filter(SolParentCount == min_sol$NbrParents) %>%
  select(-starts_with("Sol.")) %>%
  select(-starts_with("Prod.")) %>%
  distinct()

solutionRollParentsOpt %>%
  group_by(solutionNbr, SolParentCount) %>%
  summarize(Rolls = n(),
    AvgOnHand = round(sum(Sim.Parent.Avg.OH), 1),
            OTL = round(sum(Sim.Parent.OTL), 1),
            AvgDOH = round(300 * sum(Sim.Parent.Avg.OH)/
              sum(Sim.Parent.DMD), 1),
            BODays = sum(Sim.Parent.BO.Days), 
            TotDemand = round(sum(Sim.Parent.DMD), 1),
            CumDemand = sum(Sim.Parent.Cum.DMD),
            NbrOrders = sum(Sim.Parent.DMD.Count),
            NbrBO = sum(Sim.Parent.Nbr.BO),
            CumBO = sum(Sim.Parent.Cum.BO),
            OrdFillRate = round(1 - sum(Sim.Parent.Nbr.BO)/
              sum(Sim.Parent.DMD.Count), 2),
            QtyFillRate = round(1 + sum(Sim.Parent.Cum.BO)/
              sum(Sim.Parent.Cum.DMD), 2)) %>%
  ungroup()

solutionRollParentsOpt %>%
  filter(Parent.DMD.Count > 5) %>%
  group_by(solutionNbr, SolParentCount) %>%
  summarize(rolls = n(),
    AvgOnHand = round(sum(Sim.Parent.Avg.OH), 1),
            OTL = round(sum(Sim.Parent.OTL), 1),
            AvgDOH = round(300 * sum(Sim.Parent.Avg.OH)/
                             sum(Sim.Parent.DMD), 1),
            BODays = sum(Sim.Parent.BO.Days), 
            TotDemand = round(sum(Sim.Parent.DMD), 1),
            CumDemand = sum(Sim.Parent.Cum.DMD),
            NbrOrders = sum(Sim.Parent.DMD.Count),
            NbrBO = sum(Sim.Parent.Nbr.BO),
            CumBO = sum(Sim.Parent.Cum.BO),
            OrdFillRate = round(1 - sum(Sim.Parent.Nbr.BO)/
                                  sum(Sim.Parent.DMD.Count), 2),
            QtyFillRate = round(1 + sum(Sim.Parent.Cum.BO)/
                                  sum(Sim.Parent.Cum.DMD), 2)) %>%
  ungroup()

