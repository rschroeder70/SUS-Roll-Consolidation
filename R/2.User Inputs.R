#------------------------------------------------------------------------------
#
# User inputs for the model run.  Called from 1.Main.R
#

print("Reading model parameters")

# Two type of projects: Mills or Plants
projtype = "Mill" # "Mill" or "Plant"
projname = "SUS"

#projscenario <- "BevOnly"
#projscenario <- "Base" # Uses Gamma distribution
#projscenario <- "Base96"  # 96% service target.  Change in User Inputs.
#projscenario <- "CPDOnly-95-8-150-500" # Used for Pilot Roll Selection
#projscenario <- "VMI-98-8-150-725" # Service, ic, opty, prod cust
#projscenario <- "VMI-98-8-150-500v2"  # Fix the 24pt. 42.875
projscenario <- "VMI-98-8-150-500v3"  # Fix the 21pt. 37
#projscenario <- "VMI-98-3-380-500"
#projscenario <- "VMI-98-8-380-500"
#projscenario <- "Europe-95-8-140-415"
#projscenario <- "NoEurope95"
#mYear <- "2015"
mYear <- "2016"
start_date <- as.Date(paste(mYear, "01", "01", sep = "-"))
end_date <- as.Date(paste(mYear, "12", "31", sep = "-"))

# Safety Factor (95% service level = 1.65)
if (str_sub(projscenario, 1, 3) %in% c("Bev", "VMI")) {
  service_level <- 0.98
} else {
  service_level <- 0.95
}
sfty_factor <- qnorm(service_level,0,1)

# Inventory Carrying cost of funcs
cost_of_funds <- .08
#cost_of_funds <- .03

# Paper std cost per ton, excluding inbound freight because we are assuming
# we would stock at the mill and then ship to the plants
#prod_cost <- 725
prod_cost <- 500

# The opportunity cost of a ton not selling what is now going to trim
trim_opty_cost <- 150
#trim_opty_cost <- 140

#  Set the aggregation period for the demand analysis [daily, weekly, monthly]
#  The demand classification routine will calculate the statistics over this.
dmd_aggreg_period <- "weekly"

# The lead times will be in days, so we need this to convert the stastics
# to the lead time duration
if (dmd_aggreg_period == "weekly") {
  dmd_lt_conv <- 7
} else if (dmd_aggreg_period == "daily") {
  dmd_lt_conv <- 1
}

# Use a gamma distribution if the coeffic of var > 0.5
use_gamma <- TRUE
if (use_gamma) library(EnvStats)

# Use the cutter limit logic to limit roll size options
use_cutter_limits = TRUE

# if we are analyzing mill stocking, input the mill names that are 
# in the data file.  Otherwise the name of the plant(s) were we will stock
# rolls
facility_list <- c("M1", "M2", "W6", "W7")
#facility_list <- c("W6")


# Input and Output Data files are in the Data/Data+projname folder
# Generic data files are in the Data directory
# variables starting with "fp" are file paths

fpinput <- file.path(proj_root, "Data")

fpoutput <- file.path(proj_root, "Scenarios", mYear, projscenario)
if (!file.exists(fpoutput)){
  dir.create(fpoutput)
}

# name of the file containing the consumption data
consdatafile <- paste(mYear, "Shipments for Dylan.xlsx", sep = " ")
#consdatafilecsv <- "2016 Shipments for Dylan.csv"
consdatafilecsv <- "All Shipments Fixed.csv"
consdatasheet <- "Shipment data"

# input location for general data used in all model runs, 
fpdata <- file.path(basepath, "Data")

# Shipping data - Freight Rates between plants and mills
shippingData <- file.path(basepath, "Data", "ShipmentData.xlsx")

cycleLength <- file.path(basepath, "Data", "cycle_lengths.xlsx")

#name of file to save roll selection information
rollSelectionOutput <- file.path(fpoutput, paste(projname, projscenario,
                                                 "RollSelection.csv", sep = "_"))

finalresults <- file.path(fpoutput,paste(projname, projscenario, 
                                         "FinalResults.xlsx", "_"))

#list of locations with slitter limitations (can't double cut)
slitterLimitPlants <- c()
slitterLimitPlants <- c('GPI-STONE MOUNTAIN,GA', 'GPI-CAROL STREAM,IL', 'GPI-MARION,OH', 
                        'GPI-OAKS,PA', 'GPI-LUMBERTON,NC', 'GPI-FORT SMITH,AR', 
                        'GPI-CENTRALIA,IL', 'GPI-MITCHELL,SD', 'GPI-KENDALLVILLE,IN', 
                        'GPI-GORDONSVILLE,TN', 'GPI-KALAMAZOO,MI', 'GPI-LAWRENCEBURG,TN', 
                        'GPI-WAUSAU,WI', 'GPI-COBOURG', 'GPI-PORTLAND,OR', 'GPI-TUSCALOOSA,AL', 
                        'PLT00006', 'PLT00008', 'PLT00015', 'PLT00020', 'PLT00023', 'PLT00041', 
                        'PLT00042', 'PLT00043', 'PLT00044', 'PLT00047', 'PLT00048', 'PLT00053', 
                        'PLT00054', 'PLT00070', 'PLT00068')

# Big M used in IP to select parents and assign products.
# Must be larger than biggest possible number of child rolls
bigM <- 300 

#trim limit at converting plant in inches (2 --> 1" from each side)
trimLimit <- 3

# List of cutter sizes in converting plants in inches
#cutters <- c(32, 33.5, 40, 42.375, 55, 56, 64)
# This list includes the press widths at the Perry & WM
# (WM = 36, 48, 55, 67; P = 44, 50, 67)
# Beverage Only cutter limits
if (str_sub(projscenario, 1, 7) %in% c("BevOnly", "VMI98")) {
  cutters <- c(36, 44, 48, 50, 55, 56, 64, 67)
} else {
  cutters <- c(32, 33.5, 36, 40, 42.375, 44, 48, 50, 55, 56, 64, 67)
}


#days allowed for simulation to stabilize before metrics are taken 
stabilizationPeriod <- 60 

#cycle length
cycle_length_list <- read.xlsx(cycleLength, sheet = projname)
cycle_length_list <- cycle_length_list %>%
  filter(!is.na(Mean.DBR))
cycle_length_list$Mean.DBR <- as.numeric(cycle_length_list$Mean.DBR)

# These are missing - add manually
temp <- data_frame(Plant = "WM", Machine = "07", Grade = "AKPG",
           Caliper = "24", Mean.DBR = 8, Mill.Machine = "W7")
cycle_length_list <- bind_rows(cycle_length_list, temp)

temp <- data_frame(Plant = "WM", Machine = "06", Grade = "FC02",
                   Caliper = "31", Mean.DBR = 16, Mill.Machine = "W6")
cycle_length_list <- bind_rows(cycle_length_list, temp)
temp <- data_frame(Plant = "Macon", Machine = "02", Grade = "AKPG",
                   Caliper = "24", Mean.DBR = 9, Mill.Machine = "M2")
cycle_length_list <- bind_rows(cycle_length_list, temp)

rm(temp)

# transportation_length (mill to plant)

if (projtype == "Mill") {
  transportation_length <- 2
} else {
  # Need to fix this to read from a O-D file
  transportation_length <- 12
}

# Order Processing Time (Days prior to paper machine cycle)
order_proc_length <- 2
#lead_time <- cycle_length_list$cycle_length + transportation_length

#Paper Machine Trim.R 

#average tons/roll
#tonsPerRoll <- 1.92

#path to export data to use as input for python model,line 39 
#pythonInputPath <- paste(mill,"_output_for_trim.csv", sep="")

#file name of python model to run from, line 42
#pythonModel <- 'cutstock_looper.py'
#file name of results of python model, line 45

#pyTrimResults <- paste(mill,"_MachineTrimResults.csv", sep="")
#incher per ton, lines 48, 51 

#inchesPerTon <- 17.5

# Read the freight cost table
trim_freight <- read.xlsx(shippingData, sheet = "Sheet1")
trim_freight <- trim_freight %>%
  rename(Cost.Ton = `Cost/Ton`)

# Warehouse Shuttle and handling
# WM: $4.72/ton Handling + $5/ton Shuttle
# M: $4.52/ton Handling + $5/ton Shuttle
handling <- c(M = 4.52 + 5.00,
              W = 4.72 + 5.00)

# Warehouse storage - use 10 sq ft per ton
# WM: $3.00/sq ft. per year
# M: $2.16/sq ft per year
storage <- c(W = 3.00 * 10,
             M = 2.16 * 10)

# inventory holding cost per ton per day.  Adjust this by
# multiplying by the lead time days for comparison with trim costs
ihc_per_year <- cost_of_funds * prod_cost
ihc_per_day <- cost_of_funds * (1 / 365) * prod_cost

# Parameters for the demand classification
dmd_parameters <- c(nbr_dmd = 5, # Extremely Slow nbr demands
                    dmd_mean = 5, # Extremely Small demand mean
                    outlier = 10, 
                    intermit = 1.32, # Intermittency demand interval
                    variability = 4, # non-zero dmd std dev
                    dispersion = 0.49) # CV^2, non inter:smooth, 
# inter: eratic - slow, lumpy


print("Finished loading libraries and reading model parameters")
