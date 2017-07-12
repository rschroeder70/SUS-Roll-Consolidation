#------------------------------------------------------------------------------
#
# Compute estimated trimm loss by die at each plant
#
# The file comes from QlikView Fiber Analysis Consumption by Die Report, 
# Bookmark = SUS Web Consumed with the Date set to Daily.
#
#
library(tidyverse)
library(stringr)
library(openxlsx)
library(scales)
library(lubridate)

# for openxlsx writing
Sys.setenv("R_ZIPCMD" = "C:/RBuildtools/3.4/bin/zip.exe") 

proj_root <- file.path("C:", "Users", "SchroedR", "Documents",
                       "Projects", "SUS Roll Consolidation")

cons <- read.xlsx(file.path(proj_root, "Data", 
                           "QV SUS Roll Consumption 2015-2017.xlsx"),
                 detectDates = T)

cons <- cons %>%
  mutate(Plant = paste0("PLT0", Plant),
         Year = as.character(year(MonthYear)),
         Board.Width = as.numeric(Board.Width)) %>%
  filter(Board.Width > 0)

cons <- cons %>%
  mutate(M.I = str_sub(Material.Description, 1, 1))

# Grade
cons <- cons %>%
  mutate(Grade = ifelse(
        M.I == "I",
        str_sub(Material.Description, 5, 8), # Imperial
        str_sub(Material.Description, 6, 9)  # Metric
        ) 
      )

cons <- cons %>%
  filter(Grade %in% c("AKOS", "OMXX", "PKXX", "FC02", "AKPG", "PKHS", "PKSR"))

cons <- cons %>%
  mutate(Wind = str_sub(Material.Description, 
                        str_length(Material.Description), -1))
cons <- cons %>%
  mutate(Dia = ifelse(M.I == "I",
                      str_sub(Material.Description, 19, 20),
                      str_sub(Material.Description, 16, 19)))

# Aggregate over the dates
cons <- cons %>%
  group_by(Year, Plant, PlantName, 
           MRPC2.Text, Die, Die.Description,
           Grade, Board.Caliper, Dia, Wind, Board.Width, M.I) %>%
  summarize(Tons = sum(TON)) %>%
  ungroup()

cons <- cons %>%
  filter(Year == "2016")

cons %>% group_by(PlantName) %>%
  summarize(Tons = sum(Tons))

# Calculate the percentage of tons by board width for the dies with 
#   multiple widths.  Ignore Grade & Caliper
cons_mw <- cons %>%
  group_by(Year, Plant, PlantName, Die, Die.Description) %>%
  mutate(Nbr.Widths = n_distinct(Board.Width),
         Die.Tons = sum(Tons)) %>%
  mutate(Die.Pct = round(100*Tons/Die.Tons, 2)) %>%
  #filter(Nbr.Widths > 1) %>%
  ungroup() %>%
  arrange(Year, Plant, Die, desc(Board.Width)) 

# If we have less than x%, assume the preference is the next smaller roll
# with more than 10% of the volume
parent_eval <- function(x) {
  x$Trim.Width <- 0
  for (r in 1:nrow(x)) {
    if (x$Die.Pct[r] < 15) {
      # Find the next row where the Pct > 10
      mrow <- which(x$Die.Pct[r:nrow(x)] > 20)[1]
      if (!is.na(mrow) & mrow > 1) {
        mtrim <- x$Board.Width[r] - x$Board.Width[r - 1 + mrow]
        x$Trim.Width[r] <- mtrim
      } 
    }
  }
  return(x)
}

trim_eval <- cons_mw %>%
  unite(plant.die, c(Plant, Die), sep = ".") %>%
  split(.$plant.die) %>%
  map_df(parent_eval)

trim_eval <- trim_eval %>%
  mutate(Trim.Factor = (Board.Width %% (Board.Width - Trim.Width)/
                          (Board.Width - Trim.Width)))

trim_eval$Trim.Tons <- round(trim_eval$Tons * trim_eval$Trim.Factor, 2)

trim_plant <- trim_eval %>% filter(Trim.Width < 3) %>%
  group_by(PlantName) %>%
  summarize(Total.Tons = sum(Tons),
                        Trim.Tons = round(sum(Trim.Tons), 1))

trim_plant %>%
  summarize(Total.Tons = sum(Total.Tons),
            Trim.Tons = sum(Trim.Tons))

#rm()

