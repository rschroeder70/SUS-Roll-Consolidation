# -----------Get the info for the selected pilot rolls

# copy the Caliper, Grade, Mill, Dia, Wind & Roll.Width from
# the ParentSummary-pilot-rolls.xlsx spreadseet
pilot_rolls <- paste.table()

# Convert the factors to characters
pilot_rolls <- pilot_rolls %>%
  map_if(is.factor, as.character) %>%
  as_data_frame

pilot_rolls <- pilot_rolls %>%
  mutate(Caliper = as.character(Caliper),
         Dia = as.character(Dia))

# Join this with the opt_detail_plants data frame to get the plant
# info.  This comes from 7.2
if (!exists("opt_detail_plants")) source("R/7.2 Save Excel Validation.R")

pilot_rolls_plants <- 
  inner_join(pilot_rolls, opt_detail_plants,
            by = c("Mill", "Grade", "Caliper",
                   "Dia", "Wind", "Roll.Width" = "Prod.Width")) %>%
  rename(Plant.Tons = Tons)
pilot_rolls_plants <- 
  left_join(pilot_rolls_plants, facility_names,
            by = c("Plant" = "Facility.ID"))

pilot_rolls_plants_sum <- pilot_rolls_plants %>%
  group_by(Fac.Abrev) %>%
  summarize(Rolls = n(), 
            Tons = comma(sum(Plant.Tons)),
            Plant.Trim.Tons = 
              comma(sum(Plant.Tons * Prod.Trim.Fact - Plant.Tons))) %>%
  filter(!(Fac.Abrev %in% c("MEN", "HAM", "ORO", "STAU", "VF", "WAYNE")))

copy.table(pilot_rolls_plants_sum)

# Join to the dies_plants data frame to get the die info for each size roll
#   this comes from the Read Comsumption by Die.R script
pilot_rolls_by_die <- pilot_rolls_plants %>%
  left_join(mutate(filter(web, Year == "2016"), Die.Tons = round(Tons, 1)),
            by = c("Plant", "Grade", "Caliper" = "Board.Caliper", 
                   "Roll.Width" = "Board.Width", "Dia", "Wind")) %>%
  select(Caliper, Grade, Dia, Wind, Roll.Width, Plant, Fac.Abrev, Plant.Tons, 
         Parent.Width, Prod.Trim.Width, Die, Die.Description, Die.Tons) %>%
  filter(!(Fac.Abrev %in% c("MEN", "HAM", "ORO", "STAU", "VF", "WAYNE")))

copy.table(pilot_rolls_by_die)


# How much inventory
# First get the parent width
pilot_parent_rolls <-
  inner_join(pilot_rolls, opt_detail,
             by = c("Mill" = "Fac", "Grade", "Caliper", "Dia", "Wind",
                    "Roll.Width" = "Prod.Width"))

pilot_parent_rolls_inv <- pilot_parent_rolls %>%
  distinct(Caliper, Grade, Dia, Wind, Mill, Parent.Width, Sim.Parent.Avg.OH, 
           Parent.DMD, Parent.Nbr.Subs)

copy.table(pilot_parent_rolls_inv)

# What is the avage QV inventory from 2016
sr_inv_mean <- ungroup(sr_inv_mean)
pilot_qv_inv <- 
  left_join(pilot_rolls, sr_inv_mean,
            by = c("Caliper", "Grade", "Wind", "Roll.Width" = "Width")) %>%
  select(-Prod.Policy, -Prod.Type)

copy.table(pilot_qv_inv)

#------------ Check vs. the validation files
pilot_rolls_by_die$Dia <- as.numeric(pilot_rolls_by_die$Dia)

pilot_skus_to_fix <- skus_to_fix %>%
  select(Plant, Grade, Caliper, Actual.Width, Dia, Wind, Parent.Width, 
         Valid) %>% 
  mutate(Parent.Width = as.numeric(Parent.Width)) %>%
  distinct()

pilot_skus_check <- left_join(pilot_rolls_by_die, pilot_skus_to_fix,
          by = c("Caliper", "Grade", "Dia", "Wind", 
                 "Roll.Width" = "Actual.Width", "Plant", 
                 "Parent.Width" = "Parent.Width")) %>%
  select(Plant, Fac.Abrev, Caliper, Grade, Dia, Wind, Roll.Width, 
         Parent.Width = Parent.Width, Prod.Trim.Width, 
         Die, Die.Description, Die.Tons, Valid)

copy.table(pilot_skus_check)

