#------------------------------------------------------------------------------
#
# Combine the simulation runs to an overall agregate
#
#library(plyr)
library(tidyverse)
library(scales)
library(stringr)
library(openxlsx)


# Combine the files with details from the optimal run for each facility
# Use the data in the sim plot files

#mYear <- "2016"
#projscenario <- "BevOnly"

# First run the 7.1 Opt Solution Post Process.R script
if (!exists("opt_detail")) source(file.path(proj_root, "R", 
                                            "7.1 Opt Solution Post Process.R"))

# read the lists and combine into a single list
for (ifac in facility_list) {
  load(file.path(proj_root, "Scenarios", mYear, projscenario, 
                 paste0(ifac, " Sim Plots.RData")), verbose = TRUE)
}

# should be a way to do this in a loop or lapply
sim_plotlist <- c(sim_plotlist_M1, sim_plotlist_M2,
                  sim_plotlist_W6, sim_plotlist_W7)
rm(sim_plotlist_M1, sim_plotlist_M2,
   sim_plotlist_W6, sim_plotlist_W7)


#sim_plotlist[[25]]$labels$title
sim_plotlist[[32]]

# This will give a list of data frames the data points
t_data <- lapply(sim_plotlist, '[[', 'data')
# Get rid of backorders and replen orders
t_data <- lapply(t_data, function(x) select(x, -(Back.Order:Replen.Order)) )


# This will get the list of plot labels
t_labels <- lapply(sim_plotlist, '[[', 'labels')

# Extract just the title.  Sapply puts them in a vector
t_titles <- sapply(t_labels, '[[', 'title')

# Create a logical vector for the base case (all products).
# This is true if the title contains "Sub = 1"
t_base_keep <- str_detect(t_titles, "Subs = 1")
t_data_base <- t_data[t_base_keep]

# Get the Fac, Grade, CalDW and Width from all the titles
# We will use this to identify the parents in the optimal solution
t_titles_parts <- str_split(t_titles, " ")
t_fac <- sapply(t_titles_parts, "[[", 3)
t_grade <- sapply(t_titles_parts, "[[", 4)
t_calDW <- sapply(t_titles_parts, "[[", 5)
t_width <- sapply(t_titles_parts, "[[", 6)
t_nbrsubs <- sapply(t_titles_parts, "[[", 9)

rm(t_titles_parts)

t_opt <- data_frame(t_fac, t_grade, t_calDW, t_width, t_nbrsubs)
t_opt <- t_opt %>%
  mutate(t_width = as.numeric(t_width),
         t_nbrsubs = as.integer(t_nbrsubs))
t_opt <- left_join(t_opt, distinct(select(opt_detail,
                                          Fac, Grade, CalDW, Parent.Width,
                                          Parent.Nbr.Subs, Parent.DMD,
                                          Sim.Parent.Avg.OH, Parent.Policy)),
                   by = c("t_fac" = "Fac",
                          "t_grade" = "Grade",
                          "t_calDW" = "CalDW",
                          "t_width" = "Parent.Width",
                          "t_nbrsubs" = "Parent.Nbr.Subs"))

rm(t_fac, t_grade, t_calDW, t_width, t_nbrsubs)

# Include the parent policy in the title
t_titles <- paste(t_titles, t_opt$Parent.Policy, sep = " - ")

# No match to a parent means the roll is not part of the optimal solution
t_opt_keep <- !is.na(t_opt$Sim.Parent.Avg.OH)

#Check the tons and the count of parent rolls
t_opt %>% filter(!is.na(Parent.DMD)) %>% 
  summarize(Tons = sum(Parent.DMD),
            Count = n())

# Modify the list of titles for the opt plots
t_titles_opt <- t_titles[t_opt_keep]
t_titles_opt <- as.list(t_titles_opt)

t_data_opt <- t_data[t_opt_keep]

# Create a new list of the plots for the parents in the
# optimal solution
sim_plot_opt <- sim_plotlist[t_opt_keep]

# Modify the plot titles to indicate Parent MTS or MTO
#sim_plot_opt[[1]]$labels$title <- t_titles_opt[[1]]
# Original version
#sim_plot_opt <- mapply(function(x, y) {x$labels$title = y; return(x)},
#            sim_plot_opt, t_titles_opt, SIMPLIFY = FALSE)

# purrr version using map2
sim_plot_opt <- map2(sim_plot_opt, t_titles_opt, 
                     function(x, y) {x$labels$title = y; return(x)})

# write this to a pdf file
pdf(file = file.path(proj_root, "Scenarios", mYear, projscenario, 
                     "W6 Parent Sim Plots.pdf"),
    width = 10, paper = "USr", pointsize = 12,
    onefile = TRUE)
invisible(lapply(sim_plot_opt, base::print))
dev.off()

rm(t_titles, t_titles_opt, t_opt, t_labels)

# Fill in all the missing days so we can combine the data
tDays <- data.frame(seq.Date(as.Date("2016-01-01"), 
                             as.Date("2016-12-31"), by = "day" ))
colnames(tDays)[1] <- "Date"

t_data_base <- lapply(t_data_base, function(x) 
  left_join(tDays, x, by = c("Date" = "Date")))
t_data_base <- lapply(t_data_base, function(x) 
  fill(x, On.Hand, .direction = "down"))
t_data_base <- lapply(t_data_base, function(x) 
  fill(x, On.Hand, .direction = "up"))

# consolidate the list of data lists into 1 long list
# Map applies a function to the elements of given vectors and returns a 
# data frame.  It's a wrapper for mapply but does not attempt to simplify
#t_data_base <- do.call(Map, c(base::c, t_data_base))
t_data_base <- do.call(rbind, t_data_base)
#t_data_base <- as_data_frame(t_data_base)
t_data_base <- t_data_base %>%
  group_by(Date) %>%
  summarize(On.Hand = sum(On.Hand, na.rm = TRUE))
t_data_base$Scenario <- paste("Base:", sum(t_base_keep), "Rolls")

# Do the same for the opt data set
t_data_opt <- lapply(t_data_opt, function(x)
                     left_join(tDays, x, by = c("Date" = "Date")))
t_data_opt <- lapply(t_data_opt, function(x) 
  fill(x, On.Hand, .direction = "down"))
t_data_opt <- lapply(t_data_opt, function(x) 
  fill(x, On.Hand, .direction = "up"))

# consolidate the list of data lists
#t_data_opt <- do.call(Map, c(base::c, t_data_opt))
t_data_opt <- do.call(rbind, t_data_opt)
#t_data_opt <- as_data_frame(t_data_opt)
t_data_opt <- t_data_opt %>%
  group_by(Date) %>%
  summarize(On.Hand = sum(On.Hand, na.rm = TRUE))
t_data_opt$Scenario <- paste("Opt:", sum(t_opt_keep), "Rolls")

# Combine into a single long dataframe
t_data_res <- bind_rows(t_data_base, t_data_opt)

rm(tDays, t_data_base, t_data_opt, t_base_keep, t_opt_keep)

# Add the sr_inv_me data frame that we create in QV Roll Inv.R
if (!exists("sr_inv_me")) {
  source(file.path(proj_root,"R","QV Roll Inv.R"))
}

# Aggregate for all Prod.Types & Policies
sr_inv_me <- sr_inv_me %>%
  group_by(Date, Scenario) %>%
  summarize(On.Hand = sum(On.Hand),
            Annual.DMD = sum(Annual.DMD))

t_data_res <- bind_rows(t_data_res, 
                        select(sr_inv_me, Date, On.Hand, Scenario))

# Compute the mean for each scenario starting Feb 1
t <- t_data_res %>% 
  group_by(Scenario) %>%
  filter(Date >= as.Date("2016-02-01")) %>%
  mutate(Avg.OH = as.integer(mean(On.Hand))) %>%
  select(-On.Hand)
  
t_data_res <- left_join(t_data_res, t,
                        by = c("Date" = "Date",
                               "Scenario" = "Scenario"))
rm(t)

ggplot(filter(t_data_res, Scenario %in% c("Base: 109 Rolls",
                                          "Opt: 84 Rolls")),
       aes(x = Date, group = Scenario)) +
  geom_step(aes(color = Scenario, y = On.Hand)) +
  scale_color_brewer(palette = "Set1") +
  geom_line(aes(color = Scenario, y = Avg.OH), size = .1) +
  scale_y_continuous(name = "On Hand Tons", labels = comma) +
  geom_label(aes(x = max(Date), y = Avg.OH, label = 
                   format(Avg.OH, big.mark = ",")), hjust = "inward") +
  ggtitle("W6 Simulation Results")
  
ggplot(t_data_res,  
       aes(x = Date)) +
  geom_step(aes(color = Scenario, y = On.Hand, size = Scenario), 
            alpha = .8) +
  scale_y_continuous(name = "On Hand Tons", labels = comma, 
                     limits = c(NA, 50000)) +
  ggtitle("Simulation Results and QlikView Actuals") +
  #scale_color_brewer(palette = "Set1") +
  geom_line(aes(color = Scenario, y = Avg.OH), size = .1) +
  scale_size_manual(values = c(.2, .2, 1.5))
  #geom_label(aes(x = max(Date), y = Avg.OH, label = 
  #                 format(Avg.OH, big.mark = ",")), hjust = "inward") 
  #geom_smooth(aes(y = On.Hand, color = Scenario), se = FALSE,
  #            span = 0.2)
  
