################################################################################
#
# Read the QlikView daily inventory file.  The data comes from the Inventory
# cube, and the SUS Rollstock bookmark.  Select Daily Trend on the Inventory
# Overview tab ahd then Daily Trand - Pivot on the Inventory Details tab.  
# Expand all columns
#
library(tidyverse)
library(openxlsx)
library(stringr)
library(ggthemes)

#---------- Create the bar chart color vector
library(RColorBrewer)
qual_col_pals = brewer.pal.info[brewer.pal.info$category == 'qual',]
col_vector = unlist(mapply(brewer.pal, qual_col_pals$maxcolors, 
                           rownames(qual_col_pals)))


sinv <- read.xlsx("Data/QV-SUS-board-daily-inventory.xlsx",
                  detectDates = T, 
                  check.names = T)

names(sinv)[1:8] <- make.names(sinv[1,1:8])
sinv <- sinv %>% filter(Plant != "Plant")
# Convert to long format
sinv <- sinv %>% gather(myattribute, myvalue, 9:ncol(sinv))

# Id myattribute ends in a ".1" it's a cost (not ton) and we can delete
sinv <- sinv %>% filter(str_sub(myattribute, -2, -1) != ".1")
sinv <- sinv %>% filter(!is.na(myvalue))

sinv <- sinv %>% rename(Date = myattribute,
                            Tons = myvalue) %>%
  mutate(Tons = round(as.numeric(Tons), 1))

sinv <- mutate(sinv, Date = as.Date(substr(Date, 2, 11), "%m.%d.%Y"))

# Filter out data we don't want
sinv <- sinv %>% filter(!is.na(Material.Description))
sinv <- sinv %>% filter(str_sub(Material, 1, 1) == "1")
sinv <- sinv %>%
  mutate(M.I = str_sub(Material.Description, 1, 1)) %>%
  filter(M.I %in% c("M", "I"))

# Get the Grade, Caliper, Dia & Wind out of the material description
sinv <- sinv %>% 
  mutate(Width = ifelse(M.I == "I",
                        str_sub(Material.Description, 10, 17),
                        str_sub(Material.Description, 11, 14)))
sinv$Width <- as.numeric(sinv$Width)
sinv <- sinv %>% filter(!is.na(Width))

sinv <- sinv %>%
  mutate(Caliper = ifelse(M.I == "I",
                          str_sub(Material.Description, 3, 4),
                          str_sub(Material.Description, 4, 5)))
sinv <- sinv %>%
  mutate(Grade = ifelse(M.I == "I",
                        str_sub(Material.Description, 5, 8),
                        str_sub(Material.Description, 6, 9)))

sinv <- sinv %>%
  mutate(Wind = str_sub(Material.Description, 
                        str_length(Material.Description), -1))

sinv <- sinv %>%
  mutate(Dia = ifelse(M.I == "I",
                      str_sub(Material.Description, 19, 20),
                      str_sub(Material.Description, 16, 19)))

bev.elim <- 
  tribble( 
    ~Caliper, ~Grade, ~Width.e, ~Width.p,
    #-----------------------
    "18", "AKPG", 29.5, 31.125,
    "18", "AKPG", 30.125, 31.125,
    "18", "AKPG", 44.125, 45.25,
    "21", "AKPG", 45.5, 46,
    "24", "AKPG", 36.75, 37.75,
    #"24", "AKPG", 41.875, 42.25,
    #"24", "AKPG", 42.125, 42.25,
    "24", "AKPG", 42.875, 43.125,
    "24", "AKPG", 44.875, 46.625,
    "24", "AKPG", 47.5, 47.625,
    "24", "AKPG", 52.875, 54.5,
    "27", "AKPG", 40.875, 42.375,
    "27", "AKPG", 45.75, 47.125,
    "27", "AKPG", 46.625, 47.125,
    "28", "PKXX", 37.875, 38.625)

sinv.bev.elim <-
  semi_join(sinv, bev.elim,
            by = c("Caliper", "Grade", "Width" = "Width.e"))

sinv.bev.elim <- sinv.bev.elim %>%
  group_by(Date, Caliper, Grade, Width) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup() %>%
  unite(Cal.Gr.W, Caliper, Grade, Width)

# Plot the rolls to be eliminated
ggplot(data = sinv.bev.elim, aes(x = Date)) +
  geom_bar(aes(y = Tons, fill = Cal.Gr.W), stat = "identity") +
  scale_fill_manual(values = (col_vector)) +
  ggtitle("Inventory of Eliminated Beverage VMI Rolls")

sinv.bev.elim %>% group_by(Date) %>%
  summarize(Tons = sum(Tons)) %>%
  filter(Date >=  max(Date)-15)

sinv.bev.parent <- 
  semi_join(sinv, elim,
            by = c("Caliper", "Grade", "Width" = "Width.p"))

# Plot inventory of all Parent rolls
sinv.bev.parent <- sinv.bev.parent %>%
  group_by(Date, Caliper, Grade, Width) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup() %>%
  unite(Cal.Gr.W, Caliper, Grade, Width)

ggplot(data = sinv.bev.parent, aes(x = Date)) +
  geom_bar(aes(y = Tons, fill = Cal.Gr.W), stat = "identity") +
  scale_fill_manual(values = (col_vector)) +
  ggtitle("Inventory of Parent Beverage VMI Rolls")

sinv.bev.parent %>% group_by(Date) %>%
  summarize(Tons = sum(Tons)) %>%
  filter(Date >=  max(Date)-15)

