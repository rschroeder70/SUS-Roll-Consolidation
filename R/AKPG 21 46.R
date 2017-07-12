library(openxlsx)
library(tidyverse)
library(stringr)
source("~/Projects/RollAssortmentOptimization/R/0.Functions.R")


# Hard code these for now.  Better to read them from the data
max_date <- as.Date("2016-12-31")
min_date <- as.Date("2016-01-01")

# Pull in the production data from Paul
prod_hist <- read.xlsx("G:/SUS Planning Team/Paul files/For Rob/2016 Production.xlsx")
prod_hist <- prod_hist %>%
  filter(Mill.machine == "WM-06" & Cal == 21 & Grade == "AKPG")

prod_hist$MFG.Date <- as.Date(prod_hist$MFG.Date, origin = "1899-12-30")

prod_hist <- prod_hist %>%
  mutate(Caliper = ifelse(str_sub(Material.Description, 1, 1) == "I",
                                 str_sub(Material.Description, 3, 4),
                                 str_sub(Material.Description, 4, 5)))
prod_hist <- prod_hist %>%
  mutate(Width = ifelse(str_sub(Material.Description, 1, 1) == "I",
                         str_sub(Material.Description, 10, 17),
                         str_sub(Material.Description, 11, 14)))
prod_hist$Width <- as.numeric(prod_hist$Width)
prod_hist$Weight <- as.numeric(prod_hist$Weight)

prod_hist <- prod_hist %>%
  mutate(Dia = ifelse(str_sub(Material.Description, 1, 1) %in% c("I", "("),
                         str_sub(Material.Description, -7, -5),
                         str_sub(Material.Description, -10, -6)))

prod_hist <- prod_hist %>%
  filter(Width %in% c(43.25, 45.5, 46)) %>%
  mutate(Width = as.character(Width))

prod_hist <- prod_hist %>%
  group_by(Mill.machine, MFG.Date, Grade, Caliper, Width) %>%
  summarize(Tons = round(sum(Weight)/2000, 2)) %>%
  ungroup()


# The consumption history comes from QlikView Inventory /
# Fiber Analysis / Consumption by Die Report (Date = Week End) "VMI 46in Parent" bookmark
# Note that this does not include the mill/machine of manufacture
cons_hist <- read.xlsx("~/Projects/SUS Roll Consolidation/Data/AK 46 inch Consumption History.xlsx")
cons_hist$Week.End <- as.Date(cons_hist$Week.End, origin = "1899-12-30")
cons_hist <- cons_hist %>%
  filter(!is.na(Week.End))

cons_hist <- cons_hist %>%
  mutate(Grade = ifelse(str_sub(Material.Description, 1, 1) == "I",
                        str_sub(Material.Description, 5, 8),
                        str_sub(Material.Description, 6, 9)))

cons_hist <- cons_hist %>%
  group_by(Week.End, Grade, Board.Caliper, Board.Width) %>%
  summarize(Tons = sum(TON)) %>%
  ungroup()

cons_hist <- cons_hist %>%
  filter(between(Week.End, as.Date("2016-01-01"), as.Date("2016-12-31")))

# The inventory data comes from QlikView Inventory/Inventory / Inventory Details /
# "VMI 46 in Parent" Bookmark
inv_hist <- read.xlsx("~/Projects/SUS Roll Consolidation/Data/AK 46 inch Inv History.xlsx")
inv_hist <- inv_hist[as.vector(!(inv_hist[1,] == "# Batches" | inv_hist[1,] == "Cost Val"))]
nbr_cols <- ncol(inv_hist)
inv_hist <- inv_hist %>% gather(Range, Tons, 9:nbr_cols)
names(inv_hist)[1:8] <- inv_hist[1,1:8]
# Delete row 1
inv_hist <- inv_hist[-1,]
inv_hist <- inv_hist %>% rename(Material.Description = `Material Description`)
inv_hist[is.na(inv_hist)] <- 0
inv_hist <- inv_hist %>% filter(!(Plant %in% c("Plant", "Total")))
inv_hist$Tons <- round(as.numeric(gsub(",", "", inv_hist$Tons)), 2)
inv_hist$Range <- as.Date(inv_hist$Range, "%m/%d/%Y")
inv_hist <- inv_hist %>% 
  mutate(Width = ifelse(str_sub(Material.Description, 1, 1) == "I",
                        str_sub(Material.Description, 10, 17),
                        str_sub(Material.Description, 11, 14)))
inv_hist$Width <- as.character(as.numeric(inv_hist$Width))

inv_hist <- inv_hist %>%
  mutate(Caliper = ifelse(str_sub(Material.Description, 1, 1) == "I",
                          str_sub(Material.Description, 3, 4),
                          str_sub(Material.Description, 4, 5)))
inv_hist <- inv_hist %>%
  mutate(Grade = ifelse(str_sub(Material.Description, 1, 1) == "I",
                        str_sub(Material.Description, 5, 8),
                        str_sub(Material.Description, 6, 9)))

# Just look at 2016
inv_hist <- inv_hist %>%
  filter(between(Range, as.Date("2016-01-01"), as.Date("2016-12-31")))

inv_hist <- inv_hist %>%
  filter(Width != "45.875")

# Summarize over the widths
inv_hist <- inv_hist %>% 
  group_by(Range, Grade, Caliper, Width) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

# Average OH by width
inv_hist %>% group_by(Width) %>%
  summarize(Avg = sum(Tons)/12)

# Add the 45.5, 43.25 and 46 together
inv_sum <- inv_hist %>%
  filter(Width %in% c("45.5", "43.25", "46")) %>%
  group_by(Range) %>%
  summarize(Tons = sum(Tons)) %>%
  ungroup()

ggplot(data = inv_sum, aes(x = Range, y = Tons)) +
  geom_step(direction = "vh") +
  expand_limits(y = 0) +
  ggtitle(paste("24 pt AKPG 43.25 - 46 in.  Avg OH =", 
                format(mean(inv_sum$Tons), digits = 2, big.mark = ",")))

# Stacked bar (area)
ggplot(data = filter(inv_hist, Width %in% c("45.5", "43.25", "46")),
       aes(x = Range, y = Tons, fill = Width)) +
  geom_bar(stat = "identity", position = "stack") +
  geom_segment(aes(x = min(Range), xend = max(Range), y = mean(inv_sum$Tons),
                   yend = mean(inv_sum$Tons)), color = "red", size = 2) +
  ylab("Month Ending Tons") +
  xlab("Date") +
  ggtitle(paste("21 pt AKPG 43.25 - 46 in.  Avg OH =", 
                format(mean(inv_sum$Tons), digits = 2, big.mark = ",")))

# Combine the production and consumption records into a transaction table
# Make the column names the same for rbinding
cons_hist <- cons_hist %>%
  rename(Date = Week.End, Caliper = Board.Caliper,
         Width = Board.Width) %>%
  mutate(Tons = -1 * Tons) %>%
  mutate(type = "cons")
prod_hist <- prod_hist %>%
  rename(Date = MFG.Date) %>%
  mutate(type = "prod")
inv_hist <- inv_hist %>%
  select(Date = Range, Grade, Caliper, Width, Tons) %>%
  mutate(type = "inv")

trans <- bind_rows(cons_hist, prod_hist, inv_hist)

# Use the last date as the ending inventory
trans <- bind_rows(trans, filter(inv_hist, Date == max(Date)) %>%
                     mutate(type = "end"))

# Caluclate the starting inventory
trans <- bind_rows(
  trans,
  filter(trans, type != "inv") %>%
    mutate(Tons = ifelse(type != "end",-Tons, Tons)) %>%
    group_by(Grade, Caliper, Width) %>%
    summarize(Tons = sum(Tons)) %>%
    mutate(type = "start", Date = as.Date("2016-01-01"))
)

# Spread the data
trans <- trans %>%
  spread(type, Tons, fill = 0)

# Eliminate records with no data
trans <- trans %>%
  filter(!(cons == 0 & end == 0 & inv == 0 & prod == 0 & start == 0))

# Compute the net change each day
trans <- trans %>% group_by(Date, Grade, Caliper, Width) %>%
  mutate(change = sum(start + prod + cons))

# Compute the running balance
# We will use the lag fuction, so we need to seed the start row for each width
trans <- trans %>% 
  mutate(Bal = ifelse(Date == as.Date("2016-12-31"), end, 0))

trans <- trans %>% arrange(Width, desc(Date)) %>%
  group_by(Width) %>%
  mutate(Old.Bal = Bal,
         Lag.Bal = lag(Bal, default = 0),
         Cum.Sum.Bal = cumsum(Bal),
         Bal = ifelse(Date == as.Date("2016-12-31"), Bal,
                      cumsum(Bal) - cumsum(lag(change, default = 0))),
         Lag.Cum.Sum.Change = cumsum(lag(change, default = 0))) %>%
  ungroup()
  
# Delete the working columns
trans <- trans %>%
  select(-(Old.Bal:Lag.Cum.Sum.Change))

trans <- trans %>%
  arrange(Width, Date)

brd_stats <- trans %>%
  group_by(Width) %>%
  mutate(
    diffx = Date - lag(Date, default = min(Date)),
    ylag = lag(Bal, default = 0),
    auc = diffx * ylag,
    Error = ifelse(!is.na(inv), abs(inv - Bal), 0)) %>%
  summarize(auc = sum(auc),
            Sum.Cons = -sum(cons),
            Sum.Prd = sum(prod),
            Max.OH = max(Bal),
            Error = max(Error))

# Figure out the actual start and end dates for each item 
# Assume that we are at zero if within 500 cartons
first_dates <- trans %>%
  group_by(Width) %>%
  filter(Bal > 50) %>%
  summarize(mindate = min(Date))

# Get the last date where either the balance drops to zero,
# or the max date if it's > 0
last_dates <- trans %>%
  group_by(Width) %>%
  filter((Bal < 50 & lag(Bal > 50)) | 
           (Date == max_date & Bal > 50)) %>%
  summarize(maxdate = max(Date))

active_range <- full_join(first_dates, last_dates,
                          by = c("Width" = "Width"))
rm(first_dates, last_dates)
active_range <- active_range %>%
  mutate(Date.Range = as.numeric(maxdate - mindate))

brd_stats <- left_join(brd_statsa, active_range,
                        by = c("Width" = "Width"))

brd_stats <- brd_statsa %>%
  mutate(Avg.OH = as.numeric(round(auc/Date.Range, 0)),
         Avg.DIS = round(Date.Range * Avg.OH / Sum.Cons, 1))



plot_title = paste("Inventory Profile for", max(trans$Grade), max(trans$Caliper))
ggplot(ptrans, aes(x = Date)) +
  geom_step(aes(y = Bal, color = Width), direction = "hv", size = 1) +
  #geom_step(data = filter(trans, inv > 0), aes(y = inv, color = Width), 
  #          direction = "hv", size = .75, linetype = "dashed") +
  ggtitle(plot_title) +
  geom_point(data = filter(trans, inv > 0), aes(y = inv, color = Width), size = 2, shape = 8)



# Try to fix this by adjusting the inventory based on the month inv errors
# Determin adjustment transactions
adj_trans <- trans %>%
  filter(inv != 0) %>%
  mutate(Error = inv - Bal) %>%
  select(Date, Grade, Caliper, Width, Error) %>%
  group_by(Width) %>%
  mutate(type = "adj",
         Tons = (lag(Error) - Error)) %>%
  select(-Error) %>%
  filter(!is.na(Tons))
  
# Add the adjustment transactions and perform the same calculations
transa <- bind_rows(cons_hist, prod_hist, inv_hist, adj_trans)

# Use the last date as the ending inventory
transa <- bind_rows(transa, filter(inv_hist, Date == max(Date)) %>%
                     mutate(type = "end"))

# Caluclate the starting inventory
transa <- bind_rows(
  transa,
  filter(transa, type != "inv") %>%
    mutate(Tons = ifelse(type != "end",-Tons, Tons)) %>%
    group_by(Grade, Caliper, Width) %>%
    summarize(Tons = sum(Tons)) %>%
    mutate(type = "start", Date = as.Date("2016-01-01"))
)

# Spread the data
transa <- transa %>%
  spread(type, Tons, fill = 0)

# Eliminate records with no data
transa <- transa %>%
  filter(!(cons == 0 & end == 0 & inv == 0 & prod == 0 & start == 0))

# Compute the net change each day
transa <- transa %>% group_by(Date, Grade, Caliper, Width) %>%
  mutate(change = sum(start + prod + cons + adj))

# Compute the running balance
# We will use the lag fuction, so we need to seed the start row for each width
transa <- transa %>% 
  mutate(Bal = ifelse(Date == as.Date("2016-12-31"), end, 0))

transa <- transa %>% arrange(Width, desc(Date)) %>%
  group_by(Width) %>%
  mutate(Old.Bal = Bal,
         Lag.Bal = lag(Bal, default = 0),
         Cum.Sum.Bal = cumsum(Bal),
         Bal = ifelse(Date == as.Date("2016-12-31"), Bal,
                      cumsum(Bal) - cumsum(lag(change, default = 0))),
         Lag.Cum.Sum.Change = cumsum(lag(change, default = 0))) %>%
  ungroup()

# Delete the working columns
transa <- transa %>%
  select(-(Old.Bal:Lag.Cum.Sum.Change))


# Compute the average inventory by item
transa <- transa %>%
  arrange(Width, Date)

brd_statsa <- transa %>%
  group_by(Width) %>%
  mutate(
    diffx = Date - lag(Date, default = min(Date)),
    ylag = lag(Bal, default = 0),
    auc = diffx * ylag,
    Error = ifelse(!is.na(inv), abs(inv - Bal), 0)) %>%
  summarize(auc = sum(auc),
            Sum.Cons = -sum(cons),
            Sum.Prd = sum(prod),
            Max.OH = max(Bal),
            Error = max(Error))

# Figure out the actual start and end dates for each item 
# Assume that we are at zero if within 500 cartons
first_dates <- transa %>%
  group_by(Width) %>%
  filter(Bal > 50) %>%
  summarize(mindate = min(Date))

# Get the last date where either the balance drops to zero,
# or the max date if it's > 0
last_dates <- transa %>%
  group_by(Width) %>%
  filter((Bal < 50 & lag(Bal > 50)) | 
           (Date == max_date & Bal > 50)) %>%
  summarize(maxdate = max(Date))

active_range <- full_join(first_dates, last_dates,
                          by = c("Width" = "Width"))
rm(first_dates, last_dates)
active_range <- active_range %>%
  mutate(Date.Range = as.numeric(maxdate - mindate))

brd_statsa <- left_join(brd_statsa, active_range,
                       by = c("Width" = "Width"))

brd_statsa <- brd_statsa %>%
  mutate(Avg.OH = as.numeric(round(auc/Date.Range, 0)),
         Avg.DIS = round(Date.Range * Avg.OH / Sum.Cons, 1))


#ptrans <- filter(trans, Width == "46") %>%
#  mutate(inv = ifelse(inv == 0, NA, inv))
plot_title = paste("Inventory Profile for", max(trans$Grade), max(trans$Caliper))
ggplot(ptrans, aes(x = Date)) +
  geom_step(aes(y = Bal, color = Width), direction = "hv", size = 1) +
  #geom_step(data = filter(trans, inv > 0), aes(y = inv, color = Width), 
  #          direction = "hv", size = .75, linetype = "dashed") +
  ggtitle(plot_title) +
  geom_point(data = filter(trans, inv > 0), aes(y = inv, color = Width), size = 2, shape = 8)
