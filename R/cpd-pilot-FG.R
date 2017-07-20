###############################################################################
#
# Read the finished goods inventory for the CPD Pilot Dies
#
# Data comes from QlikView; Inventory / Inventory Details /
# Monthly Trend Details from the CPD Pilot Finished Goods by Die bookmark
#

cpdfg <- read.xlsx("Data/QV-FG-Inv-CPD-pilot-dies.xlsx",
                   detectDates = TRUE)

parent.dies <- c("38584A", "26474B", "6335AN", "21104B", "40461C", "33759C", 
                 "37802B", "16599H", 
                 "39360D", "32861J", "39135A")

cpdfg$Date <- as.Date(cpdfg$Date, "%m/%d/%Y")

cpdfg.ytd <- cpdfg %>% 
  filter(Date >= as.Date("2017/01/01") & 
           Date <= as.Date("2017/06/30"))

cpdfg.ytd %>% summarize(SKUs = n_distinct(Material))

# Summarize FG by roll Die
cpdfg.ytd %>% group_by(Die) %>%
  filter(!Die %in% parent.dies) %>%
  summarize(SKUs = n_distinct(Material))

# Average month end inventory by SKU
cpdfg.ytd.sku <- cpdfg.ytd %>% group_by(Material) %>%
  summarize(Avg.Mtl.Tons = mean(Qty.in.TON)) %>%
  arrange(desc(Avg.Mtl.Tons))

# YTD Average of all SKUs
cpdfg.ytd.sku %>% summarize(Tons = sum(Avg.Mtl.Tons))

# Snapsot Summary by Plant

