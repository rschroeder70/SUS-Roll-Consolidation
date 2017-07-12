library(tidyverse)
library(stringr)
library(jsonlite)
library(scales)
library(openxlsx)

#---------- Set discrete color pallet
library(RColorBrewer)
qual_col_pals = brewer.pal.info[brewer.pal.info$category == 'qual',]
col_vector = unlist(mapply(brewer.pal, qual_col_pals$maxcolors, 
                           rownames(qual_col_pals)))

#-------------------
# Build the roll set as input to the trim optimization
#
trim_scenario <-"VMI-98-8-150-725"
trim_year <- "2016"

trim_output <- file.path(proj_root,"Scenarios", trim_year, trim_scenario)

file_list <- list.files(trim_output, ".Sim Results.")
file_list <- file.path(trim_output, file_list)

trim_df <- map_df(file_list, read_csv)
trim_df <- trim_df %>%
  mutate(Caliper = str_sub(CalDW, 1, 2))

trim_fac <- "W6"
trim_grade <- "AKPG"
trim_cal <- "24"
# Avg Wt of a 62.875 77" OD roll is 3.954
tpi <- 3.954/62.875
dmd_weeks <- 2

# Filter to a single machine, grade & Caliper
trim_df <- trim_df %>%
  filter(Fac == trim_fac & Grade == trim_grade & Caliper == trim_cal)

# Compute the number of rolls
trim_df_pre <- trim_df %>%
  select(Prod.Width, Prod.Gross.Tons) %>%
  mutate(Prod.Width.mm = ceiling(Prod.Width * 25.4)) %>%
  mutate(Prod.Rolls = ceiling((Prod.Gross.Tons / (52 / dmd_weeks)) / 
                                (Prod.Width * tpi))) %>%
  arrange(Prod.Width)

pm_width_in <- 226.5625
pm_width <- floor(pm_width_in * 25.4)

copy.table(trim_df_pre)

#--------------------
#
#  Or read the trim data from an input file
#
fp <- file.path("G:", "000_Supply Chain", "_Optimization",
                "Model Optimization", "Projects", "2017 Projects",
                "Trim Testing - RSC", 
                "Pre-Consolidation Trim Test - Expanded.xlsm")

trim_data <- read.xlsx(fp, sheet = "Input", startRow = 2, 
                       cols = 1:2)

trim_df_pre <- trim_data %>%
  rename(Prod.Width = Widths, Prod.Rolls = Demands) %>%
  mutate(Prod.Width.mm = ceiling(Prod.Width * 25.4)) 

#---------- Read json output merge sizes from input file and plot

jf <- file.path("C:","Users", "SchroedR", "workspace",
                "cspsol", "data",
                "solution.json")

# The trim program sometimes inserts a row with just a comma in row 5
# Delete it
jf_fix <- scan(jf, what = "", sep = "\n")
if (jf_fix[5] == ",") jf_fix <- jf_fix[-5]
write(jf_fix, jf, sep = "\n")


sol <- fromJSON(jf_fix)
sol[["run"]][["Solution"]][["pattern"]][[1]]
sum(sol[["run"]][["Solution"]][["pattern"]][[1]])


num_pattern <- sol$run$Optimal

sol_df <- data_frame(pattern = sol$run$Solution$pattern, 
           pwidth = map_int(sol$run$Solution$pattern, sum),
           pcount = sol$run$Solution$count)

# add the widths in inches
#sol_df$pattern[1]
#trim_df_pre$Prod.Width[match(sol_df$pattern[[1]], trim_df_pre$Prod.Width.mm)]
# First create a function to look up the widths in inches
mfun <- function(x) {
  trim_df_pre$Prod.Width[match(x, trim_df_pre$Prod.Width.mm)]
}
sol_df$pattern_in <- map(sol_df$pattern, mfun)

sol_df <- sol_df %>%
  mutate(pwidth_in = map_dbl(pattern_in, sum)) %>%
  mutate(dwidth_in = pm_width_in - pwidth_in) %>%
  arrange(desc(pcount)) %>%
  mutate(pnum = row_number())

trim_width <- sol_df %>% 
  summarize(Total.Width = sum(pcount * (pm_width_in - dwidth_in)))

max_width <- num_pattern * pm_width_in

delta_width <- max_width - trim_width
delta_width_pct <- delta_width / max_width

rolls <- sol$run$SolutionWidth
# Get width inches
rolls$width_in <- 
  trim_df_pre$Prod.Width[match(rolls$width, trim_df_pre$Prod.Width.mm)]
  
extra_width <- rolls %>%
  summarize(extra_width = sum(width_in * (solution - demand)))
extra_width_pct <- extra_width / max_width

total_trim_pct <- delta_width_pct + extra_width_pct

trim_data <-
  tribble( ~ Pattern, ~Roll.Width, ~Width, ~Pattern.Count,
           #----------------------------------------
           "1",   "43.125",      	43.125, 2,
           "1",   "43.125",      	43.125, 2,
           "1",   "43.125",      	43.125, 2,
           "1",   "43.125",      	43.125, 2,
           "1",   "43.125",      	43.125, 2,
           "2",   "41.875",       41.875, 3,
           "2",   "41.875",       41.875, 3,
           "2",   "41.875",       41.875, 3,
           "2",   "41.875",       41.875, 3,
           "2",   "41.875",       41.875, 3,
           "3",   "64.25",        64.25,  10,
           "3",   "62.875",       62.875, 10,
           "3",   "62.875",       62.875, 10,
           "3",   "37.75",        37.75, 10)


data.frame(sol_df$pattern_in[1], sol_df$pcount[1], sol_df$pnum[1])

sol_df_plot <- 
  data.frame(Pattern = rep(sol_df$pnum, map_int(sol_df$pattern_in,length)),
             Pattern.Count = rep(sol_df$pcount, map_int(sol_df$pattern_in, 
                                                        length)), 
             Width = unlist(sol_df$pattern_in),
             Pattern.Delta = rep(sol_df$dwidth_in, map_int(sol_df$pattern_in, 
                                                        length)))

sol_df_plot <- sol_df_plot %>%
  mutate(Roll.Width = as.character(Width))

# scale the widths so that the max is 5 and the min is 1
width_range <- range(sol_df_plot$Pattern.Count)
sol_df_plot <- sol_df_plot %>%
  mutate(line.width = 1.2 - (1.2 - .3) *
           (width_range[2] - Pattern.Count) / 
           (width_range[2] - width_range[1]))

ggplot(data = sol_df_plot, aes(x = Pattern, y = Width)) +
  geom_col(aes(fill = Roll.Width, width = line.width), 
           color = "black", stat = "identity", alpha = .9) +
  geom_text(aes(label = Roll.Width), position = position_stack(vjust = 0.5),
            size = 4) +
  geom_hline(aes(yintercept = pm_width_in), color = "dark green",
             size = 1.3) +
  geom_text(aes(label = Pattern.Delta, y = pm_width_in + 10, 
            x = Pattern)) +
  geom_text(aes(label = Pattern.Count, x = Pattern, y = - 10)) +
  geom_text(aes(label = "Pattern\nCount", y = -10, 
                x = max(sol_df_plot$Pattern) + 1)) +
  coord_flip() +
  theme(text = element_text(size = 12), legend.position = "none") +
  ggtitle(paste0("Total Sets = ", num_pattern, 
                 ", Width Efficiency Loss = ", percent(delta_width_pct[[1]]),
                 ", Unique Patterns = ", max(sol_df_plot$Pattern))) +
  scale_x_discrete(labels = NULL) +
  scale_fill_manual(values = rev(col_vector))

rm(trim_scenario, trim_detail, trim_year, trim_output, file_list,
   trim_df, trim_df_pre, trim_fac, trim_grade, trim_cal, tpi,
   dmd_weeks, sol_df_plot, fp, trim_data)

