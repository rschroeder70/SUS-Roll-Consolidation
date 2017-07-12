#
# convert the roll widths in inches to mm
#
# Load Packages & Functions
#source(file.path(basepath, "R", "0.LoadPkgsAndFunctions.R"))
#library(openxlsx)
source("R/FuncReadResults.R")
library(jsonlite)

# This projects root directory
basepath <- file.path("C:","Users", "SchroedR", "Documents",
                      "Projects", "RollAssortmentOptimization")
proj_root <- file.path("C:", "Users", "SchroedR", "Documents",
                       "Projects", "SUS Roll Consolidation") 
# Load Packages & Functions
source(file.path(basepath, "R", "0.LoadPkgsAndFunctions.R"))

cspsol_exe <- file.path("C:", "Users", "SchroedR", "workspace", "cspsol",
                        "Debug")
cspsol_dat <- file.path("C:", "Users", "SchroedR", "workspace", "cspsol",
                        "data")

#---------- Create the bar chart color vector
library(RColorBrewer)
qual_col_pals = brewer.pal.info[brewer.pal.info$category == 'qual',]
col_vector = unlist(mapply(brewer.pal, qual_col_pals$maxcolors, 
                           rownames(qual_col_pals)))
#---------- Load these constants
pm_width <- c("W6" = 226.5625, "W7" = 234)
pm_width_16 <- pm_width * 16
#pm_width_mm <- floor(pm_width * 25.4)
#tfp <- file.path("C:","Users", "SchroedR", "workspace",
#               "cspsol", "data")

#--------- Read trim history file from John Zopf
wpi <- tribble(~Machine, ~Grade, ~Caliper, ~Dia, ~WPI,
               #------------------------------
               "W7", "AKPG", "18", "60", 80.16,
               "W7", "AKPG", "18", "69", 110.00,  #Guess
               "W7", "AKPG", "18", "72", 115.68,
               "W7", "AKPG", "18", "77", 132.386,
               "W7", "AKPG", "18", "81", 146.557,
               "W7", "AKPG", "18", "84", 157.651,
               "W6", "AKPG", "21", "60", 76.137,
               "W6", "AKPG", "21", "72", 109.873,
               "W6", "AKPG", "21", "77", 125.74,
               "W6", "AKPG", "21", "81", 139.195,
               "W6", "AKPG", "21", "84", 149.739,
               "W6", "AKPG", "24", "60", 75.555,
               "W6", "AKPG", "24", "72", 109.035,
               "W6", "AKPG", "24", "77", 124.783,
               "W6", "AKPG", "24", "81", 138.136,
               "W6", "AKPG", "24", "84", 148.598,
               "W6", "AKPG", "27", "60", 75.828,
               "W6", "AKPG", "27", "72", 109.428,
               "W6", "AKPG", "27", "77", 125.23,
               "W6", "AKPG", "27", "81", 138.365,
               "W6", "AKPG", "27", "84", 149.135,
               "W7", "PKXX", "18", "60", 78.847,
               "W7", "PKXX", "18", "72", 113.781)


vmi_opt <- read_results("opt_detail", "2016", "VMI-98-8-150-500v3")

trim_fp <- file.path("C:", "Users", "SchroedR", "Documents",
                     "Projects", "SUS Roll Consolidation",
                     "trim-analysis")

file_list <- list.files(trim_fp, "runs")
file_list <- file.path(trim_fp, file_list)

# use map_df from the purrr library
trim_hist <- map_df(file_list, read.xlsx)
trim_hist$Caliper <- str_sub(trim_hist$grade_spec, -2)
trim_hist$Grade <- "AKPG"
trim_hist$Machine <- paste0(ifelse(trim_hist$mill_id == "031", "W", "M"), 
                                   str_sub(trim_hist$machine_id, 2, 2))

trim_hist <- trim_hist %>%
  mutate(diam_ord = as.character(diam_ord),
         run_num = as.character(run_num))

trim_hist <- trim_hist %>% 
  filter(wgt_plan > 0) %>%
  select(Machine, Grade, Caliper, run_num, diam_ord, pat_num, 
         width_act, wgt_plan) %>%
  left_join(wpi,
            by = c("Machine", "Grade", "Caliper", "diam_ord" = "Dia"))

trim_hist <- trim_hist %>%
  mutate(nbr_sets = round(wgt_plan / (width_act * WPI), 0))

# Add the paper machine width as a field
trim_hist$pm_width <- pm_width[trim_hist$Machine]

# Create an index for the position of the roll in the pattern
trim_hist <- trim_hist %>%
  group_by(run_num, pat_num) %>%
  mutate(pos_nbr = row_number()) %>%
  ungroup()

trim_hist <- trim_hist %>%
  group_by(run_num, pat_num) %>%
  mutate(pat_width = sum(width_act),
         pat_gap = pm_width - sum(width_act)) %>%
  ungroup()

trim_hist <- trim_hist %>%
  group_by(run_num) %>%
  mutate(run_width = sum(width_act * nbr_sets) / 
           sum(ifelse(pos_nbr == 1, nbr_sets, 0)))

# Add the parent roll widths
trim_hist <- 
  left_join(trim_hist, 
            select(vmi_opt, Fac, Grade, Caliper, 
                   Prod.Width, Parent.Width),
            by = c("Grade", "Caliper", "width_act" = "Prod.Width",
                   "Caliper" = "Caliper", "Machine" = "Fac")) %>%
  mutate(width_adj = ifelse(is.na(Parent.Width), width_act, Parent.Width),
         has_parent = ifelse(width_act != width_adj, TRUE, FALSE))

# Compute the new trim width gap from max
trim_hist <- trim_hist %>%
  group_by(run_num, pat_num) %>%
  mutate(pat_width_adj = sum(width_adj),
         pat_gap_adj = pm_width - sum(width_adj)) %>%
  ungroup()

trim_hist <- trim_hist %>%
  group_by(temp = paste(run_num, pat_num)) %>%
  mutate(pat_has_parent = ifelse(max(has_parent) >0, T, F)) %>%
  ungroup() %>%
  select(-temp)
  
trim_hist <- trim_hist %>%
  group_by(run_num) %>%
  mutate(run_has_parent = ifelse(max(has_parent) > 0, T, F),
         run_re_trim = 
           ifelse(max(has_parent) > 0 & min(pat_gap_adj) < 0, T, F)) %>%
  ungroup()

trim_sum <- trim_hist %>% group_by(Caliper, diam_ord) %>%
  summarize(Runs = n_distinct(run_num),
            Patterns = n_distinct(paste(run_num, pat_num)),
            Sets = sum(ifelse(pos_nbr == 1, 
                              nbr_sets, 0)),
            Plan.Tons = round(sum(wgt_plan/2000), 0),
            Avg.Set.Gap = sum(ifelse(pos_nbr == 1, pat_gap * nbr_sets, 0)) /
              sum(ifelse(pos_nbr == 1, nbr_sets, 0)),
            Set.Gap.Pct = percent(sum(ifelse(pos_nbr == 1, pat_gap * nbr_sets, 0)) /
                                    sum(ifelse(pos_nbr == 1, nbr_sets, 0))/
              max(pm_width)),
            Parent.Tons = sum(WPI * nbr_sets * 
                                (Parent.Width - width_act), na.rm = T)/2000,
            Pattern.Parents = sum(ifelse(pos_nbr == 1 & pat_has_parent, 
                                         1, 0)),
            sets.Parents = 
              sum(ifelse(pos_nbr == 1 & pat_has_parent, 
                         nbr_sets, 0)),
            Sets.Exceed.Width = 
              sum(ifelse(pos_nbr == 1 & pat_has_parent & pat_gap_adj < 0, 
                         nbr_sets, 0)),
            ReTrim.Runs = 
              sum(ifelse(pat_num == 1 & pos_nbr == 1 & run_re_trim, 1, 0)),
            ReTrim.Sets = sum(ifelse(pos_nbr == 1 & run_re_trim, nbr_sets, 0))) %>%
  ungroup()

copy.table(trim_sum)

# Tons of new width parent rolls
trim_hist %>% 
  filter(Parent.Width != width_act) %>%
  summarize(Tons.Increase = sum(WPI * nbr_sets * (Parent.Width - width_act))/2000)
  
# Write this back to the excel file
# wb <- loadWorkbook(file.path(trim_fp, trim_file))
# 
# if (!("Trim Set Details" %in% names(wb))) {
#   addWorksheet(wb, "Trim Set Details")
# } else {
#   removeWorksheet(wb, "Trim Set Details")
#   addWorksheet(wb, "Trim Set Details")
# }
# 
# writeData(wb, "Trim Set Details", trim_hist,
#           startCol = 1, startRow = 1,
#           withFilter = TRUE)
# setColWidths(wb, "Trim Set Details", cols = ncol(trim_hist), widths = "auto")
# freezePane(wb, "Trim Set Details", firstActiveRow = 2)
# 
# saveWorkbook(wb, file = file.path(trim_fp, "6pm 24pnt runs.xlsx"), 
#              overwrite = TRUE)
# 
trim_hist_adj <- trim_hist %>%
  filter(as.numeric(diam_ord) > 60) %>%
  group_by(Caliper, run_num, pat_num, diam_ord) %>%
  summarize(has_parent = max(has_parent), 
            nbr_sets = max(nbr_sets),
            pat_width = max(pat_width),
            pat_gap = max(pat_gap),
            pat_width_adj = max(pat_width_adj),
            pat_gap_adj = max(pat_gap_adj)) %>%
  ungroup()

trim_hist_adj %>% filter(has_parent == 1) %>%
  summarize(total_sets = sum(nbr_sets),
            re_trim_sets = sum(ifelse(pat_gap_adj < 0, nbr_sets, 0)))

copy.table(trim_hist_adj)

# How many runs with parent rolls?
trim_hist_adj %>%
  filter(has_parent == 1) %>%
  distinct(run_num) %>% 
  summarize(count = n())

# How many runs have a trim that is too wide
re_trims <- trim_hist_adj %>%
  filter(has_parent == 1) %>%
  group_by(run_num, Caliper, diam_ord) %>%
  summarize(Sets = sum(nbr_sets),
            Max = max(pat_width_adj)) %>%
  mutate(Re.Trim = ifelse(Max > max(pm_width), T, F))

re_trims <- trim_hist %>%
  filter(run_re_trim & Caliper == "18") %>%
  distinct(diam_ord, run_num)

# Get the rolls for a trim
run_number <- re_trims$run_num[1]
run_number <- "63062"

trim_data <- trim_hist %>%
  filter(run_num == run_number) %>%
  select(Width = width_act, Width.Parent = width_adj, nbr_sets) %>%
  group_by(Width, Width.Parent) %>%
  summarize(Qty.Req = sum(nbr_sets)) %>%
  ungroup()

# What machine are we runnin on
trim_mach <- trim_hist$Machine[trim_hist$run_num == run_number][1]

#trim_data$Width.mm <- ceiling(trim_data$Width * 25.4)
trim_data$Width.16 <- trim_data$Width * 16
trim_data$Width.Parent.16 <- trim_data$Width.Parent * 16

# Create the data file for the cspsol run
write.table(pm_width_16[trim_mach],file.path(cspsol_dat, "trim_orig.txt"),
            row.names = FALSE, col.names = FALSE)
write.table(select(trim_data, Width.16, Qty.Req), 
            file.path(cspsol_dat, "trim_orig.txt"), append = TRUE,
            row.names = FALSE, col.names = FALSE)

# Run the cspsol trim software
# If it takes too long, insert the '--cgroot' option
shell(cmd = paste0(file.path(cspsol_exe, "cspsol.exe"), " --data ", 
                   file.path(cspsol_dat, "trim_orig.txt")), translate = TRUE)

# Copy the json file to the data directory
file.copy(from = file.path(proj_root, "solution.json"),
                    to = file.path(cspsol_dat, "solution_orig.json"), overwrite = TRUE)

# Run with the parent roll widths
write.table(pm_width_16[trim_mach],file.path(cspsol_dat, "trim_parent.txt"),
            row.names = FALSE, col.names = FALSE)
write.table(summarize(group_by(trim_data, Width.Parent.16), 
                      Qty.Req = sum(Qty.Req)),
            file.path(cspsol_dat, "trim_parent.txt"), append = TRUE,
            row.names = FALSE, col.names = FALSE)
shell(cmd = paste0(file.path(cspsol_exe, "cspsol.exe"), " --data ", 
                   file.path(cspsol_dat, "trim_parent.txt")), translate = TRUE)
file.copy(from = file.path(proj_root, "solution.json"),
          to = file.path(cspsol_dat, "solution_parent.json"), overwrite = TRUE)


#---------- Read results from  json files
jf_orig <- file.path(cspsol_dat, "solution_orig.json")

# The trim program sometimes inserts a row with just a comma in row 5
# Delete it
jf_fix <- scan(jf_orig, what = "", sep = "\n")
if (jf_fix[5] == ",") jf_fix <- jf_fix[-5]
write(jf_fix, file = jf_orig, sep = "\n")

sol <- fromJSON(jf_orig)

num_pattern <- sol$run$Optimal

sol_df <- data_frame(pattern = sol$run$Solution$pattern, 
                     #pwidth = map_int(sol$run$Solution$pattern, sum),
                     pcount = sol$run$Solution$count)

sol_df$pattern <- map(sol_df$pattern, ~ .x / 16)

sol_df <- sol_df %>%
  mutate(pwidth = map_dbl(pattern, sum)) %>%
  mutate(dwidth = pm_width[trim_mach] - pwidth) %>%
  arrange(desc(pcount)) %>%
  mutate(pnum = row_number())

trim_width <- sol_df %>% 
  summarize(Total.Width = sum(pcount * (pm_width[trim_mach] - dwidth)))

max_width <- num_pattern * pm_width[trim_mach]

delta_width <- max_width - trim_width
delta_width_pct <- delta_width / max_width
avg_width <- with(sol_df, weighted.mean(pwidth, pcount))

rolls <- sol$run$SolutionWidth
rolls <- rolls %>%
  mutate(width = width / 16)

extra_width <- rolls %>%
  summarize(extra_width = sum(width * (solution - demand)))
extra_width_pct <- extra_width / max_width

total_trim_pct <- delta_width_pct + extra_width_pct

sol_df_plot <- 
  data.frame(Pattern = rep(sol_df$pnum, map_int(sol_df$pattern,length)),
             Pattern.Count = rep(sol_df$pcount, map_int(sol_df$pattern, 
                                                        length)), 
             Width = unlist(sol_df$pattern),
             Pattern.Delta = rep(sol_df$dwidth, map_int(sol_df$pattern, 
                                                           length)))

sol_df_plot <- sol_df_plot %>%
  mutate(Roll.Width = as.character(Width))

# scale the widths so that the max is 1.2 and the min is .3
width_range <- range(sol_df_plot$Pattern.Count)
sol_df_plot <- sol_df_plot %>%
  mutate(line.width = 1.2 - (1.2 - .3) *
           (width_range[2] - Pattern.Count) / 
           (width_range[2] - width_range[1]),
         Scen = "Orig")
sol_df_plot <- sol_df_plot %>%
  mutate(line.width = ifelse(line.width == "NaN", 1.2, line.width))

sol_df_plot_orig <- sol_df_plot
avg_width_orig <- avg_width
num_pattern_orig <- num_pattern

# Read the json file with the parent trims
jf_parent <- file.path(cspsol_dat, "solution_parent.json")

# The trim program sometimes inserts a row with just a comma in row 5
# Delete it
jf_fix <- scan(jf_parent, what = "", sep = "\n")
if (jf_fix[5] == ",") jf_fix <- jf_fix[-5]
write(jf_fix, file = jf_parent, sep = "\n")

sol <- fromJSON(jf_parent)

num_pattern <- sol$run$Optimal

sol_df <- data_frame(pattern = sol$run$Solution$pattern, 
                     #pwidth = map_int(sol$run$Solution$pattern, sum),
                     pcount = sol$run$Solution$count)

sol_df$pattern <- map(sol_df$pattern, ~ .x / 16)

sol_df <- sol_df %>%
  mutate(pwidth = map_dbl(pattern, sum)) %>%
  mutate(dwidth = pm_width[trim_mach] - pwidth) %>%
  arrange(desc(pcount)) %>%
  mutate(pnum = row_number())

trim_width <- sol_df %>% 
  summarize(Total.Width = sum(pcount * (pm_width[trim_mach] - dwidth)))

max_width <- num_pattern * pm_width[trim_mach]

delta_width <- max_width - trim_width
delta_width_pct <- delta_width / max_width
avg_width <- with(sol_df, weighted.mean(pwidth, pcount))

rolls <- sol$run$SolutionWidth
rolls <- rolls %>%
  mutate(width = width / 16)

extra_width <- rolls %>%
  summarize(extra_width = sum(width * (solution - demand)))
extra_width_pct <- extra_width / max_width

total_trim_pct <- delta_width_pct + extra_width_pct

sol_df_plot <- 
  data.frame(Pattern = rep(sol_df$pnum, map_int(sol_df$pattern,length)),
             Pattern.Count = rep(sol_df$pcount, map_int(sol_df$pattern, 
                                                        length)), 
             Width = unlist(sol_df$pattern),
             Pattern.Delta = rep(sol_df$dwidth, map_int(sol_df$pattern, 
                                                        length)))

sol_df_plot <- sol_df_plot %>%
  mutate(Roll.Width = as.character(Width))

# scale the widths so that the max is 5 and the min is 1
width_range <- range(sol_df_plot$Pattern.Count)
sol_df_plot <- sol_df_plot %>%
  mutate(line.width = 1.2 - (1.2 - .3) *
           (width_range[2] - Pattern.Count) / 
           (width_range[2] - width_range[1]),
         Scen = "Parent")

sol_df_plot_parent <- sol_df_plot
avg_width_parent <- avg_width
num_pattern_parent <- num_pattern

# Plot 1
ggplot(data = sol_df_plot, aes(x = Pattern, y = Width)) +
  geom_col(aes(fill = Roll.Width, width = line.width), 
           color = "black", stat = "identity", alpha = .9) +
  geom_text(aes(label = Roll.Width), position = position_stack(vjust = 0.5),
            size = 4) +
  geom_hline(aes(yintercept = pm_width[trim_mach]), color = "dark green",
             size = 1.3) +
  geom_text(aes(label = Pattern.Delta, y = pm_width[trim_mach] + 10, 
                x = Pattern)) +
  geom_text(aes(label = Pattern.Count, x = Pattern, y = - 10)) +
  geom_text(aes(label = "Pattern\nCount", y = -10, 
                x = max(sol_df_plot$Pattern) + 1)) +
  coord_flip() +
  theme(text = element_text(size = 12), legend.position = "none") +
  ggtitle(paste0("Total Sets = ", num_pattern, 
                 ", Width Efficiency Loss = ", percent(delta_width_pct[[1]]),
                 ", Unique Patterns = ", max(sol_df_plot$Pattern))) +
  scale_x_discrete(labels = NULL, name = "Pattern Count") +
  scale_fill_manual(values = rev(col_vector))


# Plot both
ggplot(data = bind_rows(sol_df_plot_orig, sol_df_plot_parent), 
       aes(x = Pattern, y = Width)) +
  geom_col(aes(fill = Roll.Width, width = line.width), 
           color = "black", stat = "identity", alpha = .9) +
  geom_text(aes(label = Width), position = position_stack(vjust = 0.5),
            size = 4) +
  geom_hline(aes(yintercept = pm_width[trim_mach]), color = "dark green",
             size = 1.3) +
  geom_text(aes(label = Pattern.Delta, y = pm_width[trim_mach] + 20, 
                x = Pattern, color = ifelse(Pattern.Delta < 0, "0", "1"))) +
  scale_color_manual(values = c("red", "black")) +
  geom_text(aes(label = Pattern.Count, x = Pattern, y = - 10)) +
  #geom_text(aes(label = "Pattern\nCount", y = -10, 
  #              x = max(sol_df_plot$Pattern) + 1)) +
  coord_flip() +
  theme(text = element_text(size = 12), legend.position = "none") +
  ggtitle(paste0("Run Nbr: ", run_number, "  Original Sets = ", 
                 num_pattern_orig, 
                 ", Avg. Width = ", round(avg_width_orig, 2), 
                 " - Parent Sets = ", num_pattern_parent,
                 ", Avg. Width = ", round(avg_width_parent, 2))) +
  scale_x_discrete(labels = NULL, name = "Pattern Count") +
  scale_fill_manual(values = rev(col_vector)) +
  facet_grid(. ~ Scen)


#--------- sub parent widths from opt_detail in existing json trim
sol_df_plot_parent <- 
  left_join(sol_df_plot, 
            filter(vmi_opt, 
                   Fac == "W6" & Grade == "AKPG" & Caliper == "24"), 
            by = c("Width" = "Prod.Width")) %>%
  select(Pattern, Pattern.Count, Width, Pattern.Delta, line.width,
         Parent.Width)

sol_df_plot_parent <- sol_df_plot_parent %>%
  mutate(Parent.Width = ifelse(is.na(Parent.Width), Width, Parent.Width))
  
sol_df_plot_parent <- sol_df_plot_parent %>%
  group_by(Pattern) %>%
  mutate(Pattern.Parent.Width = sum(Parent.Width),
         Pattern.Parent.Delta = pm_width - sum(Parent.Width)) %>%
  ungroup() 

delta_parent_width_pct <- with(sol_df_plot_parent,
                               sum(Pattern.Count * Pattern.Parent.Delta)/
                                 sum(Pattern.Count * Pattern.Parent.Width))

sol_df_plot_parent <- sol_df_plot_parent %>%
  select(Pattern, Pattern.Count, Width = Parent.Width, 
         Pattern.Delta = Pattern.Parent.Delta,
         Pattern.Width = Pattern.Parent.Width, line.width) %>%
  mutate(Roll.Width = as.character(Width),
         Scen = "Post")

ggplot(data = bind_rows(sol_df_plot, sol_df_plot_parent), 
       aes(x = Pattern, y = Width)) +
  geom_col(aes(fill = Roll.Width, width = line.width), 
           color = "black", stat = "identity", alpha = .9) +
  geom_text(aes(label = Width), position = position_stack(vjust = 0.5),
            size = 4) +
  geom_hline(aes(yintercept = pm_width), color = "dark green",
             size = 1.3) +
  geom_text(aes(label = Pattern.Delta, y = pm_width + 20, 
                x = Pattern, color = ifelse(Pattern.Delta < 0, "0", "1"))) +
  scale_color_manual(values = c("red", "black")) +
  geom_text(aes(label = Pattern.Count, x = Pattern, y = - 10)) +
  geom_text(aes(label = "Pattern\nCount", y = -10, 
                x = max(sol_df_plot$Pattern) + 1)) +
  coord_flip() +
  theme(text = element_text(size = 12), legend.position = "none") +
  ggtitle(paste0("Total Sets = ", num_pattern, 
                 ", Width Efficiency Loss = ", 
                 percent(delta_parent_width_pct[[1]]),
                 ", Unique Patterns = ", max(sol_df_plot$Pattern))) +
  scale_x_discrete(labels = NULL) +
  scale_fill_manual(values = rev(col_vector)) +
  facet_grid(. ~ Scen)



#--------- Read trim results from jpg into Excel file via OCR web site
# Fix the row headings and copy into clipboard
trim_res <- paste.table()
trim_res <- trim_res %>%
  select(Pattern = PTN, Pattern.Count = SETS, starts_with("Rolls")) %>%
  arrange(desc(Pattern.Count)) %>%
  mutate(Pattern = seq(1:nrow(.)))

trim_res <- trim_res %>%
  gather(Rolls, Value, 3:ncol(.)) %>%
  filter(Value != "")

trim_res <- trim_res %>%
  separate(Value, c("Roll.Count", "Roll.Width"), sep = "x") %>%
  mutate(Roll.Count = str_replace(Roll.Count, "l", "1"),
         Width = as.numeric(Roll.Width))

trim_res <- trim_res %>%
  mutate(Roll.Count = as.numeric(Roll.Count))

# Repeat each pattern over the roll count
trim_res <- trim_res[rep(row.names(trim_res), trim_res$Roll.Count), 1:6]

trim_res <- trim_res %>% 
  select(-Rolls, -Roll.Count)

# scale the widths so that the max is 5 and the min is 1
width_range <- range(trim_res$Pattern.Count)
trim_res <- trim_res %>%
  mutate(line.width = 1.2 - (1.2 - .3) *
           (width_range[2] - Pattern.Count) / 
           (width_range[2] - width_range[1]),
         Scen = "Pre")

trim_res <- trim_res %>%
  group_by(Pattern) %>%
  mutate(Pattern.Delta = pm_width - sum(Width)) %>%
  arrange(Pattern.Count) %>%
  ungroup() 

num_pattern <- sum(distinct(trim_res, Pattern.Count))
total_width <- trim_res %>%
  distinct(Pattern, Pattern.Count, Pattern.Delta) %>%
  summarize(sum(Pattern.Count * Pattern.Delta))
delta_parent_width_pct <- total_width[[1]] / (num_pattern * pm_width)
trim_res$Scen = "Orig"

ggplot(data = trim_res, 
       aes(x = Pattern, y = Width)) +
  geom_col(aes(fill = Roll.Width, width = line.width), 
           color = "black", stat = "identity", alpha = .9) +
  geom_text(aes(label = Width), position = position_stack(vjust = 0.5),
            size = 4) +
  geom_hline(aes(yintercept = pm_width), color = "dark green",
             size = 1.3) +
  geom_text(aes(label = Pattern.Delta, y = pm_width + 10, 
                x = Pattern)) +
  geom_text(aes(label = Pattern.Count, x = Pattern, y = - 10)) +
  geom_text(aes(label = "Pattern\nCount", y = -10, 
                x = max(trim_res$Pattern) + 1)) +
  coord_flip() +
  theme(text = element_text(size = 12), legend.position = "none") +
  ggtitle(paste0("Total Sets = ", num_pattern, 
                 ", Width Efficiency Loss = ", 
                 percent(delta_parent_width_pct[[1]]),
                 ", Unique Patterns = ", max(trim_res$Pattern))) +
  scale_x_discrete(labels = NULL) +
  scale_fill_manual(values = rev(col_vector)) +
  facet_grid(. ~ Scen)

#--------- Re-trim with parent rolls

# TO DO:  Add Code

rm(trim_scenario, trim_detail, trim_year, trim_output, file_list,
   trim_df, trim_df_pre, trim_fac, trim_grade, trim_cal, tpi,
   dmd_weeks, sol_df_plot, fp, trim_data)


#---------- Copy cspsol input data from excel spreadsheet.  
# Data must be in the following form:
# PM WIdth in inches
# Width in inches | Qty Desired
trim_data <- paste.table()

names(trim_data) <- c("Width", "Qty.Req")
trim_data$Width.mm <- ceiling(trim_data$Width * 25.4)
write.table(pm_width_mm,file.path(tfp, "trim.txt"),
            row.names = FALSE, col.names = FALSE)
write.table(select(trim_data, Width.mm, Qty.Req), 
            file.path(tfp, "trim.txt"), append = TRUE,
            row.names = FALSE, col.names = FALSE)

# Run cspsol with the above as the data file

