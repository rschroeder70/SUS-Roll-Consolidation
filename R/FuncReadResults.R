#------------------------------------------------------------------------------
#
# Define a function to load the optimal solution details for the scenario
# under consideration

read_results <- function(results_type, mYear, projscenario) {
  # Results Types can be one of:
  #  opt_detail (details at the sub-roll level for the otpimal solution)
  #  sim_parents (simulation output for all parent possibilities)
  #  sol_summary (summary stats for each possible solution - cost & tons)
  #  dmd_prof_prod (Product Demand Profile Details)
  if (results_type == "opt_detail") {
    fstr <- ".Sim Results."
  } else if (results_type == "sim_parents") {
    fstr <- ".Parent Results."
  } else if (results_type == "sol_summary") {
    fstr <- ".Solutions Summary."
  } else if (results_type == "dmd_prof_prod") {
    fstr <- ".Demand Profile-Product."
  } else if (results_type == "all_details") {
    fstr <- ".All Solution Details."
  } else {
    stop("incorrect file type specified")
  }
  
  fpoutput <- file.path(proj_root,"Scenarios", mYear, projscenario)
  file_list <- list.files(fpoutput, fstr)
  file_list <- file.path(fpoutput, file_list)
  
  # use map_df from the purrr library
  #df <- adply(file_list, 1, read_csv, .id = NULL)
  df <- map_df(file_list, read_csv)
  
  # Need to get rid of the annoying UTF-8 Byte Order Marker (BOM)
  # "<U+FEFF>" that shows up in column 1
  #names(df)[1] <- str_sub(format(names(df)[1], trim = TRUE), 9, 11)
  
  # Operations specific to the opt_detail & sim_parents files
  if (results_type %in% c("opt_detail", "sim_parents")) {
    df <- df %>%
      mutate(Caliper = str_sub(CalDW, 1, 2),
             Wind = str_sub(CalDW, 7, 7),
             Dia = str_sub(CalDW, 4, 5))
  }
  
  # Operations specific to the opt_detail file
  if (results_type == "opt_detail") {
    # Replace the NaN (in the Fill.Rate column) with zeros
    df <- replace(df, is.na(df), 0)
    
    # Reduce the number of columns 
    df <- df %>%
      select(SolParentCount, Fac, Grade, Caliper, Dia, Wind, CalDW, 
             Parent.Width, Parent.Nbr.Subs, Parent.DMD,
             Parent.SS.DOH, Parent.DMD.Count, Sim.Parent.Avg.OH, 
             Sim.Parent.Qty.Fill.Rate, Parent.dclass, 
             Prod.Width, Prod.DMD.Tons, Prod.Trim.Width, Prod.Trim.Fact,
             Prod.Trim.Tons, Prod.Gross.Tons, Prod.DMD.Count,
             Parent.Distrib, Parent.OTL, Parent.DBR)
    
    df <- df %>%
      mutate(Prod.Trim.Width = round(Prod.Trim.Width, 4))
    
    df %>%
      summarize(Nbr.Prods = n(),
                Prod.DMD.Tons = sum(Prod.DMD.Tons),
                Prod.Gross.Tons = sum(Prod.Gross.Tons),
                Nbr.Parents = n_distinct(Fac, Grade, CalDW, Parent.Width))
  }  
  return(df)
}
#------------------------------------------------------------------------------
