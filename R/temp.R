calc_params <- function(x) {
  y <- list()
  is.recursive(x)
  #if (use_gamma & x$distrib == "Gamma") {
  if (use_gamma & x["distrib"] == "Gamma") {
    y$shape <- (x["dmd_mean_lt"]) ^ 2 / (x["dmd_sd_lt"]) ^ 2
    y$scale <- (x["dmd_sd_lt"]) ^ 2 / x["dmd_mean_lt"]
    y$rop <- qgamma(service_level,
                    y$shape,
                    scale = y$scale,
                    lower.tail = TRUE)
    y$sfty_stock <- y$rop - x["dmd_mean_lt"]
    y$act_distrib <- "Gamma"
    
  } else {
    y$shape <- 0
    y$scale <- 0
    y$sfty_stock <- x["dmd_sd_lt"] * sfty_factor
    y$rop <- y$sfty_stock + x["dmd_mean_lt"]
    y$act_distrib <- "Normal"
  }
  return(as.data.frame(y))
}
