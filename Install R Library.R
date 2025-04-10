load.lib<-c("devtools","Rcpp","xml2","XML","plyr","dplyr","tidyverse","vroom","lifecycle","lifecycle","lubridate","stringr","openxlsx","readr","rvest",
            "filesstrings","httr","xlsReadWrite")

install.lib <- load.lib[!load.lib %in% installed.packages()]

for(lib in install.lib) 
    install.packages(lib, dependencies=TRUE, method = "wininet")
                        
sapply(load.lib,require,character=TRUE)
