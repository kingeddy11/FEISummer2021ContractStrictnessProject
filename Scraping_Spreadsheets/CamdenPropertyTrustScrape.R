library(tidyverse)
library(dplyr)
library(lubridate)
library(readxl)
library(DescTools)
library(ggrepel)
library(viridis)
library(tibble)

#setwd("C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data")
#reading in general credit agreement dataset
#make sure to change to your own folder path
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/GUILLOTEAU CAMDEN PROPERTY TRUST.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Net Operating Income", "Valuation", "Cash Equivalents",
                 "Indebtedness Def", "Other", "Financial Cov", "Dynamic Threshold", "Investments", 
                 "Capital Expenditures and Acquisitions", "Indebtedness$", "Liens",  "Disposals", 
                 "Distribution", "Covenant Components")
category_list
#grabbing IntelligizeID number
Identification<-as_tibble(read_excel(path,  col_names = FALSE, sheet="Identification", range = "A1:E5"))
#Identification_list<-lapply(split(Identification, 1:nrow(Identification)), as.list)
for(i in seq_along(Identification)){
  for(j in seq_along(Identification[[i]])){
    if(str_detect(Identification[[i]][[j]], "^\\d{7,8}")==TRUE & is.na(Identification[[i]][[j]])==FALSE){
      Intelligizeid<-Identification[[i]][[j]]
    }
  }
}

#Identifying sheet names we want to use and creating separate datasets for each tab
for(i in seq_along(tab_names)){
  if(str_detect(tab_names[[i]], category_list[[1]])==TRUE){
    EBITDA_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[2]])==TRUE){
    Net_Income_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[3]])==TRUE){
    Total_Debt_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[4]])==TRUE){
    Net_Operating_Income_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[5]])==TRUE){
    Valuation_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[6]])==TRUE){
    Cash_Equivalents_def_data_ori<-as_tibble(read_excel(path, col_names=FALSE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[7]])==TRUE){
    Indebtedness_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[8]])==TRUE){
    Other_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[9]])==TRUE){
    Financial_Covenant_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[10]])==TRUE){
    Dynamic_Thresholds_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[11]])==TRUE){
    Investments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[12]])==TRUE){
    Capex_and_acqui_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[13]])==TRUE){
    Indebtedness_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[14]])==TRUE){
    Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[15]])==TRUE){
    Disposals_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[16]])==TRUE){
    Restricted_Payments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[17]])==TRUE){
    Covenant_Components_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(!(tab_names[[i]] %in% category_list) & tab_names!="Identification") {
    Exceptions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
}
#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Restricted Payments
Restricted_Payments_data_ori<-Restricted_Payments_data_ori %>% 
  rename(Text = 1)

#Financial Covenant
Financial_Covenant_data_ori<-Financial_Covenant_data_ori %>% 
  rename(Text = 1) %>% rename(Threshold=restriction)

#Indebtedness
Indebtedness_data_ori<-Indebtedness_data_ori %>% rename(Text=1)

#Liens
Liens_data_ori<-Liens_data_ori %>% rename(Text = 1)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Camden_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Camden_Net_Operating_Income_data_new<-add_new_columns(Net_Operating_Income_data_ori, "Net Operating Income", "Net Operating Income Definition")
Camden_Valuation_data_new<-add_new_columns(Valuation_data_ori, "Valuation", "Valuation Definition")
Camden_Financial_Covenant_data_new<-add_new_columns(Financial_Covenant_data_ori, "Financial Covenants", "Financial Covenants Definition")
Camden_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Camden_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")
Camden_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")

#selecting columns we want to keep
#Indebtedness Definition
Camden_Indebtedness_def_data_new <- Camden_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))
Camden_Indebtedness_def_data_new$IntelligizeID<-as.character(Camden_Indebtedness_def_data_new$IntelligizeID)

#Net Operating Income
Camden_Net_Operating_Income_data_new<-Camden_Net_Operating_Income_data_new %>% select(-c(Definition))
Camden_Net_Operating_Income_data_new$IntelligizeID<-as.character(Camden_Net_Operating_Income_data_new$IntelligizeID)

#Valuation
Camden_Valuation_data_new<-Camden_Valuation_data_new %>% select(-c(Definition))
Camden_Valuation_data_new$IntelligizeID<-as.character(Camden_Valuation_data_new$IntelligizeID)

#Financial Covenants
Camden_Financial_Covenant_data_new<-Camden_Financial_Covenant_data_new
Camden_Financial_Covenant_data_new$IntelligizeID<-as.character(Camden_Financial_Covenant_data_new$IntelligizeID)

#Indebtedness
Camden_Indebtedness_data_new<-Camden_Indebtedness_data_new %>% select(-c(Definition))
Camden_Indebtedness_data_new$IntelligizeID<-as.character(Camden_Indebtedness_data_new$IntelligizeID)

#Liens
Camden_Liens_data_new<-Camden_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))
Camden_Liens_data_new$IntelligizeID<-as.character(Camden_Liens_data_new$IntelligizeID)

#Restricted Payments
Camden_Restricted_Payments_data_new<-Camden_Restricted_Payments_data_new %>% select(-c(Definition))
Camden_Restricted_Payments_data_new$IntelligizeID<-as.character(Camden_Restricted_Payments_data_new$IntelligizeID)

