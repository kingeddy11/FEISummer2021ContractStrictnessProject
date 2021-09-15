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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/GUILLOTEAU CITY OFFICE REIT, INC..xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Interest Charges", "Fixed Charges", "Cash Equivalents",
                 "Indebtedness Def", "Other", "Financial Cov", "Dynamic Threshold", "Investments", 
                 "Capital Expenditures and Acquisitions", "Indebtedness$", "Liens",  "Asset Sales", 
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
    Interest_Charges_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[5]])==TRUE){
    Fixed_Charges_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
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
  else if(str_detect(tab_names[[i]], category_list[[14]])==TRUE | str_detect(tab_names[[i]], "Collateral")==TRUE){
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
  #else if(!(tab_names[[i]] %in% category_list & tab_names!="Identification")) {
  #Exceptions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
}

#separately combining M&A Dataset and Transfer Dataset to Capex and Acqui dataset
MA_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[10]]))
Affiliate_Transactions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[13]]))

Capex_and_acqui_data_ori<-full_join(MA_data_ori, Affiliate_Transactions_data_ori)
rm(MA_data_ori)
rm(Affiliate_Transactions_data_ori)
#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Capital Expenditures and Acquisitions
Capex_and_acqui_data_ori<-Capex_and_acqui_data_ori %>% rename(Text = 1)

#Disposals
Disposals_data_ori<-Disposals_data_ori %>% rename(Text = 1)

#Financial Covenant
Financial_Covenant_data_ori<-Financial_Covenant_data_ori %>% 
  rename(Text = 1) %>% drop_na(Text)

#Indebtedness
Indebtedness_data_ori<-Indebtedness_data_ori %>% rename(Text = 1)

#Investments
Investments_data_ori<-Investments_data_ori %>% rename(Text = 1)

#Liens
Liens_data_ori<-Liens_data_ori %>% rename(Text = 1) %>% drop_na(Text)

#Restricted Payments
Restricted_Payments_data_ori<-Restricted_Payments_data_ori %>% rename(Text = 1)
#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
City_Office_EBITDA_data_new<-add_new_columns(EBITDA_data_ori, "EBITDA", "Consolidated EBITDA")
City_Office_Net_Income_data_new<-add_new_columns(Net_Income_data_ori, "Net Income", "Consolidated Net Income" )
City_Office_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
City_Office_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", c("Pool NOI"))
City_Office_Financial_Covenant_data_new<-add_new_columns(Financial_Covenant_data_ori, "Financial Covenants", "Financial Covenants Definition")
City_Office_Investments_data_new<-add_new_columns(Investments_data_ori, "Investments", "Investments Definition")
City_Office_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
City_Office_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
City_Office_Capex_and_acqui_data_new<-add_new_columns(Capex_and_acqui_data_ori, "Capital Expenditures and Acquisitions", "Capital Expenditures and Acquisitions Definition")
City_Office_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
City_Office_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")

#selecting columns we want to keep
#EBITDA
City_Office_EBITDA_data_new<-City_Office_EBITDA_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))
City_Office_EBITDA_data_new$Sign<-as.double(City_Office_EBITDA_data_new$Sign)

#Net Income
City_Office_Net_Income_data_new<-City_Office_Net_Income_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))

#Indebtedness Definition
City_Office_Indebtedness_def_data_new <- City_Office_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Other Definitions
City_Office_Other_def_data_new<-City_Office_Other_def_data_new %>% select(-c(Definition))

#Financial Covenants
City_Office_Financial_Covenant_data_new<-City_Office_Financial_Covenant_data_new

#Investments
City_Office_Investments_data_new<-City_Office_Investments_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Capital Expenditures and Acquisitions
City_Office_Capex_and_acqui_data_new<-City_Office_Capex_and_acqui_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Indebtedness
City_Office_Indebtedness_data_new<-City_Office_Indebtedness_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Liens
City_Office_Liens_data_new<-City_Office_Liens_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Disposals
City_Office_Disposals_data_new<-City_Office_Disposals_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Restricted Payments
City_Office_Restricted_Payments_data_new<-City_Office_Restricted_Payments_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))
