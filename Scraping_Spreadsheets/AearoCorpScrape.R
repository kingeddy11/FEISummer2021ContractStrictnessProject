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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Aearo Corp.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Debt", "Interest Charges", "Fixed Charges", "Cash Equivalents",
                 "Indebtedness Def", "Other", "Financial Cov", "Dynamic Threshold", "Investments", 
                 "CapEx", "Indebtedness$", "Liens",  "Asset Sales", 
                 "Dividends", "Covenant Components")
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
  else if(!(tab_names[[i]] %in% category_list) & (tab_names!="Identification")) {
    Exceptions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
}
#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Cash Equivalents
Cash_Equivalents_def_data_ori<-Cash_Equivalents_def_data_ori %>% rename(Text=1) %>% 
  rename(Category=2) 

#EBITDA
EBITDA_data_ori$Sign[c(2:10)]<-1
EBITDA_data_ori$Sign[[11]]<--1


#Indebtedness Definition
Indebtedness_def_data_ori$Sign[c(2:8)]<-1

#Net Income
Net_Income_data_ori$Sign[c(2:3)]<--1

#Total Debt
Total_Debt_data_ori$Sign[[2]]<-1

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Aearo_EBITDA_data_new<-add_new_columns(EBITDA_data_ori, "EBITDA", "Consolidated EBITDA")
Aearo_Net_Income_data_new<-add_new_columns(Net_Income_data_ori, "Net Income", "Consolidated Net Income" )
Aearo_Total_Debt_data_new<-add_new_columns(Total_Debt_data_ori, "Total Debt", "Consolidated Total Debt")
Aearo_Cash_Equivalents_def_data_new<-add_new_columns(Cash_Equivalents_def_data_ori, "Cash Equivalents", "Cash Equivalents Definition")
Aearo_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Aearo_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions","Other Definitions")
Aearo_Financial_Covenant_data_new<-add_new_columns(Financial_Covenant_data_ori, "Financial Covenants", "Financial Covenants Definition")
Aearo_Investments_data_new<-add_new_columns(Investments_data_ori, "Investments", "Investments Definition")
Aearo_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Aearo_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Aearo_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Aearo_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")

#selecting columns we want to keep
#Cash Equivalents
Aearo_Cash_Equivalents_def_data_new<-Aearo_Cash_Equivalents_def_data_new 

#Disposals
Aearo_Disposals_data_new<-Aearo_Disposals_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#EBITDA
Aearo_EBITDA_data_new<-Aearo_EBITDA_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`))
Aearo_EBITDA_data_new$Sign<-as.double(Aearo_EBITDA_data_new$Sign)

#Financial Covenants
Aearo_Financial_Covenant_data_new<-Aearo_Financial_Covenant_data_new %>% 
  select(-c(Conditions))

#Indebtedness
Aearo_Indebtedness_data_new<-Aearo_Indebtedness_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#Indebtedness Definition
Aearo_Indebtedness_def_data_new <- Aearo_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category))
Aearo_Indebtedness_def_data_new$Sign<-as.double(Aearo_Indebtedness_def_data_new$Sign)

#Investments
Aearo_Investments_data_new<-Aearo_Investments_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#Liens
Aearo_Liens_data_new<-Aearo_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#Net Income
Aearo_Net_Income_data_new<-Aearo_Net_Income_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`))
Aearo_Net_Income_data_new$Sign<-as.double(Aearo_Net_Income_data_new$Sign)

#Other Definitions
Aearo_Other_def_data_new<-Aearo_Other_def_data_new %>% select(-c(`Specific Definition`)) %>% 
  rename(`Specific Definition`= Defintion) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text))

#Restricted Payments
Aearo_Restricted_Payments_data_new<-Aearo_Restricted_Payments_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))  

#Total Debt
Aearo_Total_Debt_data_new<-Aearo_Total_Debt_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, `Main category`))

