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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Benefitfocus.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Interest Charges", "Fixed Charge", "Cash Equivalents", "Indebtedness Def", "Other", 
                 "Financial Cov", "Dynamic Threshold", "Investments & Distributions", "M&A",
                 "Indebtedness$", "(Liens)",  "Asset Disposition", "Payment", "Covenant Components")
category_list
#grabbing IntelligizeID number
Identification<-as_tibble(read_excel(path,  col_names = FALSE, sheet="Identification", range = "A1:E5"))
#Identification_list<-lapply(split(Identification, 1:nrow(Identification)), as.list)
for(i in seq_along(Identification)){
  for(j in seq_along(Identification[[i]])){
    if(str_detect(Identification[[i]][[j]], "^\\d{6,8}")==TRUE & is.na(Identification[[i]][[j]])==FALSE){
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
    Investments_and_Restricted_Payments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
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
Cash_Equivalents_def_data_ori<-Cash_Equivalents_def_data_ori %>% rename(Text = 1) %>% 
  rename(Category = 2) %>% drop_na(Text)

#Indebtedness Def
Indebtedness_def_data_ori$Sign<-c(NA,1,1,1,1)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Benefit_Focus_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Benefit_Focus_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", "Other Definitions")
Benefit_Focus_Financial_Covenant_data_new<-add_new_columns(Financial_Covenant_data_ori, "Financial Covenants", "Financial Covenants Definition")
Benefit_Focus_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Benefit_Focus_Capex_and_acqui_data_new<-add_new_columns(Capex_and_acqui_data_ori, "Capital Expenditures and Acquisitions", "Capital Expenditures and Acquisitions Definition")
Benefit_Focus_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Benefit_Focus_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Benefit_Focus_Investments_and_Restricted_Payments_data_new<-add_new_columns(Investments_and_Restricted_Payments_data_ori, 
                                                              c(NA, "Restricted Payments", "Investments"), c(NA, "Restricted Payments Definition", "Investments Definition"))

#selecting columns we want to keep
#Indebtedness Definition
Benefit_Focus_Indebtedness_def_data_new<-Benefit_Focus_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item,`Specific Definition`, Text, Sign, Category))

#Other Definitions
Benefit_Focus_Other_def_data_new<-Benefit_Focus_Other_def_data_new %>% 
  select(-c(`Specific Definition`)) %>% rename(`Specific Definition`=Defintion) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text))

#Financial Covenants
Benefit_Focus_Financial_Covenant_data_new<-Benefit_Focus_Financial_Covenant_data_new %>% 
  select(-c(Conditions))

#Indebtedness
Benefit_Focus_Indebtedness_data_new<-Benefit_Focus_Indebtedness_data_new %>% 
  select(-c(`Condition 1`, `Condition 2`))

#Capital Expenditures and Acquisitions
Benefit_Focus_Capex_and_acqui_data_new<-Benefit_Focus_Capex_and_acqui_data_new %>% 
  select(-c(`Condition 1`, `Condition 2`))

#Liens
Benefit_Focus_Liens_data_new<-Benefit_Focus_Liens_data_new %>% 
  select(-c(`Condition 1`, `Condition 2`))

#Disposals
Benefit_Focus_Disposals_data_new<-Benefit_Focus_Disposals_data_new %>% 
  select(-c(`Condition 1`, `Condition 2`))

#Investments and Restricted Payments
Benefit_Focus_Investments_and_Restricted_Payments_data_new<-Benefit_Focus_Investments_and_Restricted_Payments_data_new %>% 
  select(c(IntelligizeID, Item,`Specific Definition`, Text, Category, Deduction))




