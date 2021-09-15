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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Karasinski Axogen .xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "NI", "Total Debt", "Interest Charges", "Fixed Charge", "Cash Equivalents", "Indebtedness Def", "Other", 
                 "Financial Cov", "Dynamic Threshold", "Investments", "M&A",
                 "Indebtedness$", "Liens",  "Asset Disposition", "Distribution", "Covenant Components")
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
    Cash_Equivalents_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
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
Cash_Equivalents_def_data_ori<-Cash_Equivalents_def_data_ori %>% 
  rename(Text = 1) %>% drop_na(Text)

#Disposals
Disposals_data_ori<-Disposals_data_ori %>%
  rename(Text = 1) %>% drop_na(Text)

#Investments
Investments_data_ori<-Investments_data_ori %>% rename(Text=1) %>%  
  drop_na(Text)

#Liens
Liens_data_ori<-Liens_data_ori %>% rename(Text=1) %>%  
  drop_na(Text)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Axogen_Cash_Equivalents_def_data_new<-add_new_columns(Cash_Equivalents_def_data_ori, "Cash Equivalents", "Cash Equivalents Definition")
Axogen_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Axogen_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", c(NA, "subordinated indebtedness", "subordinated indebtedness", "subordinated indebtedness"))
Axogen_Investments_data_new<-add_new_columns(Investments_data_ori, "Investments", "Investments Definition")
Axogen_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Axogen_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")


#selecting columns we want to keep
#Cash Equivalents
Axogen_Cash_Equivalents_def_data_new<-Axogen_Cash_Equivalents_def_data_new %>%
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Indebtedness Definition
Axogen_Indebtedness_def_data_new<-Axogen_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Other Definitions
Axogen_Other_def_data_new<-Axogen_Other_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Investments
Axogen_Investments_data_new<-Axogen_Investments_data_new %>%
  select(-c(Definition))

#Liens
Axogen_Liens_data_new<-Axogen_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#Disposals
Axogen_Disposals_data_new<-Axogen_Disposals_data_new %>% 
  select(-c(Definition))











