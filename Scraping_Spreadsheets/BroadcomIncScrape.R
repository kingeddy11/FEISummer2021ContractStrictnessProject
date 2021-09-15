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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/BroadcomInc.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Interest Charges", "Fixed Charges", "Cash Equivalents", "Indebtedness Def", "Other", 
"Financial Cov", "Dynamic Threshold", "Investments", "Capital Expenditures and Acquisitions",
"Indebtedness$", "Liens",  "Disposals", "Payment", "Covenant Components")
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
  else if(!(tab_names[[i]] %in% category_list & tab_names!="Identification")) {
    Exceptions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
}
#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Interest Charges
Interest_Charges_data_ori<-Interest_Charges_data_ori %>% rename(Sign = sign)

#Other Definitions
Other_def_data_ori<-Other_def_data_ori %>% rename(Text=`Does not have excess cash flow or permitted definitions`) %>%
  rename(Sign=`Sign (1,-1)`) %>%
  drop_na(Text)

#Liens
Liens_data_ori<-Liens_data_ori %>% rename(Text=`Section 7.01 Liens`)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Broadcom_EBITDA_data_new<-add_new_columns(EBITDA_data_ori, "EBITDA", "Consolidated EBITDA")
Broadcom_Net_Income_data_new<-add_new_columns(Net_Income_data_ori, "Net Income", "Consolidated Net Income" )
Broadcom_Total_Debt_data_new<-add_new_columns(Total_Debt_data_ori, "Total Debt", "Consolidated Total Debt")
Broadcom_Interest_Charges_data_new<-add_new_columns(Interest_Charges_data_ori, "Interest Charges", "Consolidated Interest Charges")
Broadcom_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness", "Indebtedness Definition")
Broadcom_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", c("Transactions", "Maximum Secured Debt Limit"))
Broadcom_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Broadcom_Exceptions_data_new<-add_new_columns(Exceptions_data_ori, "Consolidated Interest Coverage Ratio", "Consolidated Interest Coverage Ratio Definition")

#selecting columns we want to keep
#EBITDA
Broadcom_EBITDA_data_new<-Broadcom_EBITDA_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))

#Net Income
Broadcom_Net_Income_data_new<-Broadcom_Net_Income_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))

#Total Debt
Broadcom_Total_Debt_data_new<-Broadcom_Total_Debt_data_new %>% 
  select(c(IntelligizeID, Item,`Specific Definition`, Text, `Main category`))

#Interest Charges
Broadcom_Interest_Charges_data_new<-Broadcom_Interest_Charges_data_new %>% 
  select(c(IntelligizeID, Item,`Specific Definition`, Text, `Main category`))

#Indebtedness Definition
Broadcom_Indebtedness_def_data_new <- Broadcom_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Other Definitions
Broadcom_Other_def_data_new<-Broadcom_Other_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text , Sign, Category))

#Liens
Broadcom_Liens_data_new<-Broadcom_Liens_data_new %>%
  select(-c(Definition))

#Exceptions
Broadcom_Exceptions_data_new<-Broadcom_Exceptions_data_new
  

