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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Button BLOCKBUSTER.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Cash Equivalents", "Lien Definition", "Book Net Worth",
                 "Indebtedness Def", "Other", "Tangible Net Worth", "Thresholds", "Investments", 
                 "Capital Expenditures and Acquisitions", "Indebtedness$", "Liens",  "Limits on Transactions", 
                 "Restricted Payments", "Covenant Components")
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
    Cash_Equivalents_data_ori<-as_tibble(read_excel(path, col_names=FALSE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[5]])==TRUE){
    Lien_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[6]])==TRUE){
    Book_Net_Worth_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[7]])==TRUE){
    Indebtedness_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[8]])==TRUE){
    Other_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[9]])==TRUE){
    Tangible_Net_Worth_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
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
  #else if(str_detect(tab_names[[i]], category_list[[15]])==TRUE){
    #Disposals_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
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

#separately combining Asset Sales and Sale Leaseback Sheet
#Negative Cov
Asset_Sales_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[10]]))
Sale_Leaseback_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[11]]))
Asset_Sales_data_ori<-Asset_Sales_data_ori %>% 
  rename(Text=`SECTION 6.05. Asset Sales.`)
Sale_Leaseback_data_ori<-Sale_Leaseback_data_ori %>% 
  rename(Text=`SECTION 6.10. Restrictive Agreements.`)

Disposals_data_ori<-full_join(Asset_Sales_data_ori, Sale_Leaseback_data_ori)
rm(Asset_Sales_data_ori)
rm(Sale_Leaseback_data_ori)

#Appending sheets after thresholds to other definitions
Permitted_Acquisitions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[15]]))
Permitted_Encumbrances_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[16]]))
Permitted_Investments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[17]]))
Permitted_Store_Sales_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[18]]))
Permitted_Subordinated_Indebted_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[19]]))

Other_def_data_ori<-full_join(Permitted_Acquisitions_data_ori, Permitted_Encumbrances_data_ori)
Other_def_data_ori<-full_join(Other_def_data_ori, Permitted_Investments_data_ori)
Other_def_data_ori<-full_join(Other_def_data_ori, Permitted_Store_Sales_data_ori)
Other_def_data_ori<-full_join(Other_def_data_ori, Permitted_Subordinated_Indebted_data_ori)

#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Cash Equivalents
Cash_Equivalents_data_ori<-Cash_Equivalents_data_ori %>% 
  rename(Text=1) %>% rename(Category=2) %>% drop_na(Text)

#Dynamic Thresholds
Dynamic_Thresholds_data_ori<-Dynamic_Thresholds_data_ori %>% 
  select(-c(Item))

#Indebtedness
Indebtedness_data_ori<-Indebtedness_data_ori %>% 
  rename(Text=1)

#Investments
Investments_data_ori<-Investments_data_ori %>% 
  rename(Text=1) %>% drop_na(Text)

#Liens
Liens_data_ori<-Liens_data_ori %>% 
  rename(Text=1)

#Net Income
Net_Income_data_ori$Sign<-c(1)

#Other Definitions
Other_def_data_ori<-Other_def_data_ori %>% 
  rename(Text=1) %>% drop_na(Text)

#Restricted Payments
Restricted_Payments_data_ori<-Restricted_Payments_data_ori %>% 
  rename(Text=`SECTION 6.08. Restricted Payments; Certain Payments of Indebtedness.`)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Blockbuster_EBITDA_data_new<-add_new_columns(EBITDA_data_ori, "EBITDA", "Consolidated EBITDA")
Blockbuster_Net_Income_data_new<-add_new_columns(Net_Income_data_ori, "Net Income", "Consolidated Net Income" )
Blockbuster_Total_Debt_data_new<-add_new_columns(Total_Debt_data_ori, "Total Debt", "Consolidated Total Debt")
Blockbuster_Cash_Equivalents_data_new<-add_new_columns(Cash_Equivalents_data_ori, "Cash Equivalents", "Cash Equivalents Definition")
Blockbuster_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Blockbuster_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", "Other Definitions")
Blockbuster_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Blockbuster_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Blockbuster_Investments_data_new<-add_new_columns(Investments_data_ori, "Investments", "Investments Definition")
Blockbuster_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Blockbuster_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")

#Adding IntelligizeID, Item, and Definition to Thresholds
Blockbuster_Dynamic_Thresholds_data_new<-Dynamic_Thresholds_data_ori %>% 
  add_column(.before="Text", IntelligizeID=Intelligizeid) %>% 
  add_column(.after="IntelligizeID", Item="Dynamic Thresholds") %>% 
  add_column(.after="Item", `Specific Definition`="Leverage Ratio")

#selecting columns we want to keep
#Cash Equivalents
Blockbuster_Cash_Equivalents_data_new<-Blockbuster_Cash_Equivalents_data_new
  
#Disposals
Blockbuster_Disposals_data_new<-Blockbuster_Disposals_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction, Deduction_1, Deduction_2))

#Dynamic Thresholds
Blockbuster_Dynamic_Thresholds_data_new<-Blockbuster_Dynamic_Thresholds_data_new

#EBITDA
Blockbuster_EBITDA_data_new<-Blockbuster_EBITDA_data_new %>% 
  select(IntelligizeID,Item,`Specific Definition`,Text,Sign,`General category`,General_category_memo)
Blockbuster_EBITDA_data_new$Sign<-as.double(Blockbuster_EBITDA_data_new$Sign)

#Indebtedness
Blockbuster_Indebtedness_data_new<-Blockbuster_Indebtedness_data_new %>% rename(Category=Catagory) %>% 
  select(-c(Definition))

#Indebtedness Definition
Blockbuster_Indebtedness_def_data_new<-Blockbuster_Indebtedness_def_data_new %>% 
  select(IntelligizeID,Item,`Specific Definition`,Text,Sign,Category)
Blockbuster_Indebtedness_def_data_new$Sign<-as.double(Blockbuster_Indebtedness_def_data_new$Sign)

#Investments
Blockbuster_Investments_data_new<-Blockbuster_Investments_data_new %>% 
  select(-c(Definition))

#Liens
Blockbuster_Liens_data_new<-Blockbuster_Liens_data_new %>% 
  select(-c(Definition))

#Net Income
Blockbuster_Net_Income_data_new<-Blockbuster_Net_Income_data_new %>% 
  select(IntelligizeID,Item,`Specific Definition`,Text,Sign,`General category`,General_category_memo)
Blockbuster_Net_Income_data_new$Sign<-as.double(Blockbuster_Net_Income_data_new$Sign)

#Other Definitions
Blockbuster_Other_def_data_new<-Blockbuster_Other_def_data_new %>% rename(Deduction=Deductions) %>% 
  select(-c(`Specific Definition`))
Blockbuster_Other_def_data_new<-Blockbuster_Other_def_data_new %>% 
  rename(`Specific Definition`=Definition) %>% 
  select(c(IntelligizeID,Item,`Specific Definition`, Text, Sign, Category, Deduction))

#Restricted Payments
Blockbuster_Restricted_Payments_data_new<-Blockbuster_Restricted_Payments_data_new %>% 
  select(-c(Definition))

