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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Guilloteau BERKELEY LIGHTS.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Interest Charges", "Fixed Charges", "Cash Equivalents",
                 "Indebtedness Def", "Subordinated Debt", "Permitted Transfer", "Permitted Investment", "Investments", 
                 "Capital Expenditures and Acquisitions", "Indebtedness$", "Liens",  "Disposition", 
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
    Subordinated_Debt_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[9]])==TRUE){
    Permitted_Transfer_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[10]])==TRUE){
    Permitted_Investment_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
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
  #else if(!(tab_names[[i]] %in% category_list & tab_names!="Identification")) {
    #Exceptions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
}
#list<-lapply(excel_sheets(path), read_excel, path = path)

#Combining Investments and Inventory and Equipment into Investments
Unmerged_Investments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[14]]))
Inventory_and_Equipment_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[16]]))
Investments_data_ori<-full_join(Unmerged_Investments_data_ori, Inventory_and_Equipment_data_ori)

rm(Unmerged_Investments_data_ori)
rm(Inventory_and_Equipment_data_ori)
#Combining M&A and Affiliate Transactions
MA_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[11]]))
Affiliate_Transactions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[15]]))
Capex_and_acqui_data_ori<-full_join(MA_data_ori, Affiliate_Transactions_data_ori)

rm(MA_data_ori)
rm(Affiliate_Transactions_data_ori)
#adjusting sheets that do not have a "Text" Column
#Capital Expenditures and Acquisitions
Capex_and_acqui_data_ori<-Capex_and_acqui_data_ori %>% rename(Text=1)

#Disposals
Disposals_data_ori<-Disposals_data_ori %>% rename(Text=1)

#Indebtedness
Indebtedness_data_ori<-Indebtedness_data_ori %>% rename(Text=1)

#Indebtedness Definition
Indebtedness_def_data_ori<-Indebtedness_def_data_ori %>% rename(Category=Definition)
#Investments
Investments_data_ori<-Investments_data_ori %>% rename(Text=1)

#Permitted Investment
Permitted_Investment_data_ori<-Permitted_Investment_data_ori %>% rename(Category=Definition)
#Permitted Transfer
Permitted_Transfer_data_ori<-Permitted_Transfer_data_ori %>% rename(Text = 1) %>% 
  rename(Sign = `Sign (1,-1)`) %>% rename(Category=Definition)

#Restricted Payments
Restricted_Payments_data_ori<-Restricted_Payments_data_ori %>% rename(Text=1)

#Subordinated Debt
Subordinated_Debt_data_ori<-Subordinated_Debt_data_ori %>% rename(Text=1) %>% 
  rename(Sign = `Sign (1,-1)`)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Berkeley_Lights_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Berkeley_Lights_Subordinated_Debt_data_new<-add_new_columns(Subordinated_Debt_data_ori, "Subordinated Debt", "Subordinated Debt Definition" )
Berkeley_Lights_Permitted_Transfer_data_new<-add_new_columns(Permitted_Transfer_data_ori, "Permitted Transfer", "Permitted Transfer Definition")
Berkeley_Lights_Permitted_Investment_data_new<-add_new_columns(Permitted_Investment_data_ori, "Permitted Investment", "Permitted Investment Definition")
Berkeley_Lights_Investments_data_new<-add_new_columns(Investments_data_ori, "Investments", "Investments Definition")
Berkeley_Lights_Capex_and_acqui_data_new<-add_new_columns(Capex_and_acqui_data_ori, "Capital Expenditures and Acquisitions", "Capital Expenditures and Acquisitions Definition")
Berkeley_Lights_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Berkeley_Lights_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Berkeley_Lights_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Berkeley_Lights_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")


#selecting columns we want to keep
#Indebtedness Definition
Berkeley_Lights_Indebtedness_def_data_new <- Berkeley_Lights_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))
Berkeley_Lights_Indebtedness_def_data_new$IntelligizeID<-as.character(Berkeley_Lights_Indebtedness_def_data_new$IntelligizeID)

#Subordinated Debt
Berkeley_Lights_Subordinated_Debt_data_new<-Berkeley_Lights_Subordinated_Debt_data_new
Berkeley_Lights_Subordinated_Debt_data_new$IntelligizeID<-as.character(Berkeley_Lights_Subordinated_Debt_data_new$IntelligizeID)

#Permitted Transfer
Berkeley_Lights_Permitted_Transfer_data_new<-Berkeley_Lights_Permitted_Transfer_data_new %>% select(-c(exclusion))
Berkeley_Lights_Permitted_Transfer_data_new$IntelligizeID<-as.character(Berkeley_Lights_Permitted_Transfer_data_new$IntelligizeID)

#Permitted Investment
Berkeley_Lights_Permitted_Investment_data_new<-Berkeley_Lights_Permitted_Investment_data_new %>%
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))
Berkeley_Lights_Permitted_Investment_data_new$IntelligizeID<-as.character(Berkeley_Lights_Permitted_Investment_data_new$IntelligizeID)

#Investments
Berkeley_Lights_Investments_data_new<-Berkeley_Lights_Investments_data_new %>% select(-c(Definition))
Berkeley_Lights_Investments_data_new$IntelligizeID<-as.character(Berkeley_Lights_Investments_data_new$IntelligizeID)

#Capital Expenditures and Acquisitions
Berkeley_Lights_Capex_and_acqui_data_new<-Berkeley_Lights_Capex_and_acqui_data_new %>% select(-c(Definition))
Berkeley_Lights_Capex_and_acqui_data_new$IntelligizeID<-as.character(Berkeley_Lights_Capex_and_acqui_data_new$IntelligizeID)

#Indebtedness
Berkeley_Lights_Indebtedness_data_new<-Berkeley_Lights_Indebtedness_data_new %>% select(-c(definition))
Berkeley_Lights_Indebtedness_data_new$IntelligizeID<-as.character(Berkeley_Lights_Indebtedness_data_new$IntelligizeID)

#Liens
Berkeley_Lights_Liens_data_new<-Berkeley_Lights_Liens_data_new %>%
  select(c(IntelligizeID, Item, `Specific Definition`, Text, `Main category`))
Berkeley_Lights_Liens_data_new$IntelligizeID<-as.character(Berkeley_Lights_Liens_data_new$IntelligizeID)

#Disposals
Berkeley_Lights_Disposals_data_new<-Berkeley_Lights_Disposals_data_new %>% select(-c(Definition))
Berkeley_Lights_Disposals_data_new$IntelligizeID<-as.character(Berkeley_Lights_Disposals_data_new$IntelligizeID)

#Restricted Payments
Berkeley_Lights_Restricted_Payments_data_new<-Berkeley_Lights_Restricted_Payments_data_new %>% select(-c(Definition))
Berkeley_Lights_Restricted_Payments_data_new$IntelligizeID<-as.character(Berkeley_Lights_Restricted_Payments_data_new$IntelligizeID)

