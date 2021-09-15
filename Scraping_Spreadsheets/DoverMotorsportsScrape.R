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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Button DOVER MOTORSPORTS INC.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Cash Equivalents", "Lien Definition", "Book Net Worth",
                 "Indebtedness Def", "Other", "Tangible Net Worth", "Thresholds", "Investments", 
                 "Limitations on Acquistions", "Limitation on Debt", "Liens",  "Limits on Transactions", 
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
Intelligizeid<-as.character(Intelligizeid)

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
  #else if(str_detect(tab_names[[i]], category_list[[14]])==TRUE | str_detect(tab_names[[i]], "Collateral")==TRUE){
    #Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
  #else if(str_detect(tab_names[[i]], category_list[[15]])==TRUE){
  #Disposals_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
  #else if(str_detect(tab_names[[i]], category_list[[16]])==TRUE){
    #Restricted_Payments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
  else if(str_detect(tab_names[[i]], category_list[[17]])==TRUE){
    Covenant_Components_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  #else if(!(tab_names[[i]] %in% category_list & tab_names!="Identification")) {
  #Exceptions_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
}

#Combining Disposals Dataset
Limitations_on_Fundamental_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[9]]))
Sale_of_Asset_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[10]]))
Sale_and_Leaseback_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[14]]))
Disposals_data_ori<-full_join(Limitations_on_Fundamental_data_ori, Sale_of_Asset_data_ori)
Disposals_data_ori<-full_join(Disposals_data_ori, Sale_and_Leaseback_data_ori)
rm(Limitations_on_Fundamental_data_ori)
rm(Sale_of_Asset_data_ori)
rm(Sale_and_Leaseback_data_ori)

#Combining Restricted Payments
Limitation_on_Distribution_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[12]]))
Transactions_with_Affil_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[13]]))
Restricted_Payments_data_ori<-full_join(Limitation_on_Distribution_data_ori, Transactions_with_Affil_data_ori)
rm(Limitation_on_Distribution_data_ori)
rm(Transactions_with_Affil_data_ori)

#Combining Liens
Limitation_on_Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[8]]))
Limitation_on_Negative_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[15]]))
Liens_data_ori<-full_join(Limitation_on_Liens_data_ori, Limitation_on_Negative_data_ori)
rm(Limitation_on_Liens_data_ori)
rm(Limitation_on_Negative_data_ori)

#Appending sheets after covenant components to other definitions
Permitted_Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[17]]))
Permitted_Investment_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[18]]))
Permitted_Acquisition_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[19]]))

Other_def_data_ori<-full_join(Permitted_Liens_data_ori, Permitted_Investment_data_ori)
Other_def_data_ori<-full_join(Other_def_data_ori, Permitted_Acquisition_data_ori)

#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Cash Equivalents
Cash_Equivalents_data_ori<-Cash_Equivalents_data_ori %>% 
  rename(Text=1) %>% rename(Category=2) %>% drop_na(Text)

#Covenant Components
Covenant_Components_data_ori<-Covenant_Components_data_ori %>% 
  rename(Deduction=4) %>% rename(Item1=Item)

#Disposals
Disposals_data_ori<-Disposals_data_ori %>% 
  rename(Text=1)

#Indebtedness
Indebtedness_data_ori<-Indebtedness_data_ori %>% 
  rename(Text=1)

#Indebtedness Definition
Indebtedness_def_data_ori<-Indebtedness_def_data_ori %>% 
  drop_na(Text)

#Liens
Liens_data_ori<-Liens_data_ori %>% 
  rename(Text=1)

#Other Definitions
Other_def_data_ori<-Other_def_data_ori %>% 
  rename(Text=1) %>% drop_na(Text)

#Restricted Payments
Restricted_Payments_data_ori<-Restricted_Payments_data_ori %>% 
  rename(Text=1)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Dover_EBITDA_data_new<-add_new_columns(EBITDA_data_ori, "EBITDA", "Consolidated EBITDA")
Dover_Net_Income_data_new<-add_new_columns(Net_Income_data_ori, "Net Income", "Consolidated Net Income" )
Dover_Cash_Equivalents_data_new<-add_new_columns(Cash_Equivalents_data_ori, "Cash Equivalents", "Cash Equivalents Definition")
Dover_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Dover_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", "Other Definitions")
Dover_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Dover_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Dover_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Dover_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")
Dover_Covenant_Components_data_new<-add_new_columns(Covenant_Components_data_ori, "Covenant Components", "Covenant Components Definition")

#selecting columns we want to keep
#Cash Equivalents
Dover_Cash_Equivalents_data_new<-Dover_Cash_Equivalents_data_new

#Covenant Components
Dover_Covenant_Components_data_new<-Dover_Covenant_Components_data_new %>% select(-c(`Specific Definition`)) %>% 
  rename(`Specific Definition`=Item1) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, `Categorization or Calculation`, Deduction))

#Disposals
Dover_Disposals_data_new<-Dover_Disposals_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#EBITDA
Dover_EBITDA_data_new<-Dover_EBITDA_data_new %>% 
  select(IntelligizeID,Item,`Specific Definition`,Text,Sign,`General category`,General_category_memo)
Dover_EBITDA_data_new$Sign<-as.double(Dover_EBITDA_data_new$Sign)

#Indebtedness
Dover_Indebtedness_data_new<-Dover_Indebtedness_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))
Dover_Indebtedness_data_new$Sign<-as.double(Dover_Indebtedness_data_new$Sign)

#Indebtedness Definition
Dover_Indebtedness_def_data_new<-Dover_Indebtedness_def_data_new %>% 
  select(IntelligizeID,Item,`Specific Definition`,Text,Sign,Category)
Dover_Indebtedness_def_data_new$Sign<-as.double(Dover_Indebtedness_def_data_new$Sign)

#Liens
Dover_Liens_data_new<-Dover_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))
Dover_Liens_data_new$Sign<-as.double(Dover_Liens_data_new$Sign)

#Net Income
Dover_Net_Income_data_new<-Dover_Net_Income_data_new %>% 
  select(IntelligizeID,Item,`Specific Definition`,Text,Sign,`General category`,General_category_memo)
Dover_Net_Income_data_new$Sign<-as.double(Dover_Net_Income_data_new$Sign)

#Other Definitions
Dover_Other_def_data_new<-Dover_Other_def_data_new %>% rename(Deduction=Deductions) %>% 
  select(-c(`Specific Definition`))
Dover_Other_def_data_new<-Dover_Other_def_data_new %>% 
  rename(`Specific Definition`=Definition) %>% 
  select(c(IntelligizeID,Item,`Specific Definition`, Text, Category, Deduction))

#Restricted Payments
Dover_Restricted_Payments_data_new<-Dover_Restricted_Payments_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))
Dover_Restricted_Payments_data_new$Sign<-as.double(Dover_Restricted_Payments_data_new$Sign)

