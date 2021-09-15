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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Button ALTERA HEALTHCARE CORP.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Permitted Lien", "Lien Definition", "Book Net Worth",
                 "Indebtedness Def", "Other", "Tangible Net Worth", "Thresholds", "Investments", 
                 "Capital Expenditures and Acquisitions", "Indebtedness$", "Liens",  "Limits on Transactions", 
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
    Permitted_Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
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
  else if(str_detect(tab_names[[i]], category_list[[15]])==TRUE){
    Disposals_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
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

#separately combining Negative Covenants sheets
#Negative Cov
Negative_Covenant1_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[15]]))
Negative_Covenant2_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[16]]))

Negative_Covenants_data_ori<-full_join(Negative_Covenant1_data_ori, Negative_Covenant2_data_ori)
rm(Negative_Covenant1_data_ori)
rm(Negative_Covenant2_data_ori)

#Creating Restrictions Lessee Dataset
Restrictions_Lessee_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[12]]))

#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Disposals
Disposals_data_ori<-Disposals_data_ori %>% rename(Text = 1)

#Indebtedness Definition
Indebtedness_def_data_ori<-Indebtedness_def_data_ori %>% rename(Text = 1)

#Lien Definition
Lien_def_data_ori<-Lien_def_data_ori %>% rename(Text = 1)

#Liens
Liens_data_ori<-Liens_data_ori %>% rename(Text = 1)
Liens_data_ori$Sign<-c(-1,1,1,1,1,1,1,1,1)

#Negative Covenants
Negative_Covenants_data_ori<-Negative_Covenants_data_ori %>% rename(Text = 1)
Negative_Covenants_data_ori$Sign[[1]]<-NA
Negative_Covenants_data_ori$Sign[c(2:3)]<-c(-1,-1)
Negative_Covenants_data_ori$Sign[c(4:6)]<-c(1,1,1)

#Restrictions Lessee
Restrictions_Lessee_data_ori<-Restrictions_Lessee_data_ori %>% 
  rename(Text = 1) %>% rename(Sign = Sign...3)

#Tangible Net Worth
Tangible_Net_Worth_data_ori<-Tangible_Net_Worth_data_ori %>% drop_na(Text)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Altera_Book_Net_Worth_def_data_new<-add_new_columns(Book_Net_Worth_def_data_ori, "Book Net Worth", "Book Net Worth Definition")
Altera_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Altera_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Altera_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Altera_Permitted_Liens_data_new<-add_new_columns(Permitted_Liens_data_ori, "Permitted Liens", "Permitted Liens Definition")
Altera_Lien_def_data_new<-add_new_columns(Lien_def_data_ori, "Lien Definition", "Lien Definition")
Altera_Negative_Covenants_data_new<-add_new_columns(Negative_Covenants_data_ori, "Negative Covenants", "Negative Covenants Definition")
Altera_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", "Indebtedness")
Altera_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Altera_Restrictions_Lessee_data_new<-add_new_columns(Restrictions_Lessee_data_ori, "Restrictions Lessee", "Restrictions Lessee Definition")
Altera_Tangible_Net_Worth_data_new<-add_new_columns(Tangible_Net_Worth_data_ori, "Tangible Net Worth", "Tangible Net Worth Definition")

#Adding IntelligizeID, Item, and Definition to Thresholds
Altera_Dynamic_Thresholds_data_new<-Dynamic_Thresholds_data_ori %>% 
  add_column(.before="Dates", IntelligizeID=Intelligizeid) %>% 
  add_column(.after="IntelligizeID", Item="Dynamic Thresholds") %>% 
  add_column(.after="Item", `Specific Definition`="Dynamic Thresholds Definition")

#selecting columns we want to keep
#Book Net Worth
Altera_Book_Net_Worth_def_data_new<-Altera_Book_Net_Worth_def_data_new

#Disposals
Altera_Disposals_data_new<-Altera_Disposals_data_new %>% 
  select(-c(Definition))
#Dynamic Thresholds
Altera_Dynamic_Thresholds_data_new<-Altera_Dynamic_Thresholds_data_new

#Indebtedness
Altera_Indebtedness_data_new<-Altera_Indebtedness_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#Indebtedness Definition
Altera_Indebtedness_def_data_new<-Altera_Indebtedness_def_data_new %>% 
  select(-c(Definition))

#Lien Definition
Altera_Lien_def_data_new<-Altera_Lien_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#Liens
Altera_Liens_data_new<-Altera_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Negative Covenants
Altera_Negative_Covenants_data_new<-Altera_Negative_Covenants_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Other Definitions
Altera_Other_def_data_new<-Altera_Other_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text))

#Permitted Liens
Altera_Permitted_Liens_data_new<-Altera_Permitted_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Restrictions Lessee
Altera_Restrictions_Lessee_data_new<-Altera_Restrictions_Lessee_data_new %>% 
  select(-c(`Specific Definition`, Sign...5)) %>% rename(`Specific Definition`=Definition) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Tangible Net Worth
Altera_Tangible_Net_Worth_data_new<-Altera_Tangible_Net_Worth_data_new
