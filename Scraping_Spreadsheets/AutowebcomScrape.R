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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Button AUTOWEB COM INC.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Permitted Lien", "Lien Definition", "Tangible Net Worth",
                 "Indebtedness Def", "Other", "Financial Covenant$", "Threshold", "Investments", 
                 "Capital Expenditures and Acquisitions", "Indebtedness$", "Liens",  "Sale of Assets", 
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
    Tangible_Net_Worth_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
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
  #else if(str_detect(tab_names[[i]], category_list[[14]])==TRUE | str_detect(tab_names[[i]], "Collateral")==TRUE){
    #Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
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
#Liens
Indebtedness_Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[7]]))
Indebtedness_Liens_data_ori<-Indebtedness_Liens_data_ori %>% rename(Text = 1)
Permitted_Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[10]]))

Liens_data_ori<-full_join(Indebtedness_Liens_data_ori, Permitted_Liens_data_ori)
rm(Indebtedness_Liens_data_ori)
rm(Permitted_Liens_data_ori)

#Continuity of Operations
Continuity_of_Operations_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[8]]))

#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column
#Continuity of Operations
Continuity_of_Operations_data_ori<-Continuity_of_Operations_data_ori %>% 
  rename(Text=1)

#Tangible Net Worth
Tangible_Net_Worth_data_ori<-Tangible_Net_Worth_data_ori %>% 
  rename(Text=1) %>% drop_na(Text)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Autoweb_Continuity_of_Operations_data_new<-add_new_columns(Continuity_of_Operations_data_ori, "Continuity of Operations", "Continuity of Operations Definition")
Autoweb_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Autoweb_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Autoweb_Tangible_Net_Worth_data_new<-add_new_columns(Tangible_Net_Worth_data_ori, "Tangible Net Worth", "Tangible Net Worth Definition")

#Dynamic Thresholds
Autoweb_Dynamic_Thresholds_data_new<-Dynamic_Thresholds_data_ori %>% 
  add_column(.before="Dates", IntelligizeID=Intelligizeid) %>% 
  add_column(.after="IntelligizeID", Item="Dynamic Thresholds") %>% 
  add_column(.after="Item", `Specific Definition`="Dynamic Thresholds Definition")

#selecting columns we want to keep
#Continuity of Operations
Autoweb_Continuity_of_Operations_data_new<-Autoweb_Continuity_of_Operations_data_new %>% 
  select(-c(`Specific Definition`)) %>% rename(`Specific Definition`=Definition) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Dynamic Thresholds
Autoweb_Dynamic_Thresholds_data_new<-Autoweb_Dynamic_Thresholds_data_new

#Indebtedness
Autoweb_Indebtedness_def_data_new <- Autoweb_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text))

#Liens
Autoweb_Liens_data_new<-Autoweb_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Tangible Net Worth
Autoweb_Tangible_Net_Worth_data_new<-Autoweb_Tangible_Net_Worth_data_new %>% 
  select(-c(Definition))

