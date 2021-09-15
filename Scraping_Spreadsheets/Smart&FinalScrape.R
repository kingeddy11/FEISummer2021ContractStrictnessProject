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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Smart&Final.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Indebtedness Def", "Other", 
                 "Financial Cov", "Dynamic Threshold", "Investments", "Capital Expenditures and Acquisitions",
                 "Indebtedness$", "Liens",  "Disposals", "Payment")
category_list
#grabbing IntelligizeID number
Identification<-as_tibble(read_excel(path,  col_names = FALSE, sheet="Identification", range = "A1:E5"))
#Identification_list<-lapply(split(Identification, 1:nrow(Identification)), as.list)
for(i in seq_along(Identification)){
  for(j in seq_along(Identification[[i]])){
    if(str_detect(Identification[[i]][[j]], "^\\d{7,8}")==TRUE & is.na(Identification[[i]][[j]])==FALSE){
      IntelligizeID<-Identification[[i]][[j]]
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
    Indebtedness_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[5]])==TRUE){
    Other_def_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[6]])==TRUE){
    Financial_Covenant_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[7]])==TRUE){
    Dynamic_Thresholds_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[8]])==TRUE){
    Investments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[9]])==TRUE){
    Capex_and_acqui_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[10]])==TRUE){
    Indebtedness_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[11]])==TRUE){
    Liens_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[12]])==TRUE){
    Disposals_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(str_detect(tab_names[[i]], category_list[[13]])==TRUE){
    Restricted_Payments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  else if(!(tab_names[[i]] %in% category_list) & tab_names!="Identification") {
    Consolidated_Interest_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
}

#Creating new version for each aspect with correct format for datasets
#EBITDA
EBITDA_data_new<-EBITDA_data_ori %>% add_column(.before="Text", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="EBITDA") %>% 
  rename(`Specific Definition`=Definition ) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))

#Net Income
Net_Income_data_new<-Net_Income_data_ori %>% add_column(.before="Text", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Net Income") %>% 
  rename(`Specific Definition`=Definition ) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))

#Total Debt
Debt_data_new<-Debt_data_ori %>% add_column(.before="Text", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Total Debt") %>%
  add_column(.after="Item", `Specific Definition`="Consolidated Total Debt") %>% 
  select(c(IntelligizeID, Item,`Specific Definition`, Text, `Main category`))

#Indebtedness Definition
Indebtedness_def_data_new<-Indebtedness_def_data_ori %>% add_column(.before="Text", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Indebtedness") %>%
  add_column(.before = "Text", `Specific Definition`="Indebtedness Definition") %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Other Definitions
Other_def_data_new<-Other_def_data_ori %>% rename(Text=Term) %>%
  rename(`Specific Definition`=Category) %>% 
  drop_na(Text) %>%
  add_column(.before="Text", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Other Definitions") %>%
  select(c(IntelligizeID, Item, `Specific Definition`, Text))

#Financial Covenants
Fin_Cov_data_new<-Fin_Cov_data_ori %>%
  add_column(.before="Text", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Financial Covenants") %>%
  select(c(IntelligizeID, Item, Text, Category, Threshold, Calculation))

#Investments
Investments_data_new<-Investments_data_ori %>% rename(`Specific Definition`= Definition) %>% 
  add_column(.before="Section 6.04 Investments, Loans, and Advances", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Financial Covenants") %>%
  select(c(IntelligizeID, Item, `Specific Definition`, `Section 6.04 Investments, Loans, and Advances`, Category, Deductions))

#Indebtedness Definition
Indebtedness_data_new<-Indebtedness_data_ori %>% rename(`Specific Definition`= Definition) %>%
  add_column(.before="Section 6.01 Indebtedness", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Indebtedness") %>%
  select(c(IntelligizeID, Item, `Specific Definition`, `Section 6.01 Indebtedness`, Category, Deductions))

#Liens Definition
Liens_data_new<-Liens_data_ori %>% rename(`Specific Definition`= Definition) %>%
  add_column(.before="Section 6.02 Liens", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Liens") %>%
  select(c(IntelligizeID, Item, `Specific Definition`, `Section 6.02 Liens`, Category, Deductions))

#Disposals Definition
Disposals_new<-Disposals_ori %>% rename(`Specific Definition`= Definition) %>%
  add_column(.before="Section 6.05 Mergers, Consolidations, Sales of Assets and Acquisitions", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Disposals") %>%
  select(c(IntelligizeID, Item, `Specific Definition`, `Section 6.05 Mergers, Consolidations, Sales of Assets and Acquisitions`, Category, Deductions))

#Restricted Payments Definition
Restricted_Payments_data_new<-Restricted_Payments_data_ori %>% rename(`Specific Definition`= Definition) %>%
  add_column(.before="Section 6.06 Restricted Payments", IntelligizeID=8686328) %>% 
  add_column(.after="IntelligizeID", Item="Restricted Payments") %>%
  select(c(IntelligizeID, Item, `Specific Definition`, `Section 6.06 Restricted Payments`, Category, Deductions))


