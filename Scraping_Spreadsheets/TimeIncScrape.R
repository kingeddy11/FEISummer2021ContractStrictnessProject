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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/TimeInc.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Interest Charges", "Fixed Charges", "Cash Equivalents",
                 "Indebtedness Def", "Other", "Financial Cov", "Dynamic Threshold", "Investments", 
                 "Capital Expenditures and Acquisitions", "indebtedness$", "Liens",  "Disposition", 
                 "Payment", "Covenant Components")
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
Cash_Equivalents_def_data_ori<-Cash_Equivalents_def_data_ori %>% rename(Text = 1) %>% 
  rename(Category = 2) %>% drop_na(Text)

#Covenant Components
Covenant_Components_data_ori<-Covenant_Components_data_ori %>% select(-c(Item))

#Disposals
Disposals_data_ori<-Disposals_data_ori %>% 
  rename(Text = `SECTION 7.04. Dispositions. The Borrower will not, and will not permit any of its Restricted Subsidiaries to, consummate any Disposition, except:`)

#Financial Covenant
Financial_Covenant_data_ori<-Financial_Covenant_data_ori %>% 
  rename(Text = `SECTION 7.08. Financial Covenant`)

#Indebtedness 
Indebtedness_data_ori<-Indebtedness_data_ori %>% 
  rename(Text = `SECTION 7.02. Incurrence of Indebtedness and Issuance of Disqualified Stock and Preferred Stock.`)

#Indebtedness Definition
Indebtedness_def_data_ori<-Indebtedness_def_data_ori %>% 
  rename(Sign = 3)

#Investments Definition
Investments_data_ori<-Investments_data_ori %>% rename(Text = 1)

#Liens
Liens_data_ori<-Liens_data_ori %>% rename(Text = 1)

#Other Definitions
Other_def_data_ori<-Other_def_data_ori %>% rename(Text = 1) %>% 
  rename(Sign=`Sign (1,-1)`) %>%
  drop_na(Text)

#Restricted Payments
Restricted_Payments_data_ori<-Restricted_Payments_data_ori %>% 
  rename(Text = `SECTION 7.05. Restricted Payments. (a) The Borrower will not, and will not permit any of its Restricted Subsidiaries to, directly or indirectly:`)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Time_EBITDA_data_new<-add_new_columns(EBITDA_data_ori, "EBITDA", "Consolidated EBITDA")
Time_Net_Income_data_new<-add_new_columns(Net_Income_data_ori, "Net Income", "Consolidated Net Income" )
Time_Total_Debt_data_new<-add_new_columns(Total_Debt_data_ori, "Total Debt", "Consolidated Total Debt")
Time_Cash_Equivalents_def_data_new<-add_new_columns(Cash_Equivalents_def_data_ori, "Cash Equivalents", "Cash Equivalents Definition")
Time_Indebtedness_def_data_new<-add_new_columns(Indebtedness_def_data_ori, "Indebtedness Definition", "Indebtedness Definition")
Time_Other_def_data_new<-add_new_columns(Other_def_data_ori, "Other Definitions", c("Disqualified Stock", "Eligible Cash", "Disposition Deficiency"))
Time_Financial_Covenant_data_new<-add_new_columns(Financial_Covenant_data_ori, "Financial Covenants", "Financial Covenants Definition")
Time_Investments_data_new<-add_new_columns(Investments_data_ori, "Investments", "Investments Definition")
Time_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Time_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Time_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Time_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")
Time_Covenant_Components_data_new<-add_new_columns(Covenant_Components_data_ori, "Covenant Components", "Consolidated Secured Net Leverage Ratio")

#selecting columns we want to keep
#EBITDA
Time_EBITDA_data_new<-Time_EBITDA_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))

#Net Income
Time_Net_Income_data_new<-Time_Net_Income_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, `General category`, General_category_memo))

#Total Debt
Time_Total_Debt_data_new<-Time_Total_Debt_data_new %>% 
  select(c(IntelligizeID, Item,`Specific Definition`, Text, `Main category`))

#Cash Equivalents
Time_Cash_Equivalents_def_data_new<-Time_Cash_Equivalents_def_data_new

#Indebtedness Definition
Time_Indebtedness_def_data_new <- Time_Indebtedness_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category))

#Other Definitions
Time_Other_def_data_new<-Time_Other_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text , Sign, Category))

#Financial Covenants
Time_Financial_Covenant_data_new<-Time_Financial_Covenant_data_new

#Investments
Time_Investments_data_new<-Time_Investments_data_new %>% select(-c(Definition))

#Indebtedness
Time_Indebtedness_data_new<-Time_Indebtedness_data_new %>% select(-c(Definition))

#Liens
Time_Liens_data_new<-Time_Liens_data_new

#Disposals
Time_Disposals_data_new<-Time_Disposals_data_new %>% select(-c(Definition))

#Restricted Payments
Time_Restricted_Payments_data_new<-Time_Restricted_Payments_data_new %>% select(-c(Definition))

#Covenant Components
Time_Covenant_Components_data_new<-Time_Covenant_Components_data_new


