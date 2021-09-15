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
path <- "C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping/data/Button ADDVANTAGE TECHNOLOGIES GROUP INC.xlsx"
#grabbing all sheet names from excel sheet
tab_names<-excel_sheets(path)
tab_names
#tab categories we want for general data structure
category_list<-c("EBITDA", "Net Income", "Total Debt", "Permitted Lien", "Lien Definition", "Cash Equivalents",
                 "Indebtedness Def", "Other", "Financial Cov", "Dynamic Threshold", "Investments", 
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
  #else if(str_detect(tab_names[[i]], category_list[[11]])==TRUE){
    #Investments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
  else if(str_detect(tab_names[[i]], category_list[[12]])==TRUE){
    Capex_and_acqui_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  }
  #else if(str_detect(tab_names[[i]], category_list[[13]])==TRUE){
    #Indebtedness_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[i]]))
  #}
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

#separately combining M&A Dataset and Transfer Dataset to Capex and Acqui dataset
#Indebtedness
Indebtedness_uncomb_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[7]]))
Funded_Debt_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[17]]))

Indebtedness_data_ori<-full_join(Indebtedness_uncomb_data_ori, Funded_Debt_data_ori)
rm(Indebtedness_uncomb_data_ori)
rm(Funded_Debt_data_ori)

#Investments
Investments_uncomb_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[10]]))
Asset_Investments_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[16]]))

Investments_data_ori<-full_join(Investments_uncomb_data_ori, Asset_Investments_data_ori)
rm(Investments_uncomb_data_ori)
rm(Asset_Investments_data_ori)

#Restricted Payments
Dividends_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[11]]))
Sale_of_Stock_data_ori<-as_tibble(read_excel(path, col_names=TRUE, sheet=tab_names[[12]]))

Restricted_Payments_data_ori<-full_join(Dividends_data_ori, Sale_of_Stock_data_ori)
rm(Dividends_data_ori)
rm(Sale_of_Stock_data_ori)

#list<-lapply(excel_sheets(path), read_excel, path = path)

#adjusting sheets that do not have a "Text" Column

#Disposals
Disposals_data_ori<-Disposals_data_ori %>% rename(Text = 1)

#EBITDA
EBITDA_data_ori<-EBITDA_data_ori %>% rename(`General category`=Inclusions)

#Investments
Investments_data_ori<-Investments_data_ori %>% rename(Text = 1)

#Lien Definition
Lien_def_data_ori<-Lien_def_data_ori %>% rename(Text = 1)

#Liens
Liens_data_ori<-Liens_data_ori %>% rename(Text = 1)

#Restricted Payments
Restricted_Payments_data_ori<-Restricted_Payments_data_ori %>% rename(Text = 1)

#Total Debt
Total_Debt_data_ori<-Total_Debt_data_ori %>% rename(`General category`=Inclusions)

#adding IntelligizeID, Item, Specific Definition
add_new_columns <- function(df, item, definition){
  new_df<-add_column(df, .before="Text", IntelligizeID=Intelligizeid)
  new_df<-add_column(new_df, .after="IntelligizeID", Item=item)
  new_df<-add_column(new_df, .after="Item", `Specific Definition`=definition)
  return(new_df)
}
Addvantage_EBITDA_data_new<-add_new_columns(EBITDA_data_ori, "EBITDA", "Consolidated EBITDA")
Addvantage_Total_Debt_data_new<-add_new_columns(Total_Debt_data_ori, "Total Debt", "Consolidated Total Debt")
Addvantage_Indebtedness_data_new<-add_new_columns(Indebtedness_data_ori, "Indebtedness", "Indebtedness")
Addvantage_Permitted_Liens_data_new<-add_new_columns(Permitted_Liens_data_ori, "Permitted Liens", "Permitted Liens Definition")
Addvantage_Lien_def_data_new<-add_new_columns(Lien_def_data_ori, "Lien Definition", "Lien Definition")
Addvantage_Financial_Covenant_data_new<-add_new_columns(Financial_Covenant_data_ori, "Financial Covenants", "Financial Covenants Definition")
Addvantage_Investments_data_new<-add_new_columns(Investments_data_ori, "Investments", "Investments Definition")
Addvantage_Liens_data_new<-add_new_columns(Liens_data_ori, "Liens", "Liens Definition")
Addvantage_Disposals_data_new<-add_new_columns(Disposals_data_ori, "Disposals", "Disposals Definition")
Addvantage_Restricted_Payments_data_new<-add_new_columns(Restricted_Payments_data_ori, "Restricted Payments", "Restricted Payments Definition")

#selecting columns we want to keep
#EBITDA
Addvantage_EBITDA_data_new<-Addvantage_EBITDA_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, `General category`))

#Total Debt
Addvantage_Total_Debt_data_new<-Addvantage_Total_Debt_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, `General category`))

#Indebtedness
Addvantage_Indebtedness_data_new <- Addvantage_Indebtedness_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Permitted Liens
Addvantage_Permitted_Liens_data_new<-Addvantage_Permitted_Liens_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category))

#Lien Definition
Addvantage_Lien_def_data_new<-Addvantage_Lien_def_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category))

#Financial Covenants
Addvantage_Financial_Covenant_data_new<-Addvantage_Financial_Covenant_data_new

#Investments
Addvantage_Investments_data_new<-Addvantage_Investments_data_new %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Category, Deduction))

#Liens
Addvantage_Liens_data_new<-Addvantage_Liens_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text))

#Disposals
Addvantage_Disposals_data_new<-Addvantage_Disposals_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text))

#Restricted Payments
Addvantage_Restricted_Payments_data_new<-Addvantage_Restricted_Payments_data_new %>% select(c(IntelligizeID, Item, `Specific Definition`, Text))
