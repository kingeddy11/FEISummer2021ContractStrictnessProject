library(tidyverse)
library(dplyr)
library(lubridate)
library(readxl)
library(DescTools)
library(ggrepel)
library(viridis)
library(tibble)
library(writexl)

#setting work directory
setwd("C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping")
#reading and running all individual credit agreement scripts
source("BroadcomIncScrape.R")
source("TimeIncScrape.R")
source("BerkeleyLightsScrape.R")
source("BlueGreenVacationsScrape.R")
source("BREPropertiesScrape.R")
source("CamdenPropertyTrustScrape.R")
source("CityOfficeReitScrape.R")
source("AcclaimEntertainmentScrape.R")
source("AT&T Latin America.R")
source("AvizaTechnologyScrape.R")
source("BellsouthCorpScrape.R")
source("BenefitfocusScrape.R")
source("AccentiaScrape.R")
source("AdmaBiologicsScrape.R")
source("AmericanRenalAssociatesScrape.R")
source("ArquleScrape.R")
source("AxogenScrape.R")
source("AearoCorpScrape.R")
source("BungeLTDScrape.R")
source("CoreMoldingTechnologiesScrape.R")
source("HandyandHarmanScrape.R")
source("HartmarxCorp.R")
source("AddvantageTechnologiesGroupScrape.R")
source("AlteraHealthcareCorp.R")
source("AutowebcomScrape.R")
source("BlockbusterScrape.R")
source("DoverMotorsportsScrape.R")

#list.files("C:/Users/edwar/OneDrive/Documents/FEISummer2021/CreditContractStrictnessProfBattaSum2021/Scraping", full.names = TRUE) %>% map(source)

#joining EBITDA Dataset
EBITDA<-full_join(Broadcom_EBITDA_data_new, Time_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Blue_Green_Vaca_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Bre_Properties_EBITDA_data_new)
EBITDA<-full_join(EBITDA, City_Office_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Acclaim_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Aviza_Technology_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Bellsouth_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Accentia_EBITDA_data_new)
EBITDA<-full_join(EBITDA, American_Renal_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Aearo_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Handy_and_Harman_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Hartmarx_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Addvantage_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Blockbuster_EBITDA_data_new)
EBITDA<-full_join(EBITDA, Dover_EBITDA_data_new)

#joining Net Income Dataset
Net_Income<-full_join(Broadcom_Net_Income_data_new, Time_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Blue_Green_Vaca_Net_Income_data_new)
Net_Income<-full_join(Net_Income, City_Office_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Aviza_Technology_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Bellsouth_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Accentia_Net_Income_data_new)
Net_Income<-full_join(Net_Income, American_Renal_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Aearo_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Hartmarx_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Blockbuster_Net_Income_data_new)
Net_Income<-full_join(Net_Income, Dover_Net_Income_data_new)

#joining Total Debt Dataset
Total_Debt<-full_join(Broadcom_Total_Debt_data_new, Time_Total_Debt_data_new)
Total_Debt<-full_join(Total_Debt, Bre_Properties_Total_Debt_data_new)
Total_Debt<-full_join(Total_Debt, American_Renal_Total_Debt_data_new)
Total_Debt<-full_join(Total_Debt, Aearo_Total_Debt_data_new)
Total_Debt<-full_join(Total_Debt, Hartmarx_Total_Debt_data_new)
Total_Debt<-full_join(Total_Debt, Addvantage_Total_Debt_data_new)

#joining Interest Charges Dataset
Interest_Charges<-Broadcom_Interest_Charges_data_new

#joining Fixed Charges Dataset
Fixed_Charges<-Acclaim_Fixed_Charges_data_new

#joining Cash Equivalents Dataset
Cash_Equivalents<-Time_Cash_Equivalents_def_data_new
Cash_Equivalents<-full_join(Cash_Equivalents, Adma_Biologics_Cash_Equivalents_def_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, American_Renal_Cash_Equivalents_def_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, Arqule_Cash_Equivalents_def_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, Axogen_Cash_Equivalents_def_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, Aearo_Cash_Equivalents_def_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, Handy_and_Harman_Cash_Equivalents_def_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, Hartmarx_Cash_Equivalents_def_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, Blockbuster_Cash_Equivalents_data_new)
Cash_Equivalents<-full_join(Cash_Equivalents, Dover_Cash_Equivalents_data_new)

#joining Indebtedness Definition Dataset
Indebtedness_def<-full_join(Broadcom_Indebtedness_def_data_new, Time_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Blue_Green_Vaca_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Bre_Properties_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Berkeley_Lights_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Camden_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, City_Office_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Acclaim_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, ATT_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Aviza_Technology_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Bellsouth_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Benefit_Focus_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Accentia_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Adma_Biologics_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, American_Renal_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Arqule_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Axogen_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Aearo_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Bunge_LTD_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Core_Molding_Tech_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Handy_and_Harman_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Hartmarx_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Altera_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Autoweb_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Blockbuster_Indebtedness_def_data_new)
Indebtedness_def<-full_join(Indebtedness_def, Dover_Indebtedness_def_data_new)

#joining Subordinated Debt Dataset
Subordinated_Debt<-Berkeley_Lights_Subordinated_Debt_data_new

#joining Permitted Transfer Dataset
Permitted_Transfer<-Berkeley_Lights_Permitted_Transfer_data_new

#joining Permitted Investment Dataset
Permitted_Investment<-Berkeley_Lights_Permitted_Investment_data_new

#joining Net Operating Income Dataset
Net_Operating_Income<-Camden_Net_Operating_Income_data_new

#joining Valuation Dataset
Valuation<-Camden_Valuation_data_new

#joining Permitted Liens Dataset
Permitted_Liens<-Addvantage_Permitted_Liens_data_new
Permitted_Liens<-full_join(Permitted_Liens, Altera_Permitted_Liens_data_new)

#joining Lien Definition Dataset
Lien_def<-Addvantage_Lien_def_data_new
Lien_def<-full_join(Lien_def, Altera_Lien_def_data_new)

#joining Book Net Worth Dataset
Book_Net_Worth_def<-Altera_Book_Net_Worth_def_data_new

#joining Tangible Net Worth Dataset
Tangible_Net_Worth<-Altera_Tangible_Net_Worth_data_new
Tangible_Net_Worth<-full_join(Tangible_Net_Worth, Autoweb_Tangible_Net_Worth_data_new)

#joining Other Definitions Dataset
Other_def<-full_join(Broadcom_Other_def_data_new, Time_Other_def_data_new)
Other_def<-full_join(Other_def, City_Office_Other_def_data_new)
Other_def<-full_join(Other_def, Aviza_Technology_Other_def_data_new)
Other_def<-full_join(Other_def, Bellsouth_Other_def_data_new)
Other_def<-full_join(Other_def, Benefit_Focus_Other_def_data_new)
Other_def<-full_join(Other_def, Axogen_Other_def_data_new)
Other_def<-full_join(Other_def, Aearo_Other_def_data_new)
Other_def<-full_join(Other_def, Core_Molding_Tech_Other_def_data_new)
Other_def<-full_join(Other_def, Handy_and_Harman_Other_def_data_new)
Other_def<-full_join(Other_def, Hartmarx_Other_def_data_new)
Other_def<-full_join(Other_def, Altera_Other_def_data_new)
Other_def<-full_join(Other_def, Blockbuster_Other_def_data_new)
Other_def<-full_join(Other_def, Dover_Other_def_data_new)

#joining Financial Covenants Dataset
Financial_Covenant<-Time_Financial_Covenant_data_new
Financial_Covenant$Threshold<-as.character(Financial_Covenant$Threshold)
Financial_Covenant<-full_join(Financial_Covenant, Blue_Green_Vaca_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Bre_Properties_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Camden_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, City_Office_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Aviza_Technology_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Bellsouth_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Benefit_Focus_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Accentia_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Adma_Biologics_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, American_Renal_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Aearo_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Core_Molding_Tech_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Handy_and_Harman_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Hartmarx_Financial_Covenant_data_new)
Financial_Covenant<-full_join(Financial_Covenant, Addvantage_Financial_Covenant_data_new)

#joining Dynamic Thresholds Dataset
Dynamic_Thresholds<-ATT_Dynamic_Thresholds_data_new
Dynamic_Thresholds<-full_join(Dynamic_Thresholds, Altera_Dynamic_Thresholds_data_new)
Dynamic_Thresholds<-full_join(Dynamic_Thresholds, Autoweb_Dynamic_Thresholds_data_new)
Dynamic_Thresholds<-full_join(Dynamic_Thresholds, Blockbuster_Dynamic_Thresholds_data_new)

#joining Investments Dataset
Investments<-Time_Investments_data_new
Investments<-full_join(Investments, Berkeley_Lights_Investments_data_new)
Investments<-full_join(Investments, City_Office_Investments_data_new)
Investments<-full_join(Investments, Adma_Biologics_Investments_data_new)
Investments<-full_join(Investments, American_Renal_Investments_data_new)
Investments<-full_join(Investments, Arqule_Investments_data_new)
Investments<-full_join(Investments, Axogen_Investments_data_new)
Investments<-full_join(Investments, Aearo_Investments_data_new)
Investments<-full_join(Investments, Core_Molding_Tech_Investments_data_new)
Investments<-full_join(Investments, Handy_and_Harman_Investments_data_new)
Investments<-full_join(Investments, Hartmarx_Investments_data_new)
Investments<-full_join(Investments, Blockbuster_Investments_data_new)

#joining Capital Expenditures and Acquisitions Dataset
Capex_and_acqui<-Blue_Green_Vaca_Capex_and_acqui_data_new
Capex_and_acqui<-full_join(Capex_and_acqui, Berkeley_Lights_Capex_and_acqui_data_new)
Capex_and_acqui<-full_join(Capex_and_acqui, City_Office_Capex_and_acqui_data_new)
Capex_and_acqui<-full_join(Capex_and_acqui, Acclaim_Capex_and_acqui_data_new)
Capex_and_acqui<-full_join(Capex_and_acqui, Aviza_Technology_Capex_and_acqui_data_new)
Capex_and_acqui<-full_join(Capex_and_acqui, Benefit_Focus_Capex_and_acqui_data_new)
Capex_and_acqui<-full_join(Capex_and_acqui, Core_Molding_Tech_Capex_and_acqui_data_new)
Capex_and_acqui<-full_join(Capex_and_acqui, Handy_and_Harman_Capex_and_acqui_data_new)
Capex_and_acqui<-full_join(Capex_and_acqui, Hartmarx_Capex_and_acqui_data_new)

#joining Indebtedness Dataset
Indebtedness<-Time_Indebtedness_data_new
Indebtedness<-full_join(Indebtedness, Berkeley_Lights_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Camden_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, City_Office_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Acclaim_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Aviza_Technology_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Benefit_Focus_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Aearo_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Core_Molding_Tech_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Handy_and_Harman_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Hartmarx_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Addvantage_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Altera_Indebtedness_data_new)
Indebtedness<-full_join(Indebtedness, Blockbuster_Indebtedness_data_new)
Indebtedness$Sign<-as.double(Indebtedness$Sign)
Indebtedness<-full_join(Indebtedness, Dover_Indebtedness_data_new)

#joining Liens Dataset
Liens<-full_join(Broadcom_Liens_data_new, Time_Liens_data_new)
Liens<-full_join(Liens, Blue_Green_Vaca_Liens_data_new)
Liens<-full_join(Liens, Berkeley_Lights_Liens_data_new)
Liens<-full_join(Liens, Camden_Liens_data_new)
Liens<-full_join(Liens, City_Office_Liens_data_new)
Liens<-full_join(Liens, Aviza_Technology_Liens_data_new)
Liens$Deduction<-as.character(Liens$Deduction)
Liens<-full_join(Liens, Bellsouth_Liens_data_new)
Liens<-full_join(Liens, Benefit_Focus_Liens_data_new)
Liens<-full_join(Liens, Accentia_Liens_data_new)
Liens<-full_join(Liens, Adma_Biologics_Liens_data_new)
Liens<-full_join(Liens, American_Renal_Liens_data_new)
Liens<-full_join(Liens, Arqule_Liens_data_new)
Liens<-full_join(Liens, Axogen_Liens_data_new)
Liens<-full_join(Liens, Aearo_Liens_data_new)
Liens<-full_join(Liens, Core_Molding_Tech_Liens_data_new)
Liens<-full_join(Liens, Handy_and_Harman_Liens_data_new)
Liens<-full_join(Liens, Hartmarx_Liens_data_new)
Liens<-full_join(Liens, Addvantage_Liens_data_new)
Liens<-full_join(Liens, Altera_Liens_data_new)
Liens<-full_join(Liens, Autoweb_Liens_data_new)
Liens<-full_join(Liens, Blockbuster_Liens_data_new)
Liens<-full_join(Liens, Dover_Liens_data_new)

#joining Disposals Dataset
Disposals<-Time_Disposals_data_new
Disposals<-full_join(Disposals, Berkeley_Lights_Disposals_data_new)
Disposals<-full_join(Disposals, City_Office_Disposals_data_new)
Disposals<-full_join(Disposals, Aviza_Technology_Disposals_data_new)
Disposals<-full_join(Disposals, Benefit_Focus_Disposals_data_new)
Disposals<-full_join(Disposals, Adma_Biologics_Disposals_data_new)
Disposals<-full_join(Disposals, American_Renal_Disposals_data_new)
Disposals<-full_join(Disposals, Arqule_Disposals_data_new)
Disposals<-full_join(Disposals, Axogen_Disposals_data_new)
Disposals<-full_join(Disposals, Aearo_Disposals_data_new)
Disposals<-full_join(Disposals, Core_Molding_Tech_Disposals_data_new)
Disposals<-full_join(Disposals, Handy_and_Harman_Disposals_data_new)
Disposals<-full_join(Disposals, Hartmarx_Disposals_data_new)
Disposals<-full_join(Disposals, Addvantage_Disposals_data_new)
Disposals<-full_join(Disposals, Altera_Disposals_data_new)
Disposals<-full_join(Disposals, Blockbuster_Disposals_data_new)
Disposals<-full_join(Disposals, Dover_Disposals_data_new)

#joining Restricted Payments Dataset
Restricted_Payments<-Time_Restricted_Payments_data_new
Restricted_Payments<-full_join(Restricted_Payments, Berkeley_Lights_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, Camden_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, City_Office_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, Aviza_Technology_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, Aearo_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, Bunge_LTD_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, Handy_and_Harman_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, Addvantage_Restricted_Payments_data_new)
Restricted_Payments<-full_join(Restricted_Payments, Blockbuster_Restricted_Payments_data_new)
Restricted_Payments$Sign<-as.double(Restricted_Payments$Sign)
Restricted_Payments<-full_join(Restricted_Payments, Dover_Restricted_Payments_data_new)

#joining Negative Covenant Dataset
Negative_Covenants<-ATT_Negative_Covenants_data_new
Negative_Covenants<-full_join(Negative_Covenants, Altera_Negative_Covenants_data_new)

#joining Covenant Components Dataset
Covenant_Components<-Time_Covenant_Components_data_new
Covenant_Components<-full_join(Covenant_Components, Dover_Covenant_Components_data_new)

#joining Exceptions Dataset
Exceptions<-Broadcom_Exceptions_data_new

#Extra Combined Datasets
#joining Investments and Restricted Payments
Investments_and_Restricted_Payments<-Benefit_Focus_Investments_and_Restricted_Payments_data_new

#Restrictions Lessee
Restrictions_Lessee<-Altera_Restrictions_Lessee_data_new

#Continuity of Operations
Continuity_of_Operations<-Autoweb_Continuity_of_Operations_data_new


#Organizing the Datasets into implemented general structure
#EBITDA
EBITDA<-EBITDA %>% drop_na(Text)

#Net Income
Net_Income<-Net_Income %>% drop_na(Text)

#Total Debt
Total_Debt<-Total_Debt %>% 
  mutate(`General category`= coalesce(`Main category`, `General category`)) %>% 
  select(-c(`Main category`))


#Interest Charges
Interest_Charges<-Interest_Charges %>% 
  rename(`General category`= `Main category`)

#Fixed Charges
Fixed_Charges<-Fixed_Charges

#Cash Equivalents
Cash_Equivalents<-Cash_Equivalents

#Indebtedness Definition
Indebtedness_def<-Indebtedness_def %>% 
  select(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction)

#Subordinated Debt
Subordinated_Debt<-Subordinated_Debt %>% 
  select(IntelligizeID, Item, `Specific Definition`, Text)

#Permitted Transfer
Permitted_Transfer<-Permitted_Transfer

#Permitted Investment
Permitted_Investment<-Permitted_Investment

#Net Operating Income
Net_Operating_Income<-Net_Operating_Income

#Valuation
Valuation<-Valuation

#Permitted Liens
Permitted_Liens<-Permitted_Liens

#Lien Definition
Lien_def<-Lien_def

#Book Net Worth
Book_Net_Worth_def<-Book_Net_Worth_def

#Tangible Net Worth
Tangible_Net_Worth<-Tangible_Net_Worth

#Other Definitions
Other_def<-Other_def

#Financial Covenants
Financial_Covenant<-Financial_Covenant %>% drop_na(Text)

#Dynamic Thresholds
Dynamic_Thresholds<-Dynamic_Thresholds %>% 
  mutate(`Threshold 1` = coalesce(`Threshold 1`, `Categorization or Calculation`)) %>% 
  mutate(Date = coalesce(Date, Dates)) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Date, Category, `Threshold 1`, `Threshold 2`))

#Investments
Investments<-Investments %>% 
  mutate(Deduction=coalesce(Deductions, Deduction)) %>% 
  select(-c(Deductions))

#Capital Expenditures and Acquisitions
Capex_and_acqui<-Capex_and_acqui

#Indebtedness
Indebtedness<-Indebtedness %>% 
  mutate(Deduction = coalesce(Deduction, Provision)) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Liens
Liens<-Liens %>% 
  mutate(Deduction = coalesce(Deduction, Deductions)) %>% 
  mutate(Category = coalesce(Category, `Main category`)) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Disposals
Disposals<-Disposals %>% 
  mutate(Deduction=coalesce(Deduction, Deduction_1)) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction, Deduction_2))

#Restricted Payments
Restricted_Payments<-Restricted_Payments %>% 
  mutate(Deduction=coalesce(Deduction, Provision)) %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Negative Covenants
Negative_Covenants<-Negative_Covenants %>% 
  select(c(IntelligizeID, Item, `Specific Definition`, Text, Sign, Category, Deduction))

#Covenant Components
Covenant_Components<-Covenant_Components

#Exceptions
Consolidated_Interest_Coverage_Ratio<-Exceptions

#Investments and Restricted Payments
Investments_and_Restricted_Payments<-Investments_and_Restricted_Payments

#Continuity of Operations
Continuity_of_Operations<-Continuity_of_Operations

#Restrictions Lessee
Restrictions_Lessee<-Restrictions_Lessee

#Keeping all necessary datasets
tokeep<-c('EBITDA', 'Net_Income', 'Total_Debt', 'Interest_Charges', 'Fixed_Charges',
          'Cash_Equivalents', 'Indebtedness_def', 'Subordinated_Debt', 'Permitted_Transfer',
          'Permitted_Investment', 'Net_Operating_Income', 'Valuation', 'Permitted_Liens',
          'Lien_def', 'Book_Net_Worth_def', 'Tangible_Net_Worth', 'Other_def',
          'Financial_Covenant', 'Dynamic_Thresholds', 'Investments', 'Capex_and_acqui',
          'Indebtedness', 'Liens', 'Disposals', 'Restricted_Payments', 'Negative_Covenants',
          'Covenant_Components', 'Consolidated_Interest_Coverage_Ratio', 'Investments_and_Restricted_Payments',
          'Continuity_of_Operations', 'Restrictions_Lessee')

toremove <- setdiff(ls(), tokeep)
rm(list = c(toremove, 'toremove'))

#Converting all datasets to csv files
list_df <- mget(ls())

lapply(1:length(list_df), function(i) write.csv(list_df[[i]], 
                                                file = paste0(names(list_df[i]), ".csv"),
                                                row.names = FALSE))
