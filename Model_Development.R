rm(list=ls())
library(readxl)
library(Hmisc)
library(openair)
library(MASS)

##Setting working Directory
setwd("C:\\Users\\mahen\\OneDrive - University of California, Davis\\Desktop\\School\\BAX 401 Fall - Information, Insight and Impact\\Class 4\\HW2\\HW2\\")

## load data
multdata <- read_excel("HW2_MultimediaHW.xlsx",row.names("Months"))

## Column Names
colnames(multdata)

## Summary
summary(multdata)

## Zero Percentages
colSums(multdata==0)/nrow(multdata)*100

## Removing columns
## Column1 : Months are not need for the model as it just gives the time frame and is not related to sales
## Column3 : ADV_Total cannot be included in the model as it is the aggregated sum of all the advertising types
## Column4 : ADV_Offline cannot be included in the model as it is the aggregated sum of all the offline advertising types
## Column9 : Banner has 90.47% zeros which will provide any value to the model
## Column10 : ADV_Online cannot be included in the model as it is the aggregated sum of all the online advertising types
## Column12 : SocialMedia has 100% zeros which will provide any value to the model
multdata <- multdata[-c(1,3,4,9,10,12)]
pairs(multdata, panel = panel.smooth)

## Correlation Plot
corPlot(multdata)
## Search and Re-targetting are highly correlated with 83%
## Search and Portals are highly correlated with with 88%
## Catelogs_NewCust and Catelogs_Winback are moderately correlated with 57%
## Search and Portals are moderatly correlated with sales with 40% and 49% respectively


## Saving data
multdata_saved <- multdata

######################################## Model_Building with square root ####################################

Sales <- multdata$`Sales (units)`
Stm1 <- Lag(Sales,shift=1) #creating the lag of sales value
Stm1[1] <- 0
multdata_adv_modes <- multdata[-1]
multdata_adv_modes_sqrt <- cbind(as.data.frame(apply(multdata_adv_modes,2,sqrt)),Stm1)

##################### Focal Model 
#Running model with intercept - AIC 552.2 ,ADJ R2 - 0.3096
Square_root_lag1 <- step(lm(Sales~.,data=multdata_adv_modes_sqrt),direction = "backward",na.action=na.omit())
summary(Square_root_lag1)
### Best Model of square root with intercept :
### Sales ~ Catalogs_Winback + Catalogs_NewCust + Portals + Stm1

#Running model with intercept - AIC 553.59 ,ADJ R2 - 0.9805
Square_root_lag1_nointercept <- step(lm(Sales~.-1,data=multdata_adv_modes_sqrt),direction = "backward",na.action=na.omit())
summary(Square_root_lag1_nointercept)
### Best Model of square root without intercept :
### Best Model : Sales ~ Catalogs_Winback + Catalogs_NewCust + Search + Retargeting + Portals + Stm1 - 1

########################################## Model_Building with LN ##########################################

## Saving data
multdata <- multdata_saved

Sales <- multdata$`Sales (units)`
Stm1 <- Lag(Sales,shift=1) #creating the lag of sales value
Stm1[1] <- 0
multdata_adv_modes <- multdata[-1]
multdata_adv_modes_LN <- cbind(as.data.frame(apply(multdata_adv_modes+1,2,log)),Stm1)

#Running model with intercept - AIC 539.76 ,ADJ R2 - 0.4968
LN_lag1 <- step(lm(Sales~.,data=multdata_adv_modes_LN),direction = "backward",na.action=na.omit())
summary(LN_lag1)
### Best Model of log with intercept :
### Sales ~ Catalogs_ExistCust + Catalogs_Winback + Catalogs_NewCust + Newsletter + Portals

##################### Current Extension Model 
#Running model with intercept - AIC AIC=539.68 ,ADJ R2 - 0.9862
LN_lag1_nointercept <- step(lm(Sales~.-1,data=multdata_adv_modes_LN),direction = "backward",na.action=na.omit())
summary(LN_lag1_nointercept)
### Best Model of log without intercept : 
### Sales ~ Catalogs_ExistCust + Catalogs_Winback + Catalogs_NewCust + Search + Newsletter + Retargeting + Portals - 1




######################################## Model_Building with square root (With synergy) ####################################

## Saving data
multdata <- multdata_saved

Sales <- multdata$`Sales (units)`
Stm1 <- Lag(Sales,shift=1) #creating the lag of sales value
Stm1[1] <- 0
multdata_adv_modes <- multdata[-1]
colnames(syn)
syn <- cbind(multdata_adv_modes^2, do.call(cbind,combn(colnames(multdata_adv_modes), 2, 
                                             FUN= function(x) list(multdata_adv_modes[x[1]]*multdata_adv_modes[x[2]]))))
colnames(syn)[-(seq_len(ncol(multdata_adv_modes)))] <-  combn(colnames(multdata_adv_modes), 2, 
                                                    FUN = paste, collapse=":")
syn <- syn[9:36]

multdata_adv_modes <- cbind(multdata_adv_modes,syn)
multdata_adv_modes_sqrt <- cbind(as.data.frame(apply(multdata_adv_modes,2,sqrt)),Stm1)


#Running model with intercept - AIC 507.75 ,ADJ R2 - 0.6965
Square_root_lag1 <- step(lm(Sales~.,data=multdata_adv_modes_sqrt),direction = "backward",na.action=na.omit())
summary(Square_root_lag1)
### Best Model of square root with intercept with synergy:
# Sales ~ Catalogs_ExistCust + Catalogs_Winback + Catalogs_NewCust +
# Mailings + Search + Newsletter + Retargeting + Portals +
#   `Catalogs_ExistCust:Catalogs_Winback` + `Catalogs_ExistCust:Catalogs_NewCust` +
#   `Catalogs_ExistCust:Mailings` + `Catalogs_ExistCust:Search` +
#   `Catalogs_ExistCust:Newsletter` + `Catalogs_ExistCust:Retargeting` +
#   `Catalogs_ExistCust:Portals` + `Catalogs_Winback:Catalogs_NewCust` +
#   `Catalogs_Winback:Mailings` + `Catalogs_Winback:Search` +
#   `Catalogs_Winback:Newsletter` + `Catalogs_Winback:Retargeting` +
#   `Catalogs_Winback:Portals` + `Catalogs_NewCust:Mailings` +
#   `Catalogs_NewCust:Newsletter` + `Catalogs_NewCust:Retargeting` +
#   `Catalogs_NewCust:Portals` + `Mailings:Search` + `Mailings:Newsletter` +
#   `Mailings:Retargeting` + `Mailings:Portals` + `Search:Newsletter` +
#   `Search:Portals` + `Newsletter:Retargeting` + `Newsletter:Portals` +
#   Stm1

#Running model with intercept - AIC 518.13 ,ADJ R2 - 0.9929
Square_root_lag1_nointercept <- step(lm(Sales~.-1,data=multdata_adv_modes_sqrt),direction = "backward",na.action=na.omit())
summary(Square_root_lag1_nointercept)
### Best Model of square root without intercept with synergy :
# Sales ~ Catalogs_ExistCust + Catalogs_Winback + Catalogs_NewCust +
#   Newsletter + Portals + `Catalogs_ExistCust:Catalogs_Winback` +
#   `Catalogs_ExistCust:Catalogs_NewCust` + `Catalogs_ExistCust:Search` +
#   `Catalogs_ExistCust:Portals` + `Catalogs_Winback:Mailings` +
#   `Catalogs_Winback:Newsletter` + `Catalogs_Winback:Retargeting` +
#   `Catalogs_Winback:Portals` + `Catalogs_NewCust:Mailings` +
#   `Catalogs_NewCust:Search` + `Catalogs_NewCust:Newsletter` +
#   `Catalogs_NewCust:Retargeting` + `Mailings:Search` + `Mailings:Newsletter` +
#   `Mailings:Retargeting` + `Search:Newsletter` + `Search:Retargeting` +
#   `Search:Portals` + `Newsletter:Retargeting` - 1




########################################## Model_Building with LN ##########################################

## Saving data
multdata <- multdata_saved

Sales <- multdata$`Sales (units)`
Stm1 <- Lag(Sales,shift=1) #creating the lag of sales value
Stm1[1] <- 0
multdata_adv_modes <- multdata[-1]
colnames(syn)
syn <- cbind(multdata_adv_modes^2, do.call(cbind,combn(colnames(multdata_adv_modes), 2, 
                                                       FUN= function(x) list(multdata_adv_modes[x[1]]*multdata_adv_modes[x[2]]))))
colnames(syn)[-(seq_len(ncol(multdata_adv_modes)))] <-  combn(colnames(multdata_adv_modes), 2, 
                                                              FUN = paste, collapse=":")
syn <- syn[9:36]

multdata_adv_modes <- cbind(multdata_adv_modes,syn)
multdata_adv_modes_LN <- cbind(as.data.frame(apply(multdata_adv_modes+1,2,log)),Stm1)


#Running model with intercept - AIC 507.75 ,ADJ R2 - 0.6161
LN_lag1 <- step(lm(Sales~.,data=multdata_adv_modes_LN),direction = "backward",na.action=na.omit())
summary(LN_lag1)
### Best Model of log with intercept and synergy:
# Sales ~ Catalogs_ExistCust + Catalogs_Winback + Catalogs_NewCust +
#   Mailings + Newsletter + Retargeting + Portals + `Catalogs_ExistCust:Catalogs_Winback` +
#   `Catalogs_ExistCust:Catalogs_NewCust` + `Catalogs_ExistCust:Mailings` +
#   `Catalogs_ExistCust:Search` + `Catalogs_ExistCust:Newsletter` +
#   `Catalogs_ExistCust:Retargeting` + `Catalogs_ExistCust:Portals` +
#   `Catalogs_Winback:Catalogs_NewCust` + `Catalogs_Winback:Mailings` +
#   `Catalogs_Winback:Search` + `Catalogs_Winback:Retargeting` +
#   `Catalogs_Winback:Portals` + `Catalogs_NewCust:Mailings` +
#   `Catalogs_NewCust:Search` + `Catalogs_NewCust:Newsletter` +
#   `Catalogs_NewCust:Portals` + `Mailings:Search` + `Mailings:Newsletter` +
#   `Mailings:Retargeting` + `Mailings:Portals` + `Search:Newsletter` +
#   `Search:Retargeting` + `Search:Portals` + `Newsletter:Retargeting` +
#   `Newsletter:Portals` + `Retargeting:Portals` + Stm1

#Running model with intercept - AIC 511.21 ,ADJ R2 - 0.994
LN_lag1_nointercept <- step(lm(Sales~.-1,data=multdata_adv_modes_LN),direction = "backward",na.action=na.omit())
summary(LN_lag1_nointercept)
### Best Model of log without intercept and synergy :
# Sales ~ Catalogs_Winback + Catalogs_NewCust + Search + Portals +
#   `Catalogs_ExistCust:Catalogs_Winback` + `Catalogs_ExistCust:Catalogs_NewCust` +
#   `Catalogs_ExistCust:Portals` + `Catalogs_Winback:Catalogs_NewCust` +
#   `Catalogs_Winback:Mailings` + `Catalogs_Winback:Search` +
#   `Catalogs_Winback:Newsletter` + `Catalogs_Winback:Portals` +
#   `Catalogs_NewCust:Mailings` + `Catalogs_NewCust:Search` +
#   `Catalogs_NewCust:Newsletter` + `Catalogs_NewCust:Retargeting` +
#   `Catalogs_NewCust:Portals` + `Search:Newsletter` + `Search:Portals` +
#   `Newsletter:Portals` - 1
