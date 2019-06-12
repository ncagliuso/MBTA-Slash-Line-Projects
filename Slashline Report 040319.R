#Slashline Report
#Nick Cagliuso
#This project is updated weekly with "slashlines" by Buyer, Business Unit, and Threshold

library(tidyverse)
library(dplyr)
library(openxlsx)
setwd("C:/Users/NCAGLIUSO/Desktop/Slashline/Backlogs")
BACKLOG_05132019 <- read.xlsx("BACKLOG-05132019.xlsx")
BACKLOG_06102019 <- read.xlsx("BACKLOG-06102019.xlsx")
BACKLOG_05132019 <- BACKLOG_05132019 %>% mutate(Time = "Before")
BACKLOG_05132019 <- BACKLOG_05132019 %>% mutate(Unique_ID = paste(.data$Business_Unit, .data$Req_ID, sep = "_"))
BACKLOG_06102019 <- BACKLOG_06102019 %>% mutate(Time = "After")
BACKLOG_06102019 <- BACKLOG_06102019 %>% mutate(Unique_ID = paste(.data$Business_Unit, .data$Req_ID, sep = "_"))
#Backlog dates are currently manually changed every week; function should be made here to input new dates once
#Different reqs can have the same Req ID in different business units, so Unique_ID forces req uniqueness

Raw_Data <- rbind(BACKLOG_05132019, BACKLOG_06102019)
Overlap <- inner_join(BACKLOG_05132019, BACKLOG_06102019, by = "Unique_ID")
#Finds all Reqs in both dataframes

Raw_Data <- mutate(Raw_Data, Category = ifelse(Raw_Data$Unique_ID %in% Overlap$Unique_ID, "Overlap",
                                   ifelse(Raw_Data$Unique_ID %in% BACKLOG_05132019$Unique_ID, "Out",
                                   ifelse(Raw_Data$Unique_ID %in% BACKLOG_06102019$Unique_ID, "In", "Not_Found"))))
#"In" = in newer dataframe but not older one, "Out" = in older dataframe but not newer one

Raw_Data <- mutate(Raw_Data, Threshold = ifelse(Raw_Data$Req_Total < 3500, "Micro",
                                        ifelse(Raw_Data$Req_Total < 50000, "Medium",
                                        ifelse(Raw_Data$Req_Total >= 50000, "Large", "N/A"))))
True_Overlaps <- Raw_Data %>% filter(.data$Category == "Overlap")
True_Overlaps <- True_Overlaps %>% filter(.data$Time == "After")
#Buyer assignment changes mean reqs are double counted between "Before" and After", so "After" is filtered for accuracy

Raw_Data <- Raw_Data %>% filter(!.data$Category =="Overlap")
Raw_Data <- rbind(Raw_Data, True_Overlaps)
C_Raw <- Raw_Data %>% filter(.data$Buyer == "CFRANCIS")
J_Raw <- Raw_Data %>% filter(.data$Buyer == "JKIDD")

Buyer_Info <- Raw_Data %>% group_by(.data$Buyer, .data$Category) %>% summarise(cnt = n())
Slash_Raw <- Buyer_Info %>% spread(Category,cnt)
Threshold_Info <- Raw_Data %>% group_by(.data$Threshold, .data$Category) %>% summarise(cnt = n())
T_Slash_Raw <- Threshold_Info %>% spread(Category, cnt)
Unit_Info <- Raw_Data %>% group_by(.data$Business_Unit, .data$Category) %>% summarise(cnt = n())
U_Slash_Raw <- Unit_Info %>% spread(Category, cnt)
C_Info <- C_Raw %>% group_by(.data$Threshold, .data$Category) %>% summarise(cnt = n())
C_Slash_Raw <- C_Info %>% spread(Category, cnt)
J_Info <- J_Raw %>% group_by(.data$Threshold, .data$Category) %>% summarise(cnt = n())
J_Slash_Raw <- J_Info %>% spread(Category, cnt)

sourcing_execs <- (Buyer = c("AFLYNN", "TDIONNE","EWELSH", "RWEINER", "JDELALLA","KHALL","JMOYNIHAN"))
inventory_buyers <- (Buyer = c("AKNOBEL", "DMARTINOS", "KLOVE", "PHONG", "TSULLIVAN1", "NSEQUEA", "DWELCH"))
non_inventory_buyers <- (Buyer = c("CFRANCIS", "JKIDD", "JLEBBOSSIERE", "TTOUSSAINT","IATHERTON","JMIELE"))
Slash_Raw <- Slash_Raw %>% mutate(Buyer_Team = case_when(
  .data$Buyer %in% sourcing_execs ~ "SE",
  .data$Buyer %in% inventory_buyers ~ "INV",
  .data$Buyer %in% non_inventory_buyers ~ "NINV"
))

Slash_Raw[is.na(Slash_Raw)] <- 0
Slash_Raw <- Slash_Raw %>% mutate(PlusMinus = .data$In - .data$Out)
Slash_Raw <- Slash_Raw[,c("Buyer", "Buyer_Team", "In", "Out", "Overlap", "PlusMinus")]
T_Slash_Raw <- T_Slash_Raw %>% mutate(PlusMinus = .data$In - .data$Out)
U_Slash_Raw <- U_Slash_Raw %>% mutate(PlusMinus = .data$In - .data$Out)
C_Slash_Raw <- C_Slash_Raw %>% mutate(PlusMinus = .data$In - .data$Out)
J_Slash_Raw <- J_Slash_Raw %>% mutate(PlusMinus = .data$In - .data$Out)
#"PlusMinus" = the difference between req count in "After" and req count in "Before"

True_Overlaps <- True_Overlaps %>% filter(.data$Buyer == "CFRANCIS" | .data$Buyer == "JKIDD")

write.xlsx(Slash_Raw, file = "Slashline Data.xlsx")
write.xlsx(True_Overlaps, file = "Slashline Overlaps.xlsx")
write.xlsx(T_Slash_Raw, file = "Threshold Data.xlsx")
write.xlsx(Raw_Data, file = "Total Data.xlsx")
write.xlsx(U_Slash_Raw, file = "Unit Data.xlsx")
write.xlsx(J_Slash_Raw, file = "J Data.xlsx")
write.xlsx(C_Slash_Raw, file = "C Data.xlsx")
