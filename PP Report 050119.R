#PPAPAGNO Report
#Nick Cagliuso
#This is part of the overall slashline report, with a focus on the specific Approver named above

library(openxlsx)
library(dplyr)
library(tidyverse)
setwd("C:/Users/NCAGLIUSO/Desktop/Slashline/Backlogs")
ABACKLOG_05132019 <- read.xlsx("APPROVER-BACKLOG-05132019.xlsx")
ABACKLOG_06102019 <- read.xlsx("APPROVER-BACKLOG-06102019.xlsx")
ABACKLOG_05132019 <- ABACKLOG_05132019 %>% mutate(Time = "Before")
ABACKLOG_05132019 <- ABACKLOG_05132019 %>% mutate(Unique_ID = paste(.data$Unit, .data$PO_No, sep = "_"))
ABACKLOG_05132019 <- ABACKLOG_05132019 %>% filter(.data$Line == "1")
ABACKLOG_05132019 <- ABACKLOG_05132019 %>% distinct(.data$Unique_ID, .keep_all = TRUE)
ABACKLOG_06102019 <- ABACKLOG_06102019 %>% mutate(Time = "After")
ABACKLOG_06102019 <- ABACKLOG_06102019 %>% mutate(Unique_ID = paste(.data$Unit, .data$PO_No, sep = "_"))
ABACKLOG_06102019 <- ABACKLOG_06102019 %>% filter(.data$Line == "1")
ABACKLOG_06102019 <- ABACKLOG_06102019 %>% distinct(.data$Unique_ID, .keep_all = TRUE)
#Like in original slashline report, dates are changed manually weekly and a function is needed
#Unique ID is also added for the same reason as Slashline Report
#Reason is unknown but reqs are repeated in Approver Backlogs, so they need to be deduped

Raw_Data <- rbind(ABACKLOG_05132019, ABACKLOG_06102019)
Overlap <- inner_join(ABACKLOG_05132019, ABACKLOG_06102019, by = "Unique_ID")
Raw_Data <- mutate(Raw_Data, Category = ifelse(Raw_Data$Unique_ID %in% Overlap$Unique_ID, "Overlap",
                                        ifelse(Raw_Data$Unique_ID %in% ABACKLOG_05132019$Unique_ID, "Out",
                                        ifelse(Raw_Data$Unique_ID %in% ABACKLOG_06102019$Unique_ID, "In", "Not_Found"))))
True_Overlaps <- Raw_Data %>% filter(.data$Category == "Overlap")
True_Overlaps <- True_Overlaps %>% filter(.data$Time == "After")
Raw_Data <- Raw_Data %>% filter(!.data$Category =="Overlap")
Raw_Data <- rbind(Raw_Data, True_Overlaps)
Approver_Info <- Raw_Data %>% group_by(.data$User, .data$Category) %>% summarise(cnt = n())
Slash_Raw <- Approver_Info %>% spread(Category,cnt)
Slash_Raw[is.na(Slash_Raw)] <- 0
Slash_Raw <- Slash_Raw %>% mutate(PlusMinus = .data$In - .data$Out)
write.xlsx(Slash_Raw, file = "A_Slashline Data.xlsx")
write.xlsx(True_Overlaps, file = "ASlashline Overlaps.xlsx")
