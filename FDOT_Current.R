# library(rvest)
# library(dplyr)
# library(tidyr)
# library(rstudioapi)
packages = c("rvest","dplyr","tidyr","rstudioapi", "openxlsx")

package.check <- lapply(packages, FUN = function(x) {
  if (!require(x, character.only = TRUE)) {
    install.packages(x, dependencies = TRUE)
    library(x, character.only = TRUE)
  }
})

options(warn=-1)
#options(warn=0) to reset

#Specifying the url for Current FDOT Advertisements
url <- 'http://www2.dot.state.fl.us/procurement/ProfessionalServices/advertise/advnew.shtml'
webpage <- read_html(url)

#Using CSS selectors to scrap the ads
current_ads_html <- html_nodes(webpage,'pre')

#Converting the data to text
ads_data <- html_text(current_ads_html)

ads_data <- as.vector(unlist(strsplit(ads_data,"\r\n")),mode="list") 

ads_data <- ads_data[which( unlist( lapply(ads_data, nchar) ) > 0 )]

District <- c()
Major_Work <- c()
email <- c()
Minor_Work <- c()
uuwg <- c()
contract <- c()
selMethod <- c()
PLI <- c()
FMN <- c()
description <- c()
PM <- c()
Amount <- c()
AdDate <- c()
rDeadlineDate <- c()
LLdate <- c()
SLdate <- c()
FSMeeting <- c()
SSNA<- c()
spN <- c()
techMeeting <- c()
firstMeeting <- c()
attention <- c()
phone <- c()
scopeDate <- c()
respDeadline <- c()
lastMeeting <-c()
secondTechMeeting <- c()
viewScope <- c()
numContract <- c()
thirdTechMeeting <- c()

for (i in 1:20){
  ads_data <- gsub(paste0("Number of contracts that may be awarded from this advertisement:  ",i , " "), "",ads_data)
  
}

if (length(ads_data) == 0) {
  winDialog(
    type = "ok",
    "FDOT's Advertisement Website is currenttly down!!!")
  
} else {

for (i in 1:length(ads_data)) {
  
  pattern <- "Response Deadline              :"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    respDeadline <- c(respDeadline, i)
  }
  pattern <- "Scope Meeting Date             :"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    scopeDate <- c(scopeDate, i)
  }
  
  pattern <- "Phone:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    phone <- c(phone, i)
  }
  
  pattern <- "Attn\\.:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    attention <- c(attention, i)
  }
  
  pattern <- "1st Date Negotiations Meetings"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    firstMeeting <- c(firstMeeting, i)
  }
  
  pattern <- "Tech\\. Rev\\. Cmte Meeting Date"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    techMeeting <- c(techMeeting, i)
  }
  
  pattern <- "Special Notes:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    spN <- c(spN, i)
  }
  
  pattern <- "See Standard Notes Above:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    SSNA <- c(SSNA, i)
  }
  
  pattern1 <- "   DISTRICT"
  pattern2 <- "   CENTRAL OFFICE"
  pattern3 <- "           TURNPIKE"
  
  if(length(grep(pattern1, ads_data[[i]], ignore.case = TRUE)) >0 || length(grep(pattern2, ads_data[[i]], ignore.case = TRUE)) >0 || length(grep(pattern3, ads_data[[i]], ignore.case = TRUE)) >0) {
    District <- c(District, i)
  }
  
  pattern <- "Major Work   :"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    Major_Work <- c(Major_Work, i)
  }
  
  pattern <- "Minor Work   :"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    Minor_Work <- c(Minor_Work, i)
  }
  
  pattern <- "Under-Utilized Work Groups"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    uuwg <- c(uuwg, i)
  }
  
  pattern <- "Contract     :"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    contract <- c(contract, i)
  }
  
  pattern <- "Selection Method"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    selMethod <- c(selMethod, i)
  }
  
  pattern <- "Professional Liability Insurance \\$:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    PLI <- c(PLI, i)
  }
  
  pattern <- "Financial Management Number"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    FMN <- c(FMN, i)
  }
  
  pattern <- "Project Description:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    description <- c(description, i)
  }
  
  pattern <- "Project Manager:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    PM <- c(PM, i)
  }
  
  pattern <- "Contract Amount/Limit"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    Amount <- c(Amount, i)
  }
  
  pattern <- "Advertisement Date             :"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    AdDate <- c(AdDate, i)
  }
  
  pattern <- "Response Deadline Date         :"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    rDeadlineDate <- c(rDeadlineDate, i)
  }
  
  pattern <- "Longlist\\(Tech\\. Rev\\. Cmte\\."
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    LLdate <- c(LLdate, i)
  }
  
  pattern <- "Shortlist Selection Date"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    SLdate <- c(SLdate, i)
  }
  
  pattern <- "Final Selection Meeting Date"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    FSMeeting <- c(FSMeeting, i)
  }
  
  pattern1 <- "Respond To:"
  pattern2 <- "serv@dot\\.state\\.fl\\.us"
  
  if(length(grep(pattern1, ads_data[[i]], ignore.case = TRUE)) >0 && length(grep(pattern2, ads_data[[i+1]], ignore.case = TRUE)) >0) {
    email <- c(email, i)
  }
  
  
  pattern <- "Last Date Negotiations Meetings:"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    lastMeeting <- c(lastMeeting, i)
  }
  
  pattern <- "2nd Tech Rev\\. Cmte Meeting"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    secondTechMeeting <- c(secondTechMeeting, i)
  }
  
  pattern <- "3rd Tech Rev\\. Cmte Meeting"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    thirdTechMeeting <- c(thirdTechMeeting, i)
  }
  
  pattern <- "View proposed scope"
  if(length(grep(pattern, ads_data[[i]], ignore.case = TRUE)) >0) {
    viewScope <- c(viewScope, i)
  }
  
}


positions <- c(Major_Work,   email, District, Minor_Work, uuwg, contract, selMethod, PLI, FMN, description,
               PM, Amount, AdDate, rDeadlineDate, LLdate, SLdate, FSMeeting, spN, SSNA, techMeeting, 
               firstMeeting, attention, phone, scopeDate, respDeadline, lastMeeting, secondTechMeeting,
               viewScope, numContract, thirdTechMeeting)
positions <- sort(positions)

N <- length(Major_Work)

List <- data.frame(Topic = rep(-9999,length(positions)), Item = rep(-9999,length(positions)))

for (k in positions[-length(positions)]){
  
  List$Topic[match(k,positions)] <- ads_data[[k]]
  
  if(positions[match(k,positions)+1]-positions[match(k,positions)] == 1) {
    List$Item[match(k,positions)] <-""
  } else {
    
    for (i in (k+1):(positions[match(k,positions)+1]-1)) {
      List$Item[match(k,positions)] <- paste(List$Item[match(k,positions)], ads_data[i] , sep="/")
      
    }
  }
}
List$Item <- gsub("-9999", "", List$Item)
index <- which(List$Item == "")
collSeparated <- List[index,] %>%
  separate( Topic, c("beforecoll", "aftercoll"), ":")
List$Item[index] <- collSeparated$aftercoll
List$Topic[index] <- collSeparated$beforecoll
List[is.na(List)] <- ""

for (i in 1:nrow(List)) {
  
  pattern <- "Response Deadline              "
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Response Deadline"
  }
  pattern <- "Scope Meeting Date             "
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Scope Meeting Date"
  }
  
  pattern <- "Phone"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Attn\\."
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Attn."
  }
  
  pattern <- "1st Date Negotiations Meetings"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Tech\\. Rev\\. Cmte Meeting Date"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Tech. Rev. Cmte Meeting Date"
  }
  
  pattern <- "Special Notes"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "See Standard Notes Above"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  
  pattern <- "Under-Utilized Work Groups"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Contract     "
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Contract"
  }
  
  pattern <- "Selection Method"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Professional Liability Insurance \\$"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Professional Liability Insurance $"
  }
  
  pattern <- "Financial Management Number"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Project Description"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Project Manager"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Contract Amount/Limit"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Advertisement Date             "
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Advertisement Date"
  }
  
  pattern <- "Response Deadline Date         "
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Response Deadline Date"
  }
  
  pattern <- "Longlist\\(Tech\\. Rev\\. Cmte\\."
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "Longlist(Tech. Rev. Cmte.)"
  }
  
  pattern <- "Shortlist Selection Date"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "Final Selection Meeting Date"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern1 <- "Respond To"
  pattern2 <- "serv@dot\\.state\\.fl\\.us"
  
  if(length(grep(pattern1, List$Topic[i], ignore.case = TRUE)) >0 && length(grep(pattern2, List$Topic[i+1], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern1
  }
  
  
  pattern <- "Last Date Negotiations Meetings"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
  pattern <- "2nd Tech Rev\\. Cmte Meeting"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "2nd Tech Rev. Cmte Meeting"
  }
  
  pattern <- "3rd Tech Rev\\. Cmte Meeting"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- "3rd Tech Rev. Cmte Meeting"
  }
  
  pattern <- "View proposed scope"
  if(length(grep(pattern, List$Topic[i], ignore.case = TRUE)) >0) {
    List$Topic[i] <- pattern
  }
  
}

for (i in 1:20){
  List$Item <- gsub("  ", " ", List$Item)
  List$Topic <- gsub("  ", " ", List$Topic)
  
}
#HEEEEEEEERRRRREEEEE
List <- List[List$Topic!=-9999,]

index <- which(List$Item == "")
collSeparated <- List[index,] %>%
  separate( Topic, c("beforecoll", "aftercoll"), ":")
List$Item[index] <- collSeparated$aftercoll
List$Topic[index] <- collSeparated$beforecoll
List[is.na(List)] <- ""

for (i in 1:10){
  List$Item <- trimws(List$Item)
  List$Topic <- trimws(List$Topic)
}


List$Item[List$Topic != "Major Work :"] <- gsub("/", "", List$Item[List$Topic != "Major Work :"]) 

List <- List[-nrow(List),]

District1 <- length(grep("DISTRICT 1", List$Topic))
District2 <- length(grep("DISTRICT 2", List$Topic))
District3 <- length(grep("DISTRICT 3", List$Topic))
District4 <- length(grep("DISTRICT 4", List$Topic))
District5 <- length(grep("DISTRICT 5", List$Topic))
District6 <- length(grep("DISTRICT 6", List$Topic))
District7 <- length(grep("DISTRICT 7", List$Topic))

CentralOffice <- length(grep("CENTRAL OFFICE", List$Topic))

Turnpike <- length(grep("TURNPIKE", List$Topic))

NumProjects = District1 + District2 + District3 + District4 + District5 + District6 + District7 + CentralOffice + Turnpike

winDialog(
  type = "ok",
  paste0("There are totally ",NumProjects," planned advertisements",
         "\n D1: ",District1,
         "\n D2: ",District2,
         "\n D3: ",District3,
         "\n D4: ",District4,
         "\n D5: ",District5,
         "\n D6: ",District6,
         "\n D7: ",District7,
         "\n Central office: ",CentralOffice,
         "\n Turnpike: ",Turnpike
  ))

Major_works_List <- c()

for (i in 1:nrow(List)) {
  if(List$Topic[i]=="Major Work :"){
    Major_works_List <- paste0(Major_works_List,List$Item[i])
  }
  
}

Major_works_items <- unlist(strsplit(Major_works_List, "/ "))
Major_works_items <- Major_works_items[Major_works_items!=""]

freq_table <- as.data.frame(table(Major_works_items))
freq_table <- freq_table%>%
  arrange(desc(Freq))

listtobedisplayed <-""
for (i in 1:min(10, nrow(freq_table))){
  listtobedisplayed<-paste0(listtobedisplayed, "\n ",freq_table$Major_works_items[i],": ",freq_table$Freq[i])
}

if(nrow(freq_table)>10){
  message <- "Here are the frequency of top-10 Major Works under planned advertisements:"
} else {
  message <- "Here are the frequency of Major Works under planned advertisements:"
  
}
winDialog(
  type = "ok",
  paste0(message,
         listtobedisplayed)
)

List$Item <- gsub("/", "", List$Item) 
List$District <- ""


for (i in 1:nrow(List)){
  if(List$Topic[i] %in% c(paste("DISTRICT ", 1:7, sep = ""), "CENTRAL OFFICE", "TURNPIKE")){
    List$District[i] <- List$Topic[i]
  } else {
    List$District[i] <- List$District[i-1]
  }
  
}

List$ID <- 1

for (i in 2:nrow(List)){
  if(List$Topic[i] %in% c(paste("DISTRICT ", 1:7, sep = ""), "CENTRAL OFFICE", "TURNPIKE")){
    List$ID[i] <- List$ID[i-1] + 1
  } else {
    List$ID[i] <- List$ID[i-1]
  }
  
}


setwd(dirname(getSourceEditorContext()$path))

mainDir <- getwd()
subDir <- "Archive_Current"



if (!file.exists(subDir)){
  
  dir.create(file.path(mainDir, subDir))
  
}


drop <- c(paste("DISTRICT ", 1:7, sep = ""), "View proposed scope", "CENTRAL OFFICE", "TURNPIKE", "Shortlist Selection Date Not Available", "Major Work")
tAds <- List %>%
  group_by(Topic, ID, District)%>%
  summarise(Item = first(Item))%>%
  spread(key = Topic, value = Item)%>%
  arrange(ID)


tAds <- tAds[,!(names(tAds) %in% drop)]

if(!file.exists("fdot_ads_Current.xlsx")) {
  write.xlsx(tAds, "fdot_ads_Current.xlsx")
  write.xlsx(tAds, paste0("./Archive_Current/fdot_ads_Current",format(Sys.Date(), "_%b_%d_%Y"),".xlsx"), row.names=FALSE)
  
} else if (nrow(tAds) == nrow(read.xlsx("fdot_ads_Current.xlsx")) && ncol(tAds) == ncol(read.xlsx("fdot_ads_Current.xlsx"))){
  if (tAds == read.xlsx("fdot_ads_Current.xlsx")) {
  winDialog(type = "ok", "No new advertisements!")
  write.xlsx(tAds, paste0("./Archive_Current/fdot_ads_Current",format(Sys.Date(), "_%b_%d_%Y"),".xlsx"), row.names=FALSE)
  }

} else {
  newAds <- anti_join(tAds, read.xlsx("fdot_ads_Current.xlsx"), by = c( "District", "Attn.", "Contract", "Phone", "Advertisement Date" = "Advertisement.Date", "Final Selection Meeting Date" = "Final.Selection.Meeting.Date"))
  if(nrow(newAds) == 0) {
    winDialog(type = "ok", "No new advertisements!")
    write.xlsx(tAds, paste0("./Archive_Current/fdot_ads_Current",format(Sys.Date(), "_%b_%d_%Y"),".xlsx"), row.names=FALSE)
    write.xlsx(tAds, "fdot_ads_Current.xlsx", row.names=FALSE)
  } else {
  winDialog(type = "ok", paste0("There are ", nrow(newAds), " new ads that is saved in fdot_ads_Current_OnlyNewAds.xlsx"))
  write.xlsx(tAds, "fdot_ads_Current.xlsx", row.names=FALSE)
  write.xlsx( newAds,  "fdot_ads_Current_OnlyNewAds.xlsx", row.names=FALSE)
  write.xlsx(tAds, paste0("./Archive_Current/fdot_ads_Current",format(Sys.Date(), "_%b_%d_%Y"),".xlsx"), row.names=FALSE)
  }
}
}