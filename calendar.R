library(RTL)
library(xlsx)
library(timeDate)
library(lubridate)

MyZone = "UTC"

Sys.setenv(TZ = MyZone)

currentyear = getRmetricsOptions("currentYear") 

datestart = paste0(currentyear, "-01-01")
dateend = paste0(currentyear, "-12-31")

daycal = timeSequence(from = datestart, to = dateend, by="day", zone = MyZone)

tradingdays = daycal[isBizday(daycal, holidayNYSE())]

tradingdays = as.Date(tradingdays)

tradingdays = data.frame(tradingdays)

holidays = holidaysOil

expiry = expiry_table

trade_cycle = tradeCycle

write.xlsx(holidays, file="tradeCalendar.xlsx", sheetName="holidays", 
           col.names=TRUE, row.names=FALSE)

write.xlsx(expiry, file="tradeCalendar.xlsx", sheetName="expiry", 
           col.names=TRUE, row.names=FALSE, append=TRUE)

write.xlsx(tradingdays, file="tradeCalendar.xlsx", sheetName="tradingdays", 
           col.names=TRUE, row.names=FALSE, append=TRUE)

write.xlsx(trade_cycle, file="tradeCalendar.xlsx", sheetName="tradecycle", 
           col.names=TRUE, row.names=FALSE, append=TRUE)


