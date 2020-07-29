library(odbc)
library(readxl)

#my_data <- read_excel(file.choose())
resultData <- read_excel("./GitHub/DataAgg/GeneralExpTemplate.xlsx", sheet = "DB Import")
resultData

for(i in 1:length(resultData$SampleID)){
  #print(resultData$SampleID[i])

  valueStr <- if(is.na(resultData$ValueString[i])) "" else trimws(resultData$ValueString[i])

  sql = paste("INSERT INTO Results (ResultID, SampleID, Metric, MetricType, ConfidenceScore, TimePoint, ValueString, ValueNumber, Unit, AssayPerformedDate, createby, createdt AuditSequence) Values(next value for ResultIDSeq,", 
        trimws(resultData$SampleID[i]), ",'", trimws(resultData$Metric[i]), "','", trimws(resultData$MetricType[i]), "',", trimws(resultData$ConfidenceScore[i]), ",'",trimws(resultData$TimePoint[i]), "','", 
        valueStr, "',", trimws(resultData$ValueNumber[i]), ",'", trimws(resultData$Unit[i]), "','", trimws(resultData$AssayPerformedt[i]), "','njablonski', SYSDATETIME(),", 1,")", sep = "")

  print(sql)

  #try(dbGetQuery(con, sql), silent = FALSE, outFile = getOption("try.outFile", default = stderr()))
}

print("Successfully loaded Results data")



