library(RDCOMClient)
library(openxlsx)
library(xtable)

OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)

ids <- as.list(mail$ids)
i <- 1
while(i <= length(mail$id)){
    outMail[["To"]] = ids[i]
    outMail[["subject"]] = paste0("Report ", Sys.Date() - 1)
    outMail[["Attachments"]]$Add(attachments)
    outMail[["HTMLBody"]] = sprintf('
                                        Hello world, here is the table:
                                        Merry Christmas & a happy New Year!
                                        ') # add your html message content here
        outMail$Send()
        i= i+1
        
        }
