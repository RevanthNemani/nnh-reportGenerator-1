setwd('E:/nnh/reportsGenerator')
require(readxl)
require(dplyr)
require(writexl)
coll.names <-
  c(
    "docDate",
    "docType",
    "docNumber",
    "dueDate",
    "refNo",
    "narration",
    "aging",
    "currency",
    "debit",
    "credit"
  )
master <-
  read_xlsx("inputs/Customer Statement.xlsx",
            skip = 7,
            col_names = coll.names) %>% mutate(debit = as.numeric(debit),
                                               credit = as.numeric(credit))

splits <- which(master$docDate == "Net Balance")

y <- 1
reports <- list()
for (x in splits) {
  df <- as.data.frame(master[(y + 1):x,])
  z <- x - y
  df[z, 2] <- df[z, 1]
  df[z - 1, 2] <- df[z - 1, 1]
  df[c(z, z - 1), 1] <- NA
  df <- df[order(-df$aging),]
  reports[[gsub("/", "_", as.character(master[y, 1]))]] <-
    df %>% mutate(docDate = as.Date.numeric(as.numeric(docDate), origin = "1899-12-30"))
  y <- x + 1
}

rm(z, y, x, coll.names, master, splits, df)

lapply(names(reports), function(x)
  write_xlsx(reports[[x]], path = paste0("./outputs/", x, ".xlsx")))

rm(reports)
