library(data.table)
library(stringr)
library(XLConnect)


#' Summarise data columns.
#'
#' @param dt Data table to summarise.
#' @param writepath Path + file name to save xlsx.
#' @return Data summary.
#' @export
summarise <- function(dt, writepath = getwd()) {

   # rows & columns
   dims <- dim(dt)

   cat("Data has", dims[1], "rows and", dims[2], "columns.\n")
   cat("Summarising Columns...\n")

   # columns
   out <- dt[, .(N = 1:ncol(dt),
                 ColName = names(dt),
                 uniqueN = lapply(.SD, uniqueN),
                 Class = lapply(.SD, function(x) class(x)[1]),
                 Missing = lapply(.SD, function(x) sum(is.na(x))),
                 MissingRate =  lapply(.SD, function(x) sum(is.na(x))/.N),
                 Min = lapply(.SD, min, na.rm = TRUE),
                 Max = lapply(.SD, max, na.rm = TRUE))]

   cols <- out[uniqueN <= 5 & uniqueN > 2, ColName]

   out[, str_c("Level", 1:5) := ""]


   out[ColName %in% cols,
       str_c("Level", 1:5) :=
          data.table(t(dt[, lapply(.SD, function(x) c(unique(x), rep("", 5 - uniqueN(x)))),
                          .SDcols = (cols)]))]

   out <- cbind(out[, 1:6], out[, lapply(.SD, function(x) lapply(x, as.character)), .SDcols = 7:13])
   out <- out[, lapply(.SD, unlist)]
   # out[Class == "numeric", `:=`(Min = as.numeric(Min), Max = as.numeric(Max))]

   cat("Writing Excel...\n")
   file.copy(from = "D:/Users/shane.mono/00. Random/Snippets/Data Summary Template.xlsx",
             to = writepath, overwrite = TRUE)
   wb <- loadWorkbook(writepath)
   setStyleAction(wb, XLC$"STYLE_ACTION.NONE")
   writeWorksheet(wb, out, "Data Summary", startRow=9, startCol=2, header=TRUE)
   writeWorksheet(wb, dims, "Data Summary", startRow=5, startCol=3, header=FALSE)



   saveWorkbook(wb)

   return(out)

}


#' Copy to clipboard
#'
#' @param x Data to copy to clipboard.
#' @param row.names Whether to include row names. Default to \code{FALSE}.
#' @param excel.names Whether to convert col names to Title Case. Default to \code{TRUE}.
#' @export
cp <- function(x, row.names = FALSE, excel.names = TRUE) {
   if(excel.names)
      write.table(excel.names(x), "clipboard", row.names = row.names, sep = "\t") else
         write.table(x, "clipboard", row.names = row.names, sep = "\t")

}


#' Clean column names in a data table
#'
#' @param dt Data table to clean names.
#' @value Returns the data table with cleaned names.
#' @export
clean.names <- function(dt) {

   x <- names(dt)
   x <- str_to_lower(x)
   x <- gsub("[^a-zA-Z0-9]", "\\.", x)
   x <- gsub("(\\.)+", "\\.", x)
   x <- gsub("\\.$", "", x)
   x <- gsub("\\.([0-9]+)$", "\\1", x)

   names(dt) <- x
   return(dt)


}


#' Convert col names to title names
#'
#' @param x Data to convert col names back to Proper Title Case.
#' @value Returns the data table with Title Case col names.
#' @export
excel.names <- function(dt) {

   x <- names(dt)
   x <- gsub("\\.", " ", x)
   x <- str_to_title(x)
   names(dt) <- x
   return(dt)

}


#' Convert all columns with "date" in col names to Date format.
#'
#' @param dt Data table.
#' @param format Date format. Default to \code{"%Y-%m-%d"}.
#' @value Returns the data table.
#' @export
convert.dates <- function(dt, format = "%Y-%m-%d") {

   datecols <- names(dt)[names(dt) %like% "date"]
   cat(length(datecols), " date columns found.\n")

   for(i in 1:length(datecols)) {
      dt[, (datecols[i]) := as.Date(get(datecols[i]), format = format)]

   }
   return(dt)

}
