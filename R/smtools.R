# ======== SM TOOLS ======== #
# ======== 0000-00-00 ======== #
# -------- shane.mono -------- #

library(data.table)
library(stringr)
library(XLConnect)
library(progress)


#' Summarise data columns.
#'
#' @param dt Data table to summarise.
#' @param writepath Path + file name to save xlsx; can be \code{NULL}.
#' @param append Whether the summary should be appended to the workbook; any 
#' strings except "Data Summary" mean append; otherwise the summary table is
#' written on the worksheet "Data Summary".
#' @return Data summary.
#' @export
summarise <- function(dt, writepath = str_c(getwd(), "/Data Summary.xlsx"), 
                      append = "Data Summary") {
  
  # rows & columns
  dims <- dim(dt)
  
  cat("Data has", dims[1], "rows and", dims[2], "columns.\n")
  cat("Summarising Columns...\n")
  
  # summarise by column
  out <- dt[, .(N = 1:ncol(dt),
                ColName = names(dt),
                uniqueN = lapply(.SD, uniqueN),
                Class = lapply(.SD, function(x) class(x)[1]),
                Missing = lapply(.SD, function(x) sum(is.na(x))),
                MissingRate =  lapply(.SD, function(x) sum(is.na(x))/.N),
                EmptyRate = lapply(.SD, function(x) 
                  if(class(x) == "character") sum(x == "")/.N),
                Min = lapply(.SD, min, na.rm = TRUE),
                Max = lapply(.SD, max, na.rm = TRUE))]
  
  
  # find columns where unique levels are between 2 and 5 for level summary
  cols <- out[uniqueN <= 5 & uniqueN > 2, ColName]
  
  # add five extra columns for levels
  out[, str_c("Level", 1:5) := ""]
  
  # find levels and write to table
  out[ColName %in% cols,
      str_c("Level", 1:5) :=
        data.table(
          t(dt[, 
               lapply(.SD, function(x) c(unique(x), rep("", 5 - uniqueN(x)))),
               .SDcols = (cols)]
            ))]
  
  out <- cbind(out[, 1:7], out[, 
                               lapply(.SD, 
                                      function(x) lapply(x, as.character)), 
                               .SDcols = 8:14])
  out <- out[, lapply(.SD, unlist)]

  # write output
  if(!is.null (writepath) ) {
    cat("Writing Excel...\n")
    if (append == "Data Summary") 
      file.copy(from = "./data/Data Summary Template.xlsx",
                to = writepath, overwrite = TRUE)
    wb <- loadWorkbook(writepath)
    setStyleAction(wb, XLC$"STYLE_ACTION.NONE")
    if(append != "Data Summary") {
      cloneSheet(wb, 1, append)
      clearRange(wb, append, c(2, 10, 15, 150))  }
    writeWorksheet(wb, out, append, startRow=9, startCol=2, header=TRUE)
    writeWorksheet(wb, dims, append, startRow=5, startCol=3, header=FALSE)
    writeWorksheet(wb, append, append, startRow=2, startCol=2, header=FALSE)
    
    saveWorkbook(wb)
  }
  
  return(out)
  
}


#' Copy to clipboard.
#'
#' @param x Data to copy to clipboard.
#' @param row.names Whether to include row names. Default to \code{FALSE}.
#' @param excel.names Whether to convert col names to Title Case. Default to 
#' \code{TRUE}.
#' @export
cp <- function(x, row.names = FALSE, excel.names = TRUE) {
  if(excel.names)
    write.table(excel.names(x), "clipboard", row.names = row.names, sep = "\t") else
      write.table(x, "clipboard", row.names = row.names, sep = "\t")
  
}


#' Copy to clipboard (large file).
#'
#' @param x Data to copy to clipboard.
#' @param row.names Whether to include row names. Default to \code{FALSE}.
#' @param excel.names Whether to convert col names to Title Case. Default to 
#' \code{TRUE}.
#' @export
cp2 <- function(x, row.names = FALSE, excel.names = TRUE) {
  if(excel.names)
    write.table(excel.names(x), "clipboard", row.names = row.names, sep = "\t") else
      write.table(x, "clipboard-16384", row.names = row.names, sep = "\t")
  
}

#' Clean column names in a data table.
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


#' Convert col names to title names.
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
#' @param additional.cols Additional date columns to be parsed other than columns 
#' with the string "date".
#' @param format Date format. Default to \code{"%Y-%m-%d"}.
#' @value Returns the data table.
#' @details \code{convert.dates} automatically finds columns with "date" in the 
#' name and converts to \code{Date}.
#' Additional columns can be parsed on for parsing. Cannot handle multiple date 
#' formats.
#' @export
convert.dates <- function(dt, additional.cols = NULL, format = "%Y-%m-%d") {
  
  if(!is.null(additional.cols)) 
    datecols <- c(names(dt)[names(dt) %Like% "date"], additional.cols) 
  else
    datecols <- names(dt)[names(dt) %Like% "date"]
  
  cat(length(datecols), " date columns found.\n")
  
  dt[, (datecols) := lapply(.SD, as.Date, format = format), .SDcols = datecols]
  
  return(dt)
  
}



#' Generate sript headers.
#' @param title Will convert to upper letters automatically.
#' @export
gen.headers <- function(title = "WRITE STH NICE") {
  str_c("# ======== ", str_to_upper(title), " ======== #\n# ======== ", 
        format(Sys.time(), "%Y-%m-%d"),
        " ======== #\n# -------- shane.mono -------- #") %>% 
    writeClipboard()
}



#' Save some time typing ". Parse one string to a vector.
#' 
#' @param string A string to be parsed.
#' @param sep Default \code{,}.
#' @export
parse2v <- function(string, sep = ",") {
  return(str_trim( unlist(str_split(string, ",")) ))
}


#' Pivot table inspired grouping.
#' 
#' @param col Vector; use in \code{data.table}.
#' @param strgroup Grouping string in the format of \code{"Other+Unknown=Other"}.
#' Supports fuzzy match. \code{"O+U=Other"}, if Other and Unknown are the only 
#' categories starting with O and U. Supports \code{"NA+U=Other"}. For blanks, 
#' use \code{"$+O=Other"}.
#' @param suffix Suffix of created grouped column.
#' @param new Create a new column of the grouped categories. Default to 
#' \code{FALSE}.
#' @details No debugging / error catching.
#' @value A lookup table with unique values from \code{col} and grouped values.
#' @example billing[, pgroup(cust.group, "O+U=Other")] # group Other and Unknown 
#' into Other
#' billing[, cust.group2 := pgroup(cust.group, "$+O+U+NA=Other", new = TRUE)]
#' @export
pgroup <- function(col, strgroup, suffix = ".grp", new = FALSE) {
  
  colnm <- deparse(substitute(col))
  # col0 is original cols, col and col2 is the uniqued version
  col0 <- copy(col)
  col <- unique(col)
  col2 <- col
  # parsed vector of all groupings
  vec <- unlist(str_split(strgroup, ","))
  
  for(i in 1:length(vec)) {
    # parse elements of the string in one grouping
    ele <- unlist(str_split(vec[i], "\\+|="))
    # special NA treatment
    if (any(ele=="NA")) col2[is.na(col)] <- ele[length(ele)]
    # a list of values without the last element of the parsed string list, 
    # which is the new group name
    children <- unlist(lapply(ele[-length(ele)],
                              function(x) col[col %like% str_c("^", x)]))
    # replace every parsed entry of the column with the new group name
    col2[col2 %in% children] <- ele[length(ele)]
    
  }
  
  # construct output data table
  out <- data.table(col, col2)
  setnames(out, c(colnm, str_c(colnm, suffix)))
  
  if(new) {
    for(i in 1:length(col)) col0[col0 == col[i]] <- col2[i]
    return(col0)
  } else {
    return(out)
  }
  
}


#' Examine whether one column in a data table is unique by another.
#' 
#' @param keycols A list of column names or one column name without quote; the key
#' which other columns should be unique by.
#' @param ... Other columns to investigate uniqueness.
#' @details No debugging / error catching.
#' @value Returns a vector with the same length as \code{...} with values 
#' \code{TRUE} or \code{FALSE}.
#' @example dt[, ifunique(id, name, status)]
#' dt[, ifunique(list(tour.code, date), tour.name, adult.price)]
#' @export
ifunique <- function(keycols, ...) {
  
  if(class(x) == "list") n <- length(keycols) else 
    {keycols <- list(keycols); n <- 1}
  testcols <- list(...)
  out <- c()
  for(i in 1:length(testcols)) {
    testcol <- testcols[[i]]
    # create a dt of keycols + testcol
    dt <- data.table(do.call(cbind, keycols), testcol)
    # find duplicate rows by [, .N, by]
    dupe <- dt[, .N, by = c(paste0("V", 1:n), "z")] %>% 
      .[, .N, by = c(paste0("V", 1:n))][N > 1]
    # if the duplicate table has any entry, keys are not unique
    if(nrow(dupe) == 0) res <- TRUE else res <- FALSE
    out <- c(out, res)
  }
  return(out)
  
  
}


#' Fuzzy match; suggest possible match results.
#' 
#' @param x A vector to be matched.
#' @param y A vector to find the match from.
#' @param matchdt Optional. A data table with the first column being \code{y}.
#' This is to help facilitate fuzzy match to raw names of alread matched names,
#' and then merge to master names in one go.
#' @param n Number of match results; default to 2.
#' @param ... Further arguments to be passed onto \code{stringdistmatrix}.
#' @details No debugging / error catching. Used \code{stringdistmatrix} from 
#' \code{stringdist}.
#' @value Returns a \code{data.table} with the vector to be matched and \code{n} 
#' columns of closest matches.
#' @example fuzzymatch(raw.names, master.names)
#' @export
fuzzymatch <- function(x, y, matchdt = NULL, n = 2, ...) {
  
  x <- unique(x) 
  y <- unique(y)
  # generate string dist matrix
  match.matrix <- stringdist::stringdistmatrix(str_to_upper(x), 
                                               str_to_upper(y), ...)
  # rank string dist for every row
  rank.matrix <- t(apply(match.matrix, 1, frank, ties.method = "random"))
  # find the top ones
  for (i in 1:n) {
    match.index <- apply(rank.matrix, 1, function(x) which(x == i)) %>% unlist
    match.res <- y[match.index]
    if (i == 1) out <- match.res else out <- cbind(out, match.res)
  }
  
  out <- data.table(x, out)
  setnames(out, c("raw", str_c("match", 1:n)))
  
  # if matchdt is provided, merge with it
  if(!is.null(matchdt)) {
    
    out[, N := 1:.N]
    setnames(matchdt, 1, "match1")
    out <- merge(out, matchdt, by = "match1", all.x = TRUE)
    out <- out[order(N)]
    out[, N := NULL]
    
  }
  
  return(out[])
  
}


#' Set difference between two vectors; imagine the Venn Diagram.
#' 
#' @param x A vector.
#' @param y A vector.
#' @param res Return data; see \code{value}.
#' @value Returns lengths, # common entries, # entries only in x and in y. If 
#' \code{res = 0}, returns
#' common elements, \code{res = 1} returns elements in x only, and 
#' \code{res = 2} returns in y only.
#' Any other values would return the summary table only. 
#' @example set.diff(code1, code2)
#' @export
set.diff <- function(x, y, res = -1) {
  
  xnm <- deparse(substitute(x))
  ynm <- deparse(substitute(y))
  x <- unique(x)
  y <- unique(y)
  
  out <- data.table(xnm = c(length(x), length(intersect(x, y)), 
                            length(setdiff(x, y)), 0),
                    ynm = c(length(y), length(intersect(x, y)), 
                            0, length(setdiff(y, x))))
  setnames(out, c(xnm, ynm))
  out <- cbind(value = c("length", "common", "in x only", "in y only"), out)
  
  if (res == 0) outv <- intersect(x, y)
  if (res == 1) outv <- setdiff(x, y)
  if (res == 2) outv <- setdiff(y, x)
  
  if(res == 0) return(list(setdiff = out, common = outv)) else
    if(res == 1) return(list(setdiff = out, in.x.only = outv)) else
      if (res == 2) return(list(setdiff = out, in.y.only = outv)) else
        return(out)

}


#' Progress bar for loops - no need to set up every time.
#' 
#' @param ind Index in the loop.
#' @param total Size of loop.
#' @param startind Starting index of the loop.
#' @example setup.pb(100)
#' for(i in 12:100) {
#'     pb.tick(12, i , 100)
#'     Sys.sleep(0.02)
#' }
#' @details No error catching, especially around interrupted loops and 
#' start index.
#' @export
setup.pb <- function(total) {
  progressbarsm <<- progress_bar$new(
    format = "  Processed :ind of :total [:bar] :percent in :elapsed",
    clear = FALSE, total = total)
}
#' @export
pb.tick <- function(startind, ind, total) {
  if (startind > ind) stop("  Start index > index.")
  if (startind == ind) {
    cat("  Starting at", format(Sys.time(), "%H:%M"), "...\n")
    progressbarsm$tick(ind - 1)
  }
  progressbarsm$tick(tokens = list(ind = ind, total = total))
  if (ind == total) {
    cat("\n  Finished at", format(Sys.time(), "%H:%M"), ".\n") 
    rm(progressbarsm, pos = ".GlobalEnv")
  }
}


#' Detect string pattern.
#' 
#' @param col A vector of string.
#' @param pattern Default to \code{Az0}. Supports \code{Aa0.} where each 
#' transforms all big letters,
#' all small letters, all numbers, everything else to \code{A, a, 0} 
#' and \code{.}.
#' @example dt[, .N, by = str.pattern(tour.code)][order(-N)]  # count of
#'  code patterns
#' @export
str.pattern <- function(col, pattern = "Aa0") {
  
  if(pattern %like% "A") col <- gsub("[A-Z]", "A", col)
  if(pattern %like% "a") col <- gsub("[a-z]", "a", col)
  if(pattern %like% "0") col <- gsub("[0-9]", "0", col)
  if(pattern %like% "\\.") col <- gsub("[^0-9A-Za-z]", "\\.", col)
  
  return(col)
  
}


#' Case-insensitive %like%.
#' 
#' @param x String to be detected.
#' @param pattern String pattern (regex), case insensitive.
#' @export
`%Like%` <- function (x, pattern) { 
  stringi::stri_detect_regex(x, pattern, case_insensitive=TRUE)
}


#' Detect NA or blank.
#' 
#' @param x A vector
#' @example dt <- dt[!naorb(col)]
#' @export
naorb <- function(x) {
  x <- str_trim(x)
  out <- is.na(x)
  out2 <- x == ""
  out <- out | out2
  return(out)
}


#' Difference of dates in months.
#' 
#' @param date Date in format \code{%Y-%m-%d}, the \code{-} can be any single
#'  symbol, such as \code{/}.
#' @param refdate Reference date in the same format.
#' @param since If \code{TRUE}, months since date to reference date, eg from 
#' date to now (ref), else
#' month from reference date to date, eg from now (ref) to a future dat)
#' @values A vector of same length
#' @export
months.diff <- function(date, refdate, since = TRUE) {
  
  if(class(date) == "Date" & class(refdate) == "Date") {
    diff.in.months <- 12 * (year(refdate) - year(date)) + 
      month(refdate) - month(date)
  } else {
    y.refdate <- as.integer(str_sub(refdate, 1, 4))
    m.refdate <- as.integer(str_sub(refdate, 6, 7))
    y.date <- as.integer(str_sub(date, 1, 4))
    m.date <- as.integer(str_sub(date, 6, 7))
    
    diff.in.months <- 12L * (y.refdate - y.date) + m.refdate - m.date
  }
  
  if (since) return(diff.in.months) else return(-diff.in.months)
  
}