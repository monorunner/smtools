# ======== SM TOOLS ======== #
# ======== 0000-00-00 ======== #
# -------- shane.mono -------- #


library(data.table)
library(stringr)
library(XLConnect)
library(stringdist)

#' Summarise data columns.
#'
#' @param dt Data table to summarise.
#' @param writepath Path + file name to save xlsx.
#' @return Data summary.
#' @export
summarise <- function(dt, writepath = str_c(getwd(), "/Data Summary.xlsx"), append = "Data Summary") {
  
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
                EmptyRate = lapply(.SD, function(x) if(class(x) == "character") sum(x == "")/.N),
                Min = lapply(.SD, min, na.rm = TRUE),
                Max = lapply(.SD, max, na.rm = TRUE))]
  
  cols <- out[uniqueN <= 5 & uniqueN > 2, ColName]
  
  out[, str_c("Level", 1:5) := ""]
  
  
  out[ColName %in% cols,
      str_c("Level", 1:5) :=
        data.table(t(dt[, lapply(.SD, function(x) c(unique(x), rep("", 5 - uniqueN(x)))),
                        .SDcols = (cols)]))]
  
  out <- cbind(out[, 1:7], out[, lapply(.SD, function(x) lapply(x, as.character)), .SDcols = 8:14])
  out <- out[, lapply(.SD, unlist)]
  # out[Class == "numeric", `:=`(Min = as.numeric(Min), Max = as.numeric(Max))]
  
  if(!is.null (writepath) ) {
    cat("Writing Excel...\n")
    if (append == "Data Summary") 
      file.copy(from = "D:/Users/shane.mono/00. Random/Snippets/Data Summary Template.xlsx",
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


cp2 <- function(x, row.names = FALSE, excel.names = TRUE) {
  if(excel.names)
    write.table(excel.names(x), "clipboard", row.names = row.names, sep = "\t") else
      write.table(x, "clipboard-16384", row.names = row.names, sep = "\t")
  
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
convert.dates <- function(dt, additional.cols = NULL, format = "%Y-%m-%d") {
  
  if(!is.null(additional.cols)) datecols <- c(names(dt)[names(dt) %like% "date"], additional.cols) else
    datecols <- names(dt)[names(dt) %like% "date"]
  
  cat(length(datecols), " date columns found.\n")
  
  for(i in 1:length(datecols)) {
    dt[, (datecols[i]) := as.Date(get(datecols[i]), format = format)]
    
  }
  return(dt)
  
}



#' Generate sript headers.
#' @param title Will convert to upper letters automatically.
#' @export
gen.headers <- function(title = "WRITE STH NICE") {
  str_c("# ======== ", str_to_upper(title), " ======== #\n# ======== ", 
        format(Sys.time(), "%Y-%m-%d"),
        "======== #\n# -------- shane.mono -------- #") %>% 
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


#' Pivot table inspired grouping?
#' 
#' @param col Vector; use in \code{data.table}.
#' @param strgroup Grouping string in the format of \code{"Other+Unknown=Other"}.
#' Supports fuzzy match. \code{"O+U=Other"}, if Other and Unknown are the only categories
#' starting with O and U. Supports \code{"NA+U=Other"}. For blanks, use \code{"$+O=Other"}.
#' @param suffix Suffix of created grouped column.
#' @param new Create a new column of the grouped categories. Default to \code{FALSE}.
#' @details No debugging / error catching.
#' @value A lookup table with unique values from \code{col} and grouped values.
#' @example billing[, pgroup(cust.group, "O+U=Other")] # group Other and Unknown into Other
#' billing[, cust.group2 := pgroup(cust.group, "$+O+U+NA=Other", new = TRUE)]
#' @export
pgroup <- function(col, strgroup, suffix = ".grp", new = FALSE) {
  
  colnm <- deparse(substitute(col))
  col0 <- copy(col)
  col <- unique(col)
  col2 <- col
  vec <- unlist(str_split(strgroup, ","))
  
  for(i in 1:length(vec)) {
    ele <- unlist(str_split(vec[i], "\\+|="))
    if (any(ele=="NA")) col2[is.na(col)] <- ele[length(ele)]
    children <- unlist(lapply(ele[-length(ele)], function(x) col[col %like% str_c("^", x)]))
    col2[col2 %in% children] <- ele[length(ele)]
    
  }
  
  out <- data.table(col, col2)
  setnames(out, c(colnm, str_c(colnm, suffix)))
  
  
  if(new) {
    for(i in 1:length(col)) col0[col0 == col[i]] <- col2[i]
    return(col0)
  }
  
  return(out)
  
}


#' Examine whether one column in a data table is unique by another
#' 
#' @param x A list of column names or one column name without quote; the key which other columns should be unique by.
#' @param ... Other columns to investigate uniqueness.
#' @details No debugging / error catching.
#' @value Returns a vector with the same length as \code{...} with values \code{TRUE} or \code{FALSE}.
#' @example dt[, ifunique(id, name, status)]
#' dt[, ifunique(list(tour.code, date), tour.name, adult.price)]
#' @export
ifunique <- function(x, ...) {
  
  if(class(x) == "list") n <- length(x) else {x <- list(x); n <- 1}
  y <- list(...)
  out <- c()
  for(i in 1:length(y)) {
    z <- y[[i]]
    dt <- data.table(do.call(cbind, x), z)
    dupe <- dt[, .N, by = c(paste0("V", 1:n), "z")][, .N, by = c(paste0("V", 1:n))][N > 1]
    if(nrow(dupe) == 0) res <- TRUE else res <- FALSE
    out <- c(out, res)
  }
  return(out)
  
  
}


fuzzymatch <- function(x, y, n = 3, ...) {
  
  x <- unique(x)
  y <- unique(y)
  match.matrix <- stringdistmatrix(x, y, ...)
  rank.matrix <- t(apply(match.matrix, 1, frank, ties.method = "random"))
  for (i in 1:n) {
    match.index <- apply(rank.matrix, 1, function(x) which(x == i)) %>% unlist
    match.res <- y[match.index]
    if (i == 1) out <- match.res else out <- cbind(out, match.res)
  }
  out <- data.table(x, out)
  setnames(out, c("raw", str_c("match", 1:n)))
  return(out)
  
}


set.diff <- function(x, y) {
  
  xnm <- deparse(substitute(x))
  ynm <- deparse(substitute(y))
  x <- unique(x)
  y <- unique(y)
  
  out <- data.table(xnm = c(length(x), length(intersect(x, y)), length(setdiff(x, y)), 0),
                    ynm = c(length(y), length(intersect(x, y)), 0, length(setdiff(y, x))))
  setnames(out, c(xnm, ynm))
  out <- cbind(value = c("length", "common", "in x only", "in y only"), out)
  return(out)
  
  
}



