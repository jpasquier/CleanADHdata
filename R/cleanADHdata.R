#       _____ __                  ___    ___   __ __    __       __
#      / ___// /___  ___ _ ___   / _ |  / _ \ / // /___/ /___ _ / /_ ___ _
#     / /__ / // -_)/ _ `// _ \ / __ | / // // _  // _  // _ `// __// _ `/
#     \___//_/ \__/ \_,_//_//_//_/ |_|/____//_//_/ \_,_/ \_,_/ \__/ \_,_/


# -------------------------------- Libraries -------------------------------- #

# Install missing packages and load required libraries
pkgs <- c("readxl", "writexl")
npkgs <- pkgs[!(pkgs %in% installed.packages()[, "Package"])]
if(length(npkgs)) {
  install.packages(npkgs)
}
for (pkg in pkgs) {
  library(pkg, character.only = TRUE)
}
if (Sys.info()["sysname"] != "Windows") {
  if (!("tcltk" %in% installed.packages()[, "Package"])) {
    install.packages("tcltk")
  }
  library(tcltk)
}
rm(pkgs, npkgs, pkg)

# ---------------------------- Help function(s) ----------------------------- #

# bool2index: create a string from a boolean vector which indicates where are
# the positive cases
# Example : c(TRUE, TRUE, TRUE, FALSE, FALSE, TRUE) -> "1-3, 6"
bool2index <- function(b) {
  if (sum(b) == 1) {
    r <- which(b)
  } else {
    r <- cbind(
      c(1[b[1]], (2:length(b))[b[-1] & !b[-length(b)]]),
      c((1:(length(b) - 1))[b[-length(b)] & !b[-1]], length(b)[b[length(b)]])
    )
    r <- apply(r, 1, function(x) {
      if (x[1] == x[2]) x[1] else paste0(x[1], "-", x[2])
    })
    r <- paste(r, collapse = ", ")
  }
  return(r)
}

# Extract dates and times from a string
str2dateHour <- function(x) {
  if (!any(class(x) == "character")) {
    stop("A string must be provided as input")
  }
  fmts <- c("%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y")
  ptns <- c("^[0-9]{4}(-[0-9]{2}){2}$", "^([0-9]{2}/){2}[0-9]{4}$",
            "^([0-9]{2}\\.){2}[0-9]{4}$")
  fmts <- c(fmts, paste(fmts, "%H:%M"))
  ptns <- c(ptns, sub("\\$$", " [0-9]{2}:[0-9]{2}(:[0-9]{2})$", ptns))
  b <- sapply(ptns, function(ptn) all(is.na(x) | grepl(ptn, x)))
  b <- b & sapply(fmts, function(fmt) {
    r <- as.POSIXct(x, format = fmt)
    all(is.na(x) == is.na(r))
  })
  fmts[b]
  if (sum(b) != 1) {
    stop(paste("The dates contained in this character string vector are",
               "not recognized."))
  }
  r <- as.POSIXct(x, format = fmts[b])
  d <- as.Date(format(r, "%Y-%m-%d"))
  h <- format(r, "%H:%M")
  if (all(h[!is.na(h)] == "00:00")) h[!is.na(h)] <- "03:00"
  list(date = d, hour = h)
}

# ----------------------- Import auxiliary data file ------------------------ #

# Select the auxiliary data file
auxiliaryDataFile <- ""
cap <- "Select the auxiliary data file"
if (!file.exists(auxiliaryDataFile)) {
  if (Sys.info()["sysname"] == "Windows") {
    auxiliaryDataFile <- choose.files(caption = cap, multi = FALSE)
  } else {
    auxiliaryDataFile <- tk_choose.files(caption = cap, multi = FALSE)
    auxiliaryDataFile <- paste(auxiliaryDataFile, collapse = " ")
  }
}
rm(cap)

# Define working directory (this is the directory which contains the auxiliary
# data file
workingDirectory <- dirname(auxiliaryDataFile)

# Log files
now <- format(Sys.time(), "_%Y%m%d_%H%M%S")
now <- ""
logDir <- file.path(workingDirectory, "logs")
if (!dir.exists(logDir)) dir.create(logDir)
errLog <- file.path(logDir, paste0("errors", now, ".log"))
wrnLog <- file.path(logDir, paste0("warnings", now, ".log"))
file.create(errLog)
file.create(wrnLog)

# Check that all the requested tables are included in the auxiliary data file
requestedTables <- c("EMInfo", "Regimen")
optionalTables <- c("AddedOpenings", "AdverseEvents", "EMCovariables",
                    "NonMonitoredPeriods", "PatientCovariables")
auxiliaryDataTables <- excel_sheets(auxiliaryDataFile)
b <- requestedTables %in% auxiliaryDataTables
if (any(!b)) {
  cat(file = errLog, append = TRUE, paste0(
    "Missing table(s): ", paste(requestedTables[!b], collapse = ", "), "\n"
  ))
  stop(paste0("Fatal error: missing tables (see ", errLog, ")"))
}
b <- optionalTables %in% auxiliaryDataTables
if (any(!b)) {
  cat(file = wrnLog, append = TRUE, paste0(
    "The following optional table(s) were not found in the auxiliary data ",
    "file: ", paste(optionalTables[!b], collapse = ", "), "\n"
  ))
}
rm(b)

# Import auxiliary data
tabList <- auxiliaryDataTables[
  auxiliaryDataTables %in% c(requestedTables, optionalTables)]
for (tabName in tabList) {
  tab <- read_excel(auxiliaryDataFile, tabName, na = c("", "NA"))
  # If an optional table contains no data, it is deleted without warning
  if (tabName %in% optionalTables & nrow(tab) == 0) {
    tabList <- tabList[tabList != tabName]
  } else {
    tab <- as.data.frame(tab)
    # The function readxl::read_excel imports date as POSIXct variables.
    # We convert them to date variable
    for (j in which(sapply(tab, function(x) any(class(x) %in% "POSIXct")))) {
      if (all(format(tab[[j]], "%H%M%S") == "000000" | is.na(tab[[j]]))) {
        tab[[j]] <- as.Date(tab[[j]])
      }
    }
    assign(tabName, tab)
  }
}
rm(auxiliaryDataFile, auxiliaryDataTables, requestedTables,
   optionalTables, tabName, tab, j)

# ------------------------------ Id variables ------------------------------- #

# Tables are linked together by the variables PatientCode and Monitor
idvar <- c("PatientCode", "Monitor")

# ---------------- Check tables of the auxiliary data files ----------------- #

# Check that all the required variables are present and that there is no
# missing values. If a table contains a StartDate and an EndDate, check that
# StartDate <= EndDate.
requestedVariables <- list(
  AddedOpenings = c(idvar, "Date", "AddedOpenings"),
  AdverseEvents = c("PatientCode", "Date", "AdverseEvent",
                    "AdverseEventGrade"),
  EMCovariables = c(idvar, "StartDate", "EndDate"),
  EMInfo = c(idvar, "StartDate", "EndDate"),
  NonMonitoredPeriods = c(idvar, "StartDate", "EndDate"),
  PatientCovariables = c("PatientCode", "StartDate", "EndDate"),
  Regimen = c(idvar, "StartDate", "EndDate", "On", "Off", "ExpectedOpenings")
)
for (tabName in tabList) {
  V <- requestedVariables[[tabName]]
  tmpTab <- get(tabName)
  replaceTab <- FALSE
  # Missing variables
  b <- V %in% names(tmpTab)
  if (any(!b)) {
    cat(file = errLog, append = TRUE, paste0(
      "Missing variable(s) in ", tabName, ": ",
      paste(V[!b], collapse = ", "), "\n"
    ))
    tmpTab[V[!b]] <- NA
    replaceTab <- TRUE
  }
  # Check that variables containing dates are in Date format
  for (v in grep("^(Start|End)?Date$", names(tmpTab), value = TRUE)) {
    if (class(tmpTab[[v]])[1] != "Date") {
      cat(file = errLog, append = TRUE, paste0(
        tabName, ": Variable ", v, " is not in date format"
      ))
      if (class(tmpTab[[v]])[1] == "character") {
        linesBadDate <- which(!grepl("^[0-9]+$", tmpTab[[v]]))
        if (length(linesBadDate) > 0) {
          cat(file = errLog, append = TRUE, paste0(
            " (badly formatted dates are probably on the lines ",
            paste(linesBadDate, collapse = ", "), ")"
          ))
        }
      }
      cat(file = errLog, append = TRUE, "\n")
      tmpTab[[v]] <- as.Date(NA)
      replaceTab <- TRUE
    }
  }
  # Missing values (some variables are allowed to contain missing values)
  if (tabName == "AdverseEvents") {
    V2 <- V[V != "AdverseEventGrade"]
  } else if (tabName == "Regimen") {
    V2 <- V[!(V %in% c("On", "Off"))]
  } else if (tabName %in% c("EMCovariables", "PatientCovariables")) {
    V2 <- V[!(V %in% c("StartDate", "EndDate"))]
  } else {
    V2 <- V
  }
  b <- apply(is.na(tmpTab[V2]), 1, any)
  if (any(b)) {
    cat(file = errLog, append = TRUE, paste0(
      tabName, ": Missing value(s) in line(s) ", bool2index(b), "\n"
    ))
    tmpTab <- tmpTab[!b, V]
    replaceTab <- TRUE
  }
  # StartDate > EndDate
  b <- FALSE
  if ("StartDate" %in% V & "EndDate" %in% V) {
    b <- !is.na(tmpTab$StartDate) & !is.na(tmpTab$EndDate) &
      tmpTab$EndDate < tmpTab$StartDate
    if (any(b)) {
      for(i in 1:sum(b)) {
        cat(file = errLog, append = TRUE, paste(
          tabName, ": EndDate < StartDate for PatientCode",
          tmpTab[b, "PatientCode"][i], "Monitor",
          tmpTab[b, "Monitor"][i], "\n"
        ))
      }
      tmpTab <- tmpTab[!b, ]
      replaceTab <- TRUE
    }
  }
  if (replaceTab) assign(tabName, tmpTab)
}
suppressWarnings(rm(
  b, i, linesBadDate, r, replaceTab, requestedVariables, tabName, tmpTab, V, v
))

# --------------------- Event lists and daily adherence --------------------- #

# Select the data directory. T
dataDirectory <- ""
cap <- "Select the data directory"
if (!dir.exists(dataDirectory)) {
  if (Sys.info()["sysname"] == "Windows") {
    dataDirectory <- choose.dir(caption = cap)
  } else {
    dataDirectory <- tk_choose.dir(caption = cap)
    dataDirectory <- paste(dataDirectory, collapse = " ")
  }
}
rm(cap)
# Select all the files in working directory whose name contains the character
# string `Event(s)list` or `Daily adherence` (type case does not have to match)
fileList <- list.files(dataDirectory)
fileList <- lapply(c("events?list", "dailyadherence"), function(p) {
  fileList[grepl(p, tolower(fileList))]
})

if (all(sapply(fileList, length) == 0)) {
  stop('The data directory do not contains any data')
}

# Imports eventlists and daily adherence files. The files are read and the
# tables they contain are concatenated.
OpeningTables <- lapply(c(events = 1, dailyAdh = 2), function(k) {
  do.call(rbind, lapply(fileList[[k]], function(f) {
    f0 <- file.path(dataDirectory, f)
    if (grepl("\\.xlsx?", f)) {
      r <- as.data.frame(read_excel(f0, na = c("", "NA")))
    } else if (grepl("\\.csv", f)) {
      r <- if (grepl(";", readLines(f0, n = 1))) {
        #read.csv2(f0)
        read.table(text = readLines(f0, warn = FALSE), header = TRUE, sep = ";")
      } else {
        #read.csv(f0)
        read.table(text = readLines(f0, warn = FALSE), header = TRUE, sep = ",")
      }
    } else {
      cat(file = errLog, append = TRUE, paste0(
        "File", f, " is not readable\n"
      ))
      return(NULL)
    }
    # Check that the files is not empty
    if (nrow(r) == 0) {
      cat(file = wrnLog, append = TRUE, paste("File", f, "is empty\n"))
    }
    # Rename variables if necessary (for compatibility with medamigo)
    R <- c(MonitoringId = "Monitor", NbOpenings = "RecordedOpenings")
    for (u in names(R)) {
      if (any(names(r) == u) & all(names(r) != R[u])) {
        names(r)[names(r) == u] <- R[u]
      }
    }
    # Check that the necessary variables are present
    V <- c("PatientCode", "Monitor", "Date")
    if (k == 2) V <- c(V, "RecordedOpenings")
    if (!all(V %in% names(r))) {
      for (v in V) {
        if (all(names(r) != v)) {
          cat(file = errLog, append = TRUE, paste0(
            "Variable ", v, " is missing in file ", f, "\n"
          ))
          r[[v]] <- NA
        }
      }
    }
    # Select the necessary variables
    r <- r[V]
    # Check that the variables do not contain missing values
    for (v in V) {
      b <- is.na(r[[v]])
      if (all(b)) {
        cat(file = errLog, append = TRUE, paste(
          "The variable", v, "in the file", f, "is empty.\n"
        ))
      } else if (any(b)) {
        cat(file = errLog, append = TRUE, paste0(
          "The variable ", v, " in the file ", f, " contains missing values ",
          "on rows ", bool2index(b), ".\n"
        ))
      }
    }
    # Check that the variable `Date` is a date or a datetime
    if (any(!is.na(r$Date))) {
      error <- FALSE
      if (any(class(r$Date) %in% c("Date", "POSIXct"))) {
        r$Date <- as.character(r$Date)
      }
      z <- tryCatch(str2dateHour(r$Date), error = function(e) e)
      if (any(grepl("error", class(z)))) {
        cat(file = errLog, append = TRUE, paste(
          "The date variable in the file", f, "is incorrectly formatted.",
          "The following formats are accepted: yyyy-mm-dd; dd.mm.yyyy;",
          "dd/mm/yyyy; yyyy-mm-dd HH:MM(:SS); dd.mm.yyyy HH:MM(:SS);",
          "dd/mm/yyyy HH:MM(:SS)\n"
        ))
        r$Date <- as.Date(NA)
        if (k == 1) r$Hour <- as.character(NA)
      } else {
        r$Date <- z$date
        if (k == 1) r$Hour <- z$hour
      }
    }
    # Daily adherence: check that the values of the variable `RecordedOpenings`
    # are non-negative integers
    if (k == 2) {
      ro <- r$RecordedOpenings
      if (any(!is.na(ro))) {
        if (any(class(ro) %in% c("character", "numeric", "integer"))) {
          error <- FALSE
          if (any(class(ro) == "character")) {
            b <- !(is.na(ro) | grepl("^[0-9]+$", ro))
            if (any(b)) {
              r$RecordedOpenings[b] <- NA
              error <- TRUE
            }
            r$RecordedOpenings <- as.numeric(r$RecordedOpenings)
          } else if (any(class(ro) %in% c("numeric", "integer"))) {
            b <- !(is.na(ro) | ro %% 1 == 0 & ro >= 0)
            if (any(b)) {
              r$RecordedOpenings[b] <- NA
              error <- TRUE
            }
          }
          if (error) {
            cat(file = errLog, append = TRUE, paste0(
              "The variable RecordedOpenings in the file ", f, " contains ",
              "values that are not non-negative integers on rows ",
              paste(which(b), collapse = ", "), ".\n"
            ))
          }
        } else {
         cat(file = errLog, append = TRUE, paste(
            "The variable RecordedOpenings in the file", f,
            "is not a numeric variable.\n"
          ))
        }
      }
    }
    # Add the file name to check if the same patient/monitor/date combination
    # is found in several event files
    r$FileName <- f
    return(r)
  }))
})

# Check if the same patient/monitor/date combination is found in several event
# files
dup1 <- unique(do.call(rbind, lapply(OpeningTables, function(tab) {
  tab[c("PatientCode", "Monitor", "Date", "FileName")]
})))
dup2 <- aggregate(FileName ~ PatientCode + Monitor + Date, dup1,
                  function(x) length(unique(x)))
dup2 <- dup2[dup2$FileName > 1, ]
if (nrow(dup2) > 0) {
  dup2$FileName <- NULL
  dup1 <- merge(dup1, dup2, by = c("PatientCode", "Monitor", "Date"))
  dup2 <- aggregate(FileName ~ PatientCode + Monitor + Date, dup1,
                    function(x) paste(x, collapse = ", "))
  for(i in 1:nrow(dup2)) {
    cat(file = wrnLog, append = TRUE, paste0(
      "PatientCode ", dup2[i, "PatientCode"], ", Monitor ", dup2[i, "Monitor"],
      ", Date ", dup2[i, "Date"], ": Events found in multiple files: ",
      dup2[i, "FileName"], "\n"
    ))
  }
}
suppressWarnings(rm(dup1, dup2, i))

# Remove FileName variable
OpeningTables <- lapply(OpeningTables, function(tab) {
 tab$FileName <- NULL
 return(tab)
})

# ---------- Eventlist: Number of openings by patient/monitor/date ---------- #

events <- OpeningTables$events

# Attribution of the opening to the previous day if it takes place between
# 00:00 and 02:59
a <- strptime("00:00", format = "%H:%M")
b <- strptime("03:00", format = "%H:%M")
x <- strptime(events$Hour, format = "%H:%M")
i <- x >= a & x < b
events$Date[i] <- events$Date[i] - 1
rm(a, b, i, x)

# Number of openings by patient / monitor / date
events <- aggregate(Hour ~ PatientCode + Monitor + Date, events, length)
names(events)[names(events) == "Hour"] <- "RecordedOpenings"

# ------------------- Daily adherence: duplicate entries -------------------- #

dailyAdh <- OpeningTables$dailyAdh
b <- duplicated(dailyAdh[c("PatientCode", "Monitor", "Date")])
if (any(b)) {
  dup <- unique(dailyAdh[b, c("PatientCode", "Monitor")])
  for(i in 1:nrow(dup)) {
    cat(file = wrnLog, append = TRUE, paste0(
      "PatientCode ", dup[i, "PatientCode"], ", Monitor ", dup[i, "Monitor"],
      ": Some dates are duplicated, openings are added.\n"
    ))
  }
  dailyAdh <- aggregate(RecordedOpenings ~ PatientCode + Monitor + Date,
                        dailyAdh, sum)
}
suppressWarnings(rm(b, dup, i))

# ----------------------- Final implementation table ------------------------ #

# Concatenate tables
mems <- rbind(events, dailyAdh)
rm(events, dailyAdh)

# Check that the two tables do not overlap.
b <- duplicated(mems[c("PatientCode", "Monitor", "Date")])
if (any(b)) {
  dup <- unique(mems[b, c("PatientCode", "Monitor")])
  for(i in 1:nrow(dup)) {
    cat(file = wrnLog, append = TRUE, paste0(
      "PatientCode ", dup[i, "PatientCode"], ", Monitor ", dup[i, "Monitor"],
      ": Openings found in both Eventlist and Dailyadherence files. If ",
      "overlapping, openings are added.\n"
    ))
  }
  mems <- aggregate(RecordedOpenings ~ PatientCode + Monitor + Date, mems, sum)
}
suppressWarnings(rm(b, dup, i))

# Add a warning if no data is found for a patient/monitor
da <- merge(EMInfo[c(idvar)], cbind(unique(mems[idvar]), DataAvailable = 1),
            by = idvar, all.x = TRUE)
da$DataAvailable[is.na(da$DataAvailable)] <- 0
if (any(da$DataAvailable == 0)) {
  da <- da[da$DataAvailable == 0, ]
  for(i in 1:nrow(da)) {
    cat(file = wrnLog, append = TRUE, paste0(
      "PatientCode ", da[i, "PatientCode"], ", Monitor ", da[i, "Monitor"],
      ": Present in the auxiliary data file, but no event found\n"
    ))
  }
}
suppressWarnings(rm(da, i))

# ----------------------- Observation period (EMInfo) ----------------------- #

# Check that we have at most one one observation period per patient/monitor
obs <- EMInfo[c(idvar, c("StartDate", "EndDate"))]
dup <- unique(obs[duplicated(obs[idvar]), idvar])
if (nrow(dup) > 0) {
  for (i in 1:nrow(dup)) {
    cat(file = errLog, append = TRUE, paste(
      "EMInfo: More than one observation period are defined for PatientCode",
      dup[i, "PatientCode"], "Monitor", dup[i, "Monitor"], "\n"
    ))
  }
}
suppressWarnings(rm(dup, i))

# Convert periods to long format
obsL <- do.call(rbind, lapply(1:nrow(obs), function(i) {
  d <- seq(obs[i, "StartDate"], obs[i, "EndDate"], 1)
  data.frame(
    PatientCode = obs[i, "PatientCode"],
    Monitor = obs[i, "Monitor"],
    Date = d
  )
}))

# Keep only the dates which are within the observation period and add zeros
# at the which are not covered by the MEMS data
mems <- merge(obsL, mems, by = c(idvar, "Date"), all.x = TRUE)

# Define zero opening if there was no recording for a day included in the
# observation period
mems$RecordedOpenings[is.na(mems$RecordedOpenings)] <- 0
rm(obs, obsL)

# Define the number of corrected openings as the number of recorded openings
mems$CorrectedOpenings <- mems$RecordedOpenings

# ----------------------------- Added openings ------------------------------ #

                   # ------------------------------------ #
                   # Only if a table AddedOpenings exists #
                   # ------------------------------------ #

if (any(ls() == "AddedOpenings")) {

  # Check that the table contains at least one row
  if (nrow(AddedOpenings) > 0) {

    # Sum added openings per PatientCode/Monitor/Date
    ao <- aggregate(AddedOpenings ~ PatientCode + Monitor + Date,
                    AddedOpenings, sum)

    # Add "AddedOpenings" to MEMS table
    mems <- merge(mems, ao, by = c(idvar, "Date"), all.x = TRUE)
    mems$AddedOpenings[is.na(mems$AddedOpenings)] <- 0
    mems$CorrectedOpenings <- mems$CorrectedOpenings + mems$AddedOpenings
    rm(ao)

    # Invert columns AddedOpenings and CorrectedOpenings
    J <- 1:ncol(mems)
    j1 <- which(names(mems) == "AddedOpenings")
    j2 <- which(names(mems) == "CorrectedOpenings")
    if (j1 > j2) {
      J[j1] <- j2
      J[j2] <- j1
    }
    mems <- mems[J]
    rm(J, j1, j2)

  }

}

# -------------------------- Non monitored periods -------------------------- #

                 # ------------------------------------------ #
                 # Only if a table NonMonitoredPeriods exists #
                 # ------------------------------------------ #

if (any(ls() == "NonMonitoredPeriods")) {

  # Check that the table contains at least one row
  if (nrow(NonMonitoredPeriods) > 0) {

    # Convert periods to long format
    nmp <- NonMonitoredPeriods[c(idvar, c("StartDate", "EndDate"))]
    nmpL <- do.call(rbind, lapply(1:nrow(nmp), function(i) {
      d <- seq(nmp[i, "StartDate"], nmp[i, "EndDate"], 1)
      data.frame(
        PatientCode = nmp[i, "PatientCode"],
        Monitor = nmp[i, "Monitor"],
        Date = d,
        NonMonitored = TRUE
      )
    }))
    rm(nmp)

    # Warn about overlaps if any
    ol <- aggregate(NonMonitored ~ PatientCode + Monitor + Date, nmpL, length)
    b <- ol$NonMonitored > 1
    if (any(b)) {
      for(i in 1:sum(b)) {
        cat(file = wrnLog, append = TRUE, paste(
          "NonMonitoredPeriods: Overlaps for PatientCode",
          ol[b, "PatientCode"][i], "Monitor", ol[b, "Monitor"][i],
          "Date", ol[b, "Date"][i], "\n"
        ))
      }
      nmpL <- unique(nmpL)
    }
    rm(ol, b)
    if (any(ls() == "i")) rm(i)

    # Add NonMonitored variable to the MEMS dataset
    mems <- merge(mems, nmpL, by = c(idvar, "Date"), all.x = TRUE)
    mems$NonMonitored[is.na(mems$NonMonitored)] <- FALSE
    rm(nmpL)

  }

}

# --------------------------------- Regimen --------------------------------- #

if (nrow(Regimen) >= 1) {

  rgn <- Regimen

  # Check that cyclic regimens are well defined
  b1 <- any(names(rgn) == "On")
  b2 <- any(names(rgn) == "Off")
  if (!all(b1, b2)) {
    if (xor(b1, b2)) {
      cat(file = errLog, append = TRUE, paste(
        "Regimen: There is a variable", c("'On'", "'Off'")[c(b1, b2)],
        "but no variable", c("'Off'.", "'On'.")[c(b1, b2)],
        "Check rgn table\n"
      ))
    }
    rgn$On <- NA
    rgn$Off <- NA
  }
  rgn <- rgn[c(idvar, "StartDate", "EndDate", "On", "Off", "ExpectedOpenings")]
  b <- is.na(rgn$On) != is.na(rgn$Off)
  if (any(b)) {
    for(i in 1:sum(b)) {
      cat(file = errLog, append = TRUE, paste(
        "Regimen: Missing 'On' or 'Off' for PatientCode",
        rgn[b, "PatientCode"][i], "Monitor", rgn[b, "Monitor"][i],
        "StartDate", rgn[b, "StartDate"][i], "\n"
      ))
    }
    rgn[b, "On"] <- NA
    rgn[b, "Off"] <- NA
  }
  b <- is.na(rgn$On) & is.na(rgn$Off)
  rgn[b, "On"] <- 1
  rgn[b, "Off"] <- 0
  suppressWarnings(rm(b, b1, b2, i))

  # Convert periods to long format
  rgnL <- do.call(rbind, lapply(1:nrow(rgn), function(i) {
    d <- seq(rgn[i, "StartDate"], rgn[i, "EndDate"], 1)
    e <- c(rep(rgn[i, "ExpectedOpenings"], rgn[i, "On"]),
           rep(0, rgn[i, "Off"]))
    e <- rep(e, ceiling(length(d) / length(e)))[1:length(d)]
    data.frame(
      PatientCode = rgn[i, "PatientCode"],
      Monitor = rgn[i, "Monitor"],
      Date = d,
      ExpectedOpenings = e
    )
  }))
  rm(rgn)

  # Check that there are no overlaps
  ol <- aggregate(ExpectedOpenings ~ PatientCode + Monitor + Date,
                  rgnL, length)
  b <- ol$ExpectedOpenings > 1
  if (any(b)) {
    for(i in 1:sum(b)) {
      cat(file = errLog, append = TRUE, paste(
        "Regimen: Overlaps for PatientCode", ol[b, "PatientCode"][i],
        "Monitor", ol[b, "Monitor"][i], "Date", ol[b, "Date"][i], "\n"
      ))
    }
    ol <- ol[b, names(ol) != "ExpectedOpenings"]
    ol$RegimenOverlaps <- TRUE
    rgnL <- merge(rgnL, ol, by = c("PatientCode", "Monitor", "Date"),
                  all.x = TRUE)
    rgnL[is.na(rgnL$RegimenOverlaps), "RegimenOverlaps"] <- FALSE
    rgnL[rgnL$RegimenOverlaps, "ExpectedOpenings"] <- NA
    rgnL <- unique(rgnL)
  }
  suppressWarnings(rm(b, i, ol))

  # Add ExpectedOpenings variable to the MEMS dataset
  mems <- merge(mems, rgnL, by = c(idvar, "Date"), all.x = TRUE)
  rm(rgnL)

} else {

  mems$ExpectedOpenings <- NA

}

# Invert columns NonMonitored and ExpectedOpenings
if ("NonMonitored" %in% names(mems)) {
  J <- 1:ncol(mems)
  j1 <- which(names(mems) == "ExpectedOpenings")
  j2 <- which(names(mems) == "NonMonitored")
  if (j1 > j2) {
    J[j1] <- j2
    J[j2] <- j1
  }
  mems <- mems[J]
  rm(J, j1, j2)
}

# Check that ExpectedOpenings is defined for each Patient/Monitor/Date in
# monitored periods
b <- is.na(mems$ExpectedOpenings)
if (all(b)) {
  cat(file = errLog, append = TRUE, paste(
    "The variable ExpectedOpenings is not defined. Check previous errors.\n"
  ))
} else {
  b2 <- aggregate(b, mems[c("PatientCode", "Monitor")], all)
  if (any(b2[[3]])) {
    b2 <- b2[b2[[3]], 1:2]
    for (i in nrow(b2)) {
      cat(file = errLog, append = TRUE, paste(
        "All ExpectedOpenings are missing for PatientCode",
        b2[i, "PatientCode"], "Monitor", b2[i, "Monitor"],
        "(possibly due to previous errors)\n"
      ))
      b[mems$PatientCode == b2[i, "PatientCode"] &
          mems$Monitor == b2[i, "Monitor"]] <- FALSE
    }
  }
  if ("NonMonitored" %in% names(mems)) {
    b <- b & !mems$NonMonitored
  }
  if (any(b)) {
    for(i in 1:sum(b)) {
      cat(file = errLog, append = TRUE, paste(
        "Missing ExpectedOpenings for PatientCode", mems[b, "PatientCode"][i],
        "Monitor", mems[b, "Monitor"][i], "Date", mems[b, "Date"][i],
        "(missing values can be caused by overlaps)\n"
      ))
    }
  }
}
suppressWarnings(rm(b, b2, i))

# ----------------------- Implementation calculation ------------------------ #

# By monitor
mems$Implementation <- mems$CorrectedOpenings >= mems$ExpectedOpenings
mems$Implementation <- as.numeric(mems$Implementation)

# Set Implementation = NA for non-monitored periods (if any)
if ("NonMonitored" %in% names(mems)) {
  mems$Implementation[mems$NonMonitored] <- NA
}

# By patient
imp <- aggregate(Implementation ~ PatientCode + Date, mems, prod,
                 na.action = NULL)
mon <- aggregate(Monitor ~ PatientCode + Date, mems,
                 function(z) length(unique(z)))
names(mon)[names(mon) == "Monitor"] <- "MonitorsNb"
pmems <- merge(mon, imp, by = c("PatientCode", "Date"))
rm(imp, mon)

# ----------------------------- Relative dates ------------------------------ #

# Sort MEMS dataset by PatientCode/Monitor/Date
mems <- mems[order(mems$PatientCode, mems$Monitor, mems$Date), ]
pmems <- pmems[order(pmems$PatientCode, pmems$Date), ]

# Check that dates are consecutive for each PatientCode/Monitor
b <- aggregate(Date ~ PatientCode + Monitor, mems, function(x) {
  length(x) > 1 & any(x[-1] - x[-length(x)] != 1)
})
if (any(b[[3]])) {
  for(i in 1:sum(b[[3]])) {
    cat(file = errLog, append = TRUE, paste(
      "Dates are not consecutive for PatientCode", b[b[[3]], "PatientCode"][i],
      "Monitor", b[b[[3]], "Monitor"][i], "\n"
    ))
  }
}
b <- aggregate(Date ~ PatientCode, pmems, function(x) {
  length(x) > 1 & any(x[-1] - x[-length(x)] != 1)
})
if (any(b[[2]])) {
  for(i in 1:sum(b[[2]])) {
    cat(file = errLog, append = TRUE, paste0(
      "Dates are not consecutive for PatientCode ",
      b[b[[2]], "PatientCode"][i], "\n"
    ))
  }
}
rm(b)

# Add relative dates
mems$RelativeDate <-
  with(mems, ave(Implementation, PatientCode, Monitor, FUN = seq_along))
pmems$RelativeDate <-
  with(pmems, ave(Implementation, PatientCode, FUN = seq_along))

# ----------------------------- Adverse Events ------------------------------ #

                   # ------------------------------------ #
                   # Only if a table AdverseEvents exists #
                   # ------------------------------------ #

if (any(ls() == "AdverseEvents")) {

  # Concatenate adverse events by patient and dates
  ae <- AdverseEvents
  ae$AdverseEvents <- with(AdverseEvents, {
    ifelse(is.na(AdverseEventGrade), AdverseEvent,
           paste0(AdverseEvent, " (", AdverseEventGrade, ")"))
  })
  ae <- aggregate(AdverseEvents ~ PatientCode + Date, ae, paste,
                  collapse = ", ")

  # Merge adverse events with MEMS datasets
  mems <- merge(mems, ae, by = c("PatientCode", "Date"), all.x = TRUE)
  pmems <- merge(pmems, ae, by = c("PatientCode", "Date"), all.x = TRUE)
  rm(ae)

}

# --------------------------- Monitor covariables --------------------------- #

                   # ------------------------------------ #
                   # Only if a table EMCovariables exists #
                   # ------------------------------------ #

if (any(ls() == "EMCovariables")) {

  mcov <- EMCovariables

  # Keep only patients/monitors who are in the MEMS dataframe
  mcov <- merge(unique(mems[idvar]), mcov, by = idvar)

  # Fill empty dates
  b <- is.na(mcov$StartDate)
  for (i in which(b)) {
    d <- min(mems[mems$PatientCode == mcov[i, "PatientCode"] &
                    mems$Monitor == mcov[i, "Monitor"], "Date"])
    if (!is.na(mcov[i, "EndDate"])) {
      d <- min(d, mcov[i, "EndDate"])
    }
    mcov[i, "StartDate"] <- d
  }
  b <- is.na(mcov$EndDate)
  for (i in which(b)) {
    d <- max(mems[mems$PatientCode == mcov[i, "PatientCode"] &
                    mems$Monitor == mcov[i, "Monitor"], "Date"])
    if (!is.na(mcov[i, "StartDate"])) {
      d <- max(d, mcov[i, "StartDate"])
    }
    mcov[i, "EndDate"] <- d
  }
  suppressWarnings(rm(b, d, i))

  # Convert periods to long format
  v <- names(mcov)[!(names(mcov) %in% c("StartDate", "EndDate"))]
  mcovL <- do.call(rbind, lapply(1:nrow(mcov), function(i) {
    d <- seq(mcov[i, "StartDate"], mcov[i, "EndDate"], 1)
    cbind(mcov[i, v, drop = FALSE], Date = d, row.names = NULL)
  }))
  rm(mcov, v)

  # Check that there are no overlaps
  ol <- aggregate(.count ~ PatientCode + Monitor + Date,
                  cbind(mcovL, .count = 1), sum)
  b <- ol$.count > 1
  if (any(b)) {
    for(i in 1:sum(b)) {
      cat(file = errLog, append = TRUE, paste(
        "EMCovariables: Overlaps for PatientCode", ol[b, "PatientCode"][i],
        "Monitor", ol[b, "Monitor"][i], "Date", ol[b, "Date"][i], "\n"
      ))
    }
    mcovL <- merge(mcovL, ol, by = c(idvar, "Date"))
    mcovL <- mcovL[mcovL$.count == 1, ]
    mcovL$.count <- NULL
  }
  suppressWarnings(rm(b, i, ol))

  # Add EM covariables to the MEMS dataset
  mems <- merge(mems, mcovL, by = c(idvar, "Date"), all.x = TRUE)
  rm(mcovL)

}

# --------------------------- Patient covariables --------------------------- #

                 # ----------------------------------------- #
                 # Only if a table PatientCovariables exists #
                 # ----------------------------------------- #

if (any(ls() == "PatientCovariables")) {

  pcov <- PatientCovariables

  # Keep only patients who are in the MEMS dataframe
  pcov <- pcov[pcov$PatientCode %in% unique(pmems$PatientCode), ]

  # Fill empty dates
  b <- is.na(pcov$StartDate)
  for (i in which(b)) {
    d <- min(pmems[pmems$PatientCode == pcov[i, "PatientCode"], "Date"])
    if (!is.na(pcov[i, "EndDate"])) {
      d <- min(d, pcov[i, "EndDate"])
    }
    pcov[i, "StartDate"] <- d
  }
  b <- is.na(pcov$EndDate)
  for (i in which(b)) {
    d <- max(pmems[pmems$PatientCode == pcov[i, "PatientCode"], "Date"])
    if (!is.na(pcov[i, "StartDate"])) {
      d <- max(d, pcov[i, "StartDate"])
    }
    pcov[i, "EndDate"] <- d
  }
  suppressWarnings(rm(b, d, i))

  # Convert periods to long format
  v <- names(pcov)[!(names(pcov) %in% c("StartDate", "EndDate"))]
  pcovL <- do.call(rbind, lapply(1:nrow(pcov), function(i) {
    d <- seq(pcov[i, "StartDate"], pcov[i, "EndDate"], 1)
    cbind(pcov[i, v, drop = FALSE], Date = d, row.names = NULL)
  }))
  rm(pcov, v)

  # Check that there are no overlaps
  ol <- aggregate(.count ~ PatientCode + Date, cbind(pcovL, .count = 1), sum)
  b <- ol$.count > 1
  if (any(b)) {
    for(i in 1:sum(b)) {
      cat(file = errLog, append = TRUE, paste(
        "PatientCovariables: Overlaps for PatientCode",
        ol[b, "PatientCode"][i], "Date", ol[b, "Date"][i], "\n"
      ))
    }
    pcovL <- merge(pcovL, ol, by = c("PatientCode", "Date"))
    pcovL <- pcovL[pcovL$.count == 1, ]
    pcovL$.count <- NULL
  }
  suppressWarnings(rm(b, i, ol))

  # Add covariables
  pmems <- merge(pmems, pcovL, by = c("PatientCode", "Date"), all.x = TRUE)
  rm(pcovL)

}

# ----------------------------- Sort dataframes ----------------------------- #

mems <- mems[order(mems$PatientCode, mems$Monitor, mems$Date), ]
pmems <- pmems[order(pmems$PatientCode, pmems$Date), ]

# --------------------------------- Summary --------------------------------- #

# Summary by monitor and by patient
smy.mems <- aggregate(Implementation ~ PatientCode + Monitor, mems, mean,
                      na.rm = TRUE)
smy.pmems <- aggregate(Implementation ~ PatientCode, pmems, mean, na.rm = TRUE)

# ----------------------------- Export results ------------------------------ #

L <- list(
  `by monitor` = mems,
  `by patient` = pmems,
  `summary by monitor` = smy.mems,
  `summary by patient` = smy.pmems
)
outputFile <- file.path(workingDirectory,
                        paste0("implementation", now, ".xlsx"))
write_xlsx(L, outputFile)
rm(L)
