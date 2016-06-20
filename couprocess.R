# Process Guatemalan Environmentally ext Supply and Use tables.
# International Development Bank
# Prepared by Renato Vargas
# Overseen by Martin Cicowiez
# Team Leader: Onil Banerjee

# Preamble
# This prevents java from crashing due to memory constraints
# And we load important libraries.
options(java.parameters = "-Xmx3000m")
library("RPostgreSQL")
library(reshape)
library(xlsx)

# Function for java garbage collection
# From http://stackoverflow.com/questions/21937640
jgc <- function()
{
  .jcall("java/lang/System", method = "gc")
}    

# read in equivalence tables
# Columns
etc <- read.xlsx2("equivalence.xlsx", sheetIndex=1, header=TRUE, startRow=1, endRow=181, colIndex= c(1:11), stringsAsFactors=FALSE, colClasses= c(rep("character",9), "integer", "integer"))

jgc()

# Rows
etr <- read.xlsx2("equivalence.xlsx", sheetIndex=2, header=TRUE, startRow=1, endRow=353 , colIndex= c(1:15), stringsAsFactors=FALSE, colClasses= c(rep("character",12), "integer", "character", "integer"))

jgc()

# Function to replace NaN with zeros, from: 
# http://stackoverflow.com/questions/18142117
# This converts all missing values into zeros
# In this case it's OK
is.nan.data.frame <- function(x)
  do.call(cbind, lapply(x, is.nan))


# I. Supply Table
# ===============

# With package xlsx
st <- read.xlsx2("COU 2010-2012.xlsx", sheetIndex=1, header=FALSE, startRow=19, endRow=314, colIndex= c(1:165), stringsAsFactors=FALSE, colClasses= c("character","character", rep("numeric", 163)))

jgc()

# Use the is.nan function we created above
st[is.nan(st)] <- 0

# Give rows and columns correlative names that
# we can later match to specific classifications
# R = Rows (both tables are equal on this dimension)
# C = Columns (Tables differ in transactions.
# We'll fix that on the use table below.)

# Rename row and column names
a <- c(1:dim(st)[2])
colnames(st) <- sprintf("C%03d", a)
a <- c(1:dim(st)[1])
rownames(st) <- sprintf("R%03d", a)


# II. Use Table
# =============

# With package xlsx
ut <- read.xlsx2("COU 2010-2012.xlsx", sheetIndex=1, header=FALSE, startRow=328, endRow=623, colIndex= c(1:164), stringsAsFactors=FALSE, colClasses= c("character","character", rep("numeric", 162)))

jgc()

# Use function we created previously to replace NaN with zeros
ut[is.nan(ut)] <- 0

# Change row and column names
a <- c(1:dim(ut)[1])
rownames(ut) <- sprintf("R%03d", a)

a <- c(1:dim(ut)[2])
colnames(ut) <- sprintf("C%03d", a)

# And give different codes to columns that are different
# from the supply table.
a <- c(166:179)
colnames(ut)[c(151:164)] <- sprintf("C%03d", a)

# Do last minute replacements
# Just making sure that both tables have no duplicates
st$C001 <- ut$C001
st$C002 <- ut$C002
# Do some last minute replacements
st$C001[st$C001 == "Ajustes:"] <- "6701"
ut$C001[ut$C001 == "Ajustes:"] <- "6701"

# Extract relevant rows and columns in both tables
st 		<- st[c(nchar(st$C001) > 2)	, -c(130, 134, 148, 149, 150, 153, 155, 159, 164, 165)]
ut 		<- ut[c(nchar(ut$C001) > 2)	, -c(130, 134, 148, 149, 150, 153, 158, 159, 163,164)]


# III. Value Added Table
# =============

# With package xlsx
va <- read.xlsx2("COU 2010-2012.xlsx", sheetIndex=1, header=FALSE, startRow=631, endRow=653, colIndex= c(1:149), stringsAsFactors=FALSE, colClasses= c("character","character", rep("numeric", 147)))

jgc()

# Use function we created previously to replace NaN with zeros
va[is.nan(va)] <- 0

# Change row and column names

a <- c(297:(296+dim(va)[1]))
rownames(va) <- sprintf("R%03d", a)

a <- c(1:dim(va)[2])
colnames(va) <- sprintf("C%03d", a)

# Extract relevant rows and columns
va 		<- va[-c(1, 2, 5, 8, 9, 15)	, -c(130, 134, 148, 149)]


# IV. Employment Table
# ====================

# With package xlsx
em <- read.xlsx2("COU 2010-2012.xlsx", sheetIndex=1, header=FALSE, startRow=657, endRow=662, colIndex= c(1:147), stringsAsFactors=FALSE, colClasses= c("character","character", rep("numeric", 145)))

jgc()

# Use function we created previously to replace NaN with zeros
em[is.nan(em)] <- 0

# Change row and column names

a <- c(321:(320+dim(em)[1]))
rownames(em) <- sprintf("R%03d", a)

a <- c(1:dim(em)[2])
colnames(em) <- sprintf("C%03d", a)

# Extract relevant rows and columns
em 		<- em[						, -c(130, 134)]


# We make sure our Ad-hoc classification ends up in the Flat files.
st$C000 <- rownames(st)
ut$C000 <- rownames(ut)
va$C000 <- rownames(va)
em$C000 <- rownames(em)

# From supply and use tables
stF <- melt.data.frame(st, id.vars=c("C000", "C001", "C002"), na.rm=FALSE)
utF <- melt.data.frame(ut, id.vars=c("C000", "C001", "C002"), na.rm=FALSE)
vaF <- melt.data.frame(va, id.vars=c("C000", "C001", "C002"), na.rm=FALSE)
emF <- melt.data.frame(em, id.vars=c("C000", "C001", "C002"), na.rm=FALSE)

# Join supply and use tables
stF$supuse <- "01. Supply (Monetary)"
utF$supuse <- "02. Use (Monetary)"
vaF$supuse <- "03. Value Added (Monetary)"
emF$supuse <- "04. Employment"
sutF <- rbind(stF,utF, vaF, emF)

# Tidy up supply and use tables
sutF$"T" <- etc$"ntg.p.summ"[match(sutF$"variable", etc$"COL")]
sutF$"A" <- etc$"act.cod"[match(sutF$"variable", etc$"COL")]
sutF <- sutF[c("supuse", "C000", "T", "A", "value")]
colnames(sutF) <- c("supuse", "R", "T", "A", "value")


# V. SNA registered and SNA unregistered water
# =============================================

wt <- read.xlsx2("Base de datos O&U_2001-2010_v05.xlsx", sheetIndex=4, header=TRUE, startRow=1, endRow=1961, colIndex= c(1:17), stringsAsFactors=FALSE, colClasses= c("numeric", rep("character", 3),"numeric", rep("character", 11), "numeric" )) 

jgc()

# Extract relevant rows and columns
wt 		<- wt[ wt$A == 2010  		, -c(2, 3, c(5:11),13, 16)]


# VI. Crop Area, Irrigation, and Return Water
# ==========================================

agwt <- read.xlsx2("Base de datos agropecuaria_2001-2010_v03.xlsx", sheetIndex=4, header=TRUE, startRow=4, endRow=404, colIndex= c(1:27), stringsAsFactors=FALSE, colClasses= c( rep("character", 13),  rep("numeric", 14)))

jgc()

# Extract relevant rows and columns
agwt 	<- agwt[ agwt$year == 2010	, -c(c(1:7), 10, 12, 13, 22, 27)] 

# From water tables
agwtF <- melt.data.frame(agwt, id.vars=c("transaction", "act.cod", "C000", "year"), na.rm=FALSE)

# Tidy up registered/unregistered water table
wt <- wt[c("supuse", "C000","transaction", "col", "Dato")]
colnames(wt) <- c("supuse", "R", "T", "A", "value")


# Tidy up Crop Area, Irrigation, and Return Water table
agwtF[,c("variable")] <- sapply(agwtF[,c("variable")], as.character)
i <- unique(agwtF$variable)
agwtF$"variable"[agwtF$"variable" == i[1]]  <- "07. Cultivated Area (Ha)"
agwtF$"variable"[agwtF$"variable" == i[2]]  <- "08. Rainfed irrigation use (m3)"
agwtF$"variable"[agwtF$"variable" == i[3]]  <- "09. Sprinkler irrigation use (m3)"
agwtF$"variable"[agwtF$"variable" == i[4]]  <- "10. Drip irrigation use (m3)"
agwtF$"variable"[agwtF$"variable" == i[5]]  <- "11. Gravity use (m3)"
agwtF$"variable"[agwtF$"variable" == i[6]]  <- "12. Other use (m3)"
agwtF$"variable"[agwtF$"variable" == i[7]]  <- "13. All irrigation (m3)"
agwtF$"variable"[agwtF$"variable" == i[8]]  <- "14. Sprinkler irrigation return (m3)"
agwtF$"variable"[agwtF$"variable" == i[9]]  <- "15. Drip irrigation return (m3)"
agwtF$"variable"[agwtF$"variable" == i[10]] <- "16. Gravity return (m3)"
agwtF$"variable"[agwtF$"variable" == i[11]] <- "17. Other return (m3)"
agwtF <- agwtF[c("variable", "C000", "transaction", "act.cod", "value")]
colnames(agwtF) <- c("supuse", "R", "T", "A", "value")


# VII. Energy and Emissions
# =========================

# Energy table
eet <- read.xlsx2("BDCIEE.xlsx", sheetIndex=2, header=TRUE, startRow=7, endRow=31761, colIndex= c(2:38), stringsAsFactors=FALSE, colClasses= c("numeric", rep("character", 33),"numeric", "character", "character" ))

jgc()


# Equivalence
eetc <- read.xlsx2("equivalence.xlsx", sheetIndex=4, header=TRUE, startRow=1, endRow=126, colIndex= c(1:5), stringsAsFactors=FALSE, colClasses= c(rep("character",4)))

jgc()

eet <- eet[ eet$an==2010 , ] 

# Create consistent columns
eet$"T" <- eetc$"ntg.p.summ"[match(eet$"naet1", eetc$"naet1")]
eet$"A" <- eetc$"act.cod"[match(eet$"naet1", eetc$"naet1")]
eet$"R" <- etr$"R"[match(eet$"npt222", etr$"C001")]
eet$"supuse" = paste(eet$"cdr.scn", eet$"CódigoProducto", sep=" ")
colnames(eet)[35] <- "value"
eet <- eet[c("supuse", "R", "T", "A", "value")]

# tidy up individual table component names
i <- unique(eet$"supuse")
eet$"supuse"[eet$"supuse" == i[1]] <- "18. Energy supply (terajoule)"
eet$"supuse"[eet$"supuse" == i[4]] <- "19. Energy use (terajoule)"
eet$"supuse"[eet$"supuse" == i[5]] <- "20. Carbon Dioxide supply (CO2 tonnes)"
eet$"supuse"[eet$"supuse" == i[2]] <- "21. Nitrous Oxide supply (CO2 tonnes equivalent)"
eet$"supuse"[eet$"supuse" == i[3]] <- "22. Methane supply (CO2 tonnes equivalent)"


# VIII. Forest products
# =====================

# Forest table
ft <- read.xlsx2("BDCIB.xlsx", sheetIndex=1, header=TRUE, startRow=9, endRow=12679, colIndex= c(1:28), stringsAsFactors=FALSE, colClasses= c("character","numeric", rep("character", 23),"numeric", "character","character"))

jgc()

ft <- ft[ ft$"Año" == 2010 , ] 

# Create consistent columns
ft$"T" <- eetc$"ntg.p.summ"[match(ft$"NAET123", eetc$"cib.col")]
ft$"A" <- eetc$"act.cod"[match(ft$"NAET123", eetc$"cib.col")]
ft$"R" <- etr$"R"[match(ft$"RECODE.NPT", etr$"C001")]
ft$"supuse" = paste(ft$"Cuadro", ft$"DIM", sep=" ")
colnames(ft)[26] <- "value"
ft <- ft[c("supuse", "R", "T", "A", "value")]

# tidy up individual table components names
i <- unique(ft$"supuse")
ft$"supuse"[ft$"supuse" == i[1]] <- "23. Forest products supply (m3)"
ft$"supuse"[ft$"supuse" == i[2]] <- "24. Forest products use (m3)"
ft$"supuse"[ft$"supuse" == i[3]] <- "25. Animal species supply (number of individuals)"
ft$"supuse"[ft$"supuse" == i[4]] <- "26. Animal species use (number of individuals)"


# IX. Residuals
# =============

# Read in the data
ret <- as.data.frame(read.xlsx2("BDCIRE.xlsx", sheetIndex=1, header=TRUE, startRow=5, endRow=24195, colIndex= c(1:24), stringsAsFactors=FALSE, colClasses= c("character", "numeric" , rep("character", 6),  rep("numeric", 4), rep("character", 2),  "numeric", rep("character", 3), rep("numeric", 6) ) ))

jgc()

# Replace blanks with zeros
ret[is.nan(ret)] <- 0

# Keep the desired year alone
ret <- ret[ ret$"Año" == 2010 , ] 

# We extract product codes from a concatenated column for matching
ret$npt <- substring(ret$NPT227, 1, 4)

# Note: some codes don't match with our ad-hoc classif, so some extra
# replacements are needed. We've added a new equivalence table for this.

# Equivalence
reetc <- read.xlsx2("equivalence.xlsx", sheetIndex= "cirecol", header=TRUE, startRow=1, endRow=45, colIndex= c(1:4), stringsAsFactors=FALSE, colClasses= c(rep("character",4)))

jgc()

# Create consistent columns
ret$"T" <- reetc$"T"[match(ret$"NAET123", reetc$"CIRENAEG")]
ret$"A" <- reetc$"A"[match(ret$"NAET123", reetc$"CIRENAEG")]
ret$"R" <- etr$"R"[match(ret$"RECODE.NPT", etr$"C001")]
colnames(ret)[15] <- "value"
colnames(ret)[3] <- "supuse"

# Information domains
i <- unique(ret$"supuse")
ret$"supuse"[ret$"supuse" == i[1]] <- "27. Residuals supply (tonnes)"
ret$"supuse"[ret$"supuse" == i[2]] <- "28. Residuals use (tonnes)"

ret <- ret[c("supuse", "R", "T", "A", "value")]


# X. Subsoil resources
# ====================

sst <- as.data.frame(read.xlsx2("BDCIRS.xlsx", sheetIndex=1, header=TRUE, startRow=8, endRow=28691, colIndex= c(1:28), stringsAsFactors=FALSE, colClasses= c("character", "numeric", rep("character", 6),  rep("numeric", 9), rep("character", 7),  "numeric", rep("character", 3) ) ))

jgc()

# Replace blanks with zeros
sst[is.nan(sst)] <- 0

# Keep the desired year alone
sst <- sst[ sst$"Año" == 2010 , ] 

sst$C001 <- substring(sst$RECODE.NPT, 1, 6)

ssetc <- read.xlsx2("equivalence.xlsx", sheetIndex= "cirscol", header=TRUE, startRow=1, endRow=65, colIndex= c(1:4), stringsAsFactors=FALSE, colClasses= c(rep("character",4)))
jgc()

# Create consistent columns
sst$"T" <- ssetc$"T"[match(sst$"NAET123", ssetc$"CIRSNAEG")]
sst$"A" <- ssetc$"A"[match(sst$"NAET123", ssetc$"CIRSNAEG")]
sst$"R" <- etr$"R"[match(sst$"C001", etr$"C001")]
colnames(sst)[25] <- "value"
colnames(sst)[3] <- "supuse"

# Information domains
i <- unique(sst$"supuse")
sst$"supuse"[sst$"supuse" == i[1]] <- "29. Subsoil resource supply (tonnes)"
sst$"supuse"[sst$"supuse" == i[2]] <- "30. Subsoil resource use (tonnes)"

sst <- sst[c("supuse", "R", "T", "A", "value")]


# XI. Fisheries
# =============

# Connect to database
drv <- dbDriver("PostgreSQL")
con <- dbConnect(drv, dbname="naturacc_cuentas", 
                 host="212.83.58.14", port="5432", user="naturacc_onil", 
                 password="onilidb1234")

sys <- Sys.info()["sysname"]
if(sys["sysname"] == "Windows"){
  postgresqlpqExec(con, "SET client_encoding = 'windows-1252'");} 

# Query database
# We are interested in the output in tonnes for all 
# agricultural products.

# This query extracts the information that we need:
fsht <- dbGetQuery(con, "
                   SELECT
                   flujo as supuse,
                   scn.npg as npg,
                   npg336.producto as product,
                   scn.naeg as naeg,
                   naeg100.actividad as industry,
                   scn.ntg as ntg,
                   ntg20.trans as transaction,
                   scn.datofisico AS value
                   FROM scn
                   LEFT JOIN npg336
                   ON scn.npg = npg336.cod
                   LEFT JOIN ntg20
                   ON scn.ntg = ntg20.cod
                   LEFT JOIN naeg100
                   ON scn.naeg = naeg100.cod
                   WHERE
                   scn.ann = 2010
                   AND
                   (scn.npg BETWEEN 160100 AND 169900
                   OR
                   scn.npg BETWEEN 210100 AND 219900)
                   ORDER BY
                   scn.npg;");
dbDisconnect(con)
dbUnloadDriver(drv)
rm("con")
rm("drv")

# Equivalence
fshett <- read.xlsx2("equivalence.xlsx", sheetIndex= "scaebdt", header=TRUE, startRow=1, endRow=33, colIndex= c(1:3), stringsAsFactors=FALSE, colClasses= c("integer", rep("character",2)))
jgc()

fsheta <- read.xlsx2("equivalence.xlsx", sheetIndex= "scaebdt", header=TRUE, startRow=1, endRow=16, colIndex= c(5:7), stringsAsFactors=FALSE, colClasses= c("integer", rep("character",2)))
jgc()

fshetr <- read.xlsx2("equivalence.xlsx", sheetIndex= "scaebdr", header=TRUE, startRow=1, endRow=230, colIndex= c(1:5), stringsAsFactors=FALSE, colClasses= c("integer", "integer", "character", "integer", "character"))
jgc()


# Create consistent columns
fsht$"T" <- fshett$"T"[match(fsht$"ntg", fshett$"ntg")]
fsht$"A" <- fsheta$"A"[match(fsht$"naeg", fsheta$"naeg")]
fsht$"R" <- etr$"R"[match(fsht$"npg" , fshetr$"npg336")]

# Information domains
i <- unique(fsht$"supuse")
fsht$"supuse"[fsht$"supuse" == i[1]] <- "31. Fishery supply (tonnes)"
fsht$"supuse"[fsht$"supuse" == i[2]] <- "32. Fishery use (tonnes)"

fsht <- fsht[c("supuse", "R", "T", "A", "value")]


# XIII. Reaggregate according to new classification
# =================================================

# Bind all tables
eF <- rbind(sutF,wt,agwtF, eet, ft, ret, sst, fsht)


# Derive new row classification
eF$"RCONCAT" <- etr$"RCONCAT"[match(eF$"R", etr$"R")]

#Derive new column classification
eF <- within(eF, iabd.cod <- paste(T, A, sep=''))
eF$"CCONCAT" <- etc$"CCONCAT"[match(eF$"iabd.cod", etc$"iadb.cod")] 

# Clean up
eF <- eF [c("supuse", "R", "RCONCAT", "T", "A", "CCONCAT", "value")]
colnames(eF)[1] <- "DOMAIN"


# Write out the Excel

write.xlsx2(eF, "cou-e.xlsx", sheetName="cou-e-database", col.names=TRUE, row.names=FALSE, append=FALSE)


library("openxlsx")
openXL("cou-e.xlsx")

