#Directorio donde estan los datos
setwd('/home/guillermo/Documentos/Medidas algas')

install.packages('xlsx')
library(xlsx)

#Función para calcular la moda
Mode <- function(x) {
  ux <- unique(x)
  ux[which.max(tabulate(match(x, ux)))]
}

#Función para calcular outliers
Outliers <- function(x) {
  Q3=quantile(x,0.75)
  Q1=quantile(x,0.25)
  IQR=Q3-Q1  # interquartile
  outSup=Q3+IQR*1.5
  outInf=Q1-IQR*1.5
  outliers = which(c(x < outInf | x > outSup))
  
  outliers
}

wb <- loadWorkbook("p1 nanno 9_9 1300 alls 1_X.xlsx")
sheets <- getSheets(wb)
sheet <- sheets[[1]]

df <- readColumns(sheet, startColumn = 1, endColumn = 353, startRow = 11, endRow = 108)

#Carga de los valores para la frecuencia 534
A534 <- df[2:13, 121]
B534 <- df[14:25, 121]
C534 <- df[26:37, 121]
D534 <- df[38:49, 121]
E534 <- df[50:61, 121]
F534 <- df[62:73, 121]
G534 <- df[74:85, 121]
H534 <- df[86:97, 121]

#Valores para primer blanco (534)
Blanco1 <- c(A534[1],B534[1],C534[1],D534[1],E534[1],F534[1],G534[1],H534[1])
#Valores para control (534)
Control <- c(A534[2],A534[3],B534[2],B534[3],C534[2],C534[3],D534[2],D534[3],E534[2],E534[3],F534[2],F534[3],G534[2],G534[3],H534[2],H534[3])
#Valores para SW (534)
SW <- c(A534[4],A534[5],B534[4],B534[5],C534[4],C534[5],D534[4],D534[5],E534[4],E534[5],F534[4],F534[5],G534[4],G534[5],H534[4],H534[5])
#Valores para PRD (534)
PRD <- c(A534[6],A534[7],B534[6],B534[7],C534[6],C534[7],D534[6],D534[7],E534[6],E534[7],F534[6],F534[7],G534[6],G534[7],H534[6],H534[7])
#Valores para MLH (534)
MLH <- c(A534[8],A534[9],B534[8],B534[9],C534[8],C534[9],D534[8],D534[9],E534[8],E534[8],F534[9],F534[8],G534[9],G534[8],H534[9],H534[8])
#Valores para Cal (534)
Cal <- c(A534[10],A534[11],B534[10],B534[11],C534[10],C534[11],D534[10],D534[11],E534[10],E534[11],F534[10],F534[11],G534[10],G534[11],H534[10],H534[11])
#Valores para segundo blanco (534)
Blanco2 <- c(A534[12],B534[12],C534[12],D534[12],E534[12],F534[12],G534[12],H534[12])



#Calculo de la media, mediana, moda, máximo y mínimo para el campo 1 (Blanco)
m <- mean(Blanco1)
me <- median(Blanco1)
mo <- Mode(Blanco1)
maxValue <- max(Blanco1)
minValue <- min(Blanco1)
sprintf("Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)

#detección de outlier
ou <- Outliers(Blanco1)

#Recalcular valores sin Outliers
if (length(ou)>0){
  m <-  mean(Blanco1[-ou])
  me <-  median(Blanco1[-ou])
  mo <-  Mode(Blanco1[-ou])
  maxValue <- max(Blanco1[-ou])
  minValue <- min(Blanco1[-ou])
  sprintf("WITHOUT OUTLIERS: Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)
} else {
  sprintf("No outliers")
}

#Calculo de la media, mediana, moda, máximo y mínimo para el campo 2 y 3 (Control)
m <- mean(Control)
me <- median(Control)
mo <- Mode(Control)
maxValue <- max(Control)
minValue <- min(Control)
sprintf("Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)

#detección de outlier
ou <- Outliers(Control)

#Recalcular valores sin Outliers
if (length(ou)>0){
  m <-  mean(Control[-ou])
  me <-  median(Control[-ou])
  mo <-  Mode(Control[-ou])
  maxValue <- max(Control[-ou])
  minValue <- min(Control[-ou])
  sprintf("WITHOUT OUTLIERS: Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)
} else {
  sprintf("No outliers")
}

#Calculo de la media, mediana, moda, máximo y mínimo para el campo 4 y 5 (SW)
m <- mean(SW)
me <- median(SW)
mo <- Mode(SW)
maxValue <- max(SW)
minValue <- min(SW)
sprintf("Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)

#detección de outlier
ou <- Outliers(SW)

#Recalcular valores sin Outliers
if (length(ou)>0){
  m <-  mean(SW[-ou])
  me <-  median(SW[-ou])
  mo <-  Mode(SW[-ou])
  maxValue <- max(SW[-ou])
  minValue <- min(SW[-ou])
  sprintf("WITHOUT OUTLIERS: Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)
} else {
  sprintf("No outliers")
}

#Calculo de la media, mediana, moda, máximo y mínimo para el campo 6 y 7 (PRD)
m <- mean(PRD)
me <- median(PRD)
mo <- Mode(PRD)
maxValue <- max(PRD)
minValue <- min(PRD)
sprintf("Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)

#detección de outlier
ou <- Outliers(PRD)

#Recalcular valores sin Outliers
if (length(ou)>0){
  m <-  mean(PRD[-ou])
  me <-  median(PRD[-ou])
  mo <-  Mode(PRD[-ou])
  maxValue <- max(PRD[-ou])
  minValue <- min(PRD[-ou])
  sprintf("WITHOUT OUTLIERS: Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)
} else {
  sprintf("No outliers")
}

#Calculo de la media, mediana, moda, máximo y mínimo para el campo 8 y 9 (MLH)
m <- mean(MLH)
me <- median(MLH)
mo <- Mode(MLH)
maxValue <- max(MLH)
minValue <- min(MLH)
sprintf("Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)

#detección de outlier
ou <- Outliers(MLH)

#Recalcular valores sin Outliers
if (length(ou)>0){
  m <-  mean(MLH[-ou])
  me <-  median(MLH[-ou])
  mo <-  Mode(MLH[-ou])
  maxValue <- max(MLH[-ou])
  minValue <- min(MLH[-ou])
  sprintf("WITHOUT OUTLIERS: Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)
} else {
  sprintf("No outliers")
}

#Calculo de la media, mediana, moda, máximo y mínimo para el campo 10 y 11 (Cal)
m <- mean(Cal)
me <- median(Cal)
mo <- Mode(Cal)
maxValue <- max(Cal)
minValue <- min(Cal)
sprintf("Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)

#detección de outlier
ou <- Outliers(Cal)

#Recalcular valores sin Outliers
if (length(ou)>0){
  m <-  mean(Cal[-ou])
  me <-  median(Cal[-ou])
  mo <-  Mode(Cal[-ou])
  maxValue <- max(Cal[-ou])
  minValue <- min(Cal[-ou])
  sprintf("WITHOUT OUTLIERS: Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)
} else {
  sprintf("No outliers")
}

#Calculo de la media, mediana, moda, máximo y mínimo para el campo 12 (Blanco2)
m <- mean(Blanco2)
me <- median(Blanco2)
mo <- Mode(Blanco2)
maxValue <- max(Blanco2)
minValue <- min(Blanco2)
sprintf("Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)

#detección de outlier
ou <- Outliers(Blanco2)

#Recalcular valores sin Outliers
if (length(ou)>0){
  m <-  mean(Blanco2[-ou])
  me <-  median(Blanco2[-ou])
  mo <-  Mode(Blanco2[-ou])
  maxValue <- max(Blanco2[-ou])
  minValue <- min(Blanco2[-ou])
  sprintf("WITHOUT OUTLIERS: Mean: %f, Median: %f, Mode: %f, Max: %f, Min: %f", m, me, mo, maxValue, minValue)
} else {
  sprintf("No outliers")
}