setwd("XLS")
library("lmtest")
library("readxl")

excel = read_excel("for_regr.xlsx")
new_data = data.frame(excel)

modelR = lm(data=new_data, log(y)~log(x))
bptest(modelR) # no heteroscedascisity
