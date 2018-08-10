# Assignment 2 - 6/12/18
auto_data <- read.csv("C:/Users/prp12.000/OneDrive for Business/Courses/BIA-6309-Stats-And-Machine-Learning/06-06-2018/auto_data.csv")
attach(auto_data)
library(psych)
describe(auto_data)

# Regression 1 - mpg / horsepower 
regModel1<-lm(mpg ~ horsepower)
summary(regModel1)
anova(regModel1)
plot(regModel1)
plot(horsepower, mpg)
abline(regModel1)
# Comments:
# P values less than .05 are typically significant
# R2 value of .6049 also signals high correlation
# The coefficient estimates also tell us that for every
# 1 unit increase in hp, mpg decreases by .16 
# Predicted mpg = 39.94 - .1578(98{hp}) = 24.48 mpg

# Regression 2 - mpg / cylinders, displacement and hp
regModel2<-lm(mpg ~ cylinders+displacement+horsepower)
summary(regModel2)
anova(regModel2)
plot(regModel2)
# Comments:
# Looking at the negative coefficient data shows us
# that there is an inverse relationship between the 
# dependent and independent variables - i.e. the 
# mpg will shrink as hp, cylinders and displace increase

