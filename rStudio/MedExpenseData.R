medexpense_data<-read.csv("C:/Users/prp12.000/OneDrive for Business/Courses/BIA-6309-Stats-And-Machine-Learning/06-06-2018/medexpense_data.csv")
attach(medexpense_data)
options(scipen=999)             # This changes the scientific notation we see
str(medexpense_data)
#
library(psych)
describe(medexpense_data)       # Get the descriptive statistics for the dataset
plot(age, medical_expenses)     # Medical expenses as a function of age

#Issue 1 - Interpreting confidence intervals in Multiple Linear Regression
regModel1<-lm(medical_expenses~age+bmi+children)
summary(regModel1)
# medical expenses = -7056 + 264 age + 300bmi + 390 children
# Medical expenses rise as age, bmi and number of kids increases 
# T-Value = coefficient / standard error = 10.498
confint(regModel1)

# Issue 2 - Dummy Variables - factoring variables
SMOKER<-ifelse(smoker=="yes",1,0)   # if smoker 1, nonsmoker = 0
SMOKER
GENDER<-ifelse(gender=="male",1,0)  # male = 1, female = 0
GENDER
regModel2<-lm(medical_expenses~age+bmi+children+smoker+gender)
summary(regModel2)              # R is able to define the strings as factors
                                # Adjusted R2 went up
# All else equal, the cost of being a smoker increases by $23812.57
# All else equal, the cost of being a male is $267.17 lower than being a female


