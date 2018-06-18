# Assignment 2 - 6/12/18
auto_data <- read.csv("C:/Users/prp12.000/OneDrive for Business/Courses/BIA-6309-Stats-And-Machine-Learning/06-06-2018/auto_data.csv")
attach(auto_data)
library(psych)
describe(auto_data)

# Perform linear regression - mpg dependent, hp as independent
