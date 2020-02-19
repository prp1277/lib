attach(medexpense_data)
library(psych)
options(scipen = 999)
describe(medexpense_data)
# Pay attention to the means and sd of age and medical expenses

# Issues in MLRL continued 
# Issue 7 - Determining Variable Importance Using Beta Regressions
Model <- lm(medical_expenses~
              age + bmi )
summary(Model)

# Med_exp = -6726 + 265(age) + 301(bmi)
#   no conclusions can be made by looking at the size of the coefficients because
#   there is not an apples to apples measure (age[yrs] vs bmi[amt])

AgeModel <- lm(medical_expenses~
              age )
summary(AgeModel)

# Increase age by 1 standard deviation
# Finding Correllation by age
# Med_exp = 1983 + 280(14.15)
#    5945 = 1983 + 3962
#    3962 = 3962
#    [3962 / 11986] = [3962 / 11986]
# Med_exp = 0 + .33(age)

# The .33 tells us for every 1 sd increase in age, 
#   medical expenses increases by .33 std units

#<-- Create a Beta Regression -->
# Need to get rid of yes and no
scaled_data <- scale(medexpense_data) # Results in error

numeric_data <-data.frame(medical_expenses, age, bmi) 
View(numeric_data)

z_data <- data.frame(scale(numeric_data))
View(z_data)

# Z = [Xbar - mu] / sd
# [19 - 39.62] / 14.15
# Z = -1.4564

attach(z_data)
beta_model <- lm(medical_expenses~
                   age + bmi, data = z_data)
summary(beta_model)

#<-- Interpreting Results -->
# Med_exp = 0 + .31(age) + .15(bmi)

# Age and bmi are measures of medical expenses expressed in standard deviation
# For every 1 sd unit increase in age medical expenses increase
#   by .31 sd units and for every 1 sd increase in bmi, medical expenses
#   increase by .15 sd unit

# CANNOT DO BETA REGRESSION ON DUMMY VARIABLES
#   Why? Gender increasing by 1 sd unit does not make any sense


# Issue 8 - Multicollinearity
# Run a regression that measures weight by left / right shoe size

attach(shoe_sizes_data)

weight_model <- lm(weight~
                     right_shoe_size + left_shoe_size)

summary(weight_model)

# Weight = 79 + 10 Right shoe size
# Weight = 83 + 10 Left Shoes size
# Weight = 79 + 14.26(right_shoe_size) + 4.14(left_shoe_size)
# This doesn't make any sense - they are coolinear (complemantary)
#   which will confuse how the independent variables affect in dependent variables

# Prove the results are garbage by finding the correllation
cor(right_shoe_size, left_shoe_size)

#<-- Explaining Multi-collinearity -->
# Multi-collinearity refers to two variables that have a 
#  correllation greater than .90
# If the correllation is highly correlated the regression
#  results can be nonsensical and one of the variables needs to be dropped
