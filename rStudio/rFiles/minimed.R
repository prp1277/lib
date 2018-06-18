setwd("C:/Users/prp12.000/OneDrive for Business/Courses/BIA-6309-Stats-And-Machine-Learning/06-06-2018")
attach(mini_medexpense_data)
library(psych)
describe(mini_medexpense_data)

#NEXT SECTION
set.seed(123)
RANDOMIZED_DATA<-mini_medexpense_data[order(runif(100)),] #Rows and columns notation - 100 rows, all columns
RANDOMIZED_DATA

TRAINING_SET<-RANDOMIZED_DATA[1:70,] #Picking the first 70 for training
TRAINING_SET

TEST_SET<-RANDOMIZED_DATA[71:100,] #Pick 71 - 100 for test 
TEST_SET
#NEXT

#REGRESSION MODEL
REG_MODEL<-lm(medical_expenses~age+bmi, data=TRAINING_SET) #Need to specify the data for the regression model
summary(REG_MODEL)

FITTED_VALUES<-fitted(REG_MODEL)
FITTED_VALUES

#RMSE -> install & load the new package("Metrics")
#RMSE - root means squared error - how accurate is the model?

#TRAINING SET
library(Metrics)
rmse(TRAINING_SET$medical_expenses, FITTED_VALUES) # Have to specify the training set bc the different sources
# On average they're screwing up a lot 

#TEST SET
PREDICT<-predict(REG_MODEL, TEST_SET)
rmse(TEST_SET$medical_expenses, PREDICT)
# Since we are only taking a broad look at this dataset the model is not great



