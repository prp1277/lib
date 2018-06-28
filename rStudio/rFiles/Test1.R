attach(sherwin_williams_data)
dim(sherwin_williams_data)
str(sherwin_williams_data)
summary(sherwin_williams_data)

# <-- Regression -->
options(scipen = 999)
RegTest <- lm(monthly_sales~
                price_per_gallon + ad_expenditures + average_income + growth_region)
summary(RegTest)

predict(RegTest)
resid(RegTest)

confint(RegTest)

priceIncrease <- (price_per_gallon + 2.00)
RegTestIncrease <- lm(monthly_sales~
                      priceIncrease + price_per_gallon + ad_expenditures +
                        average_income + growth_region)
summary(RegTestIncrease)

priceIncrease

245504.52-255905.78

adIncrease <- (ad_expenditures - 10000)
RegTestAdIncrease <- lm(monthly_sales~
                          price_per_gallon + ad_expenditures + adIncrease + average_income + growth_region)
summary(RegTestAdIncrease)

anova(RegTestAdIncrease)
anova(RegTestIncrease)
246111-245504

summary(RegTest)
summary(RegTestIncrease)
summary(RegTestAdIncrease)

sqrt(.9023)
(-5200*13)+(.06*140000)-(1.06*21500)
