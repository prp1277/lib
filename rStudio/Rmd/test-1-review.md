---
title: "test-1-review"
date: "2018-06-20"
---

# CONCEPT REVIEW SHEET FOR EXAM 1

Symbols:
```
    √ - Square Root
    σ - Standard Deviation
    μ - Mean - 
    x̅ - X-bar - mean of all values
    β - Beta
    R2 - Residual Sum of Squares
```
## Linear Functions & Linear Regression

Simple linear regression: 
```
Y = β + β(X)
    Y = Dependent
    X = Independent
    β = Unknown
```
## The Concept behind Linear Regression

### Dependent variable - left side of equation
 * _What_ we're measuring - i.e. Return

### Independent variable - right side of equation
 * What we're measuring _~by_ - i.e. Risk

### 2 Tests for Linear Regression
  1. The function should be a straight line
  2. If the exponent on X is anything other than 1, it's not linear

## Regression Models & Overfitting

If you can arrive at the same conclusion using 5 results instead of 5 million, using 5 million is overkill.

## Total Sum of Squares (TSS), Regression Sum of Squares (RSS), Error Sum of Squares (ESS)

```    
Total Sum of Squares = Regression SS + Residual SS
100% = Known + Error
``` 

## Relationship between TSS, RSS, ESS and R2

"Think of this as a pie chart"

The goal is to maximize RSS and minimize ESS

Since `TSS = RSS + ESS` 

TSS - Total Sum of Squares = RSS + ESS
  * The lower the residual sum of squares the higher the R2 and the greater the explanatory power

## Computing TSS, RR and ESS for simple numeric values

**Total Sum of Squares** `TSS = [ RSS + ESS ]`

**Coefficient of Determination** `R2 = [ RSS + ESS ]`

## Relationship between correlation and R2

R2 - coefficient of determination - ranges from 0 - 1

R2 is the proportion of the variance in the dependent variable that is predictable from the independent variable.

R2 measures how well observed outcomes are replicated by the model based on the proportion of total variation of outcomes explained by the model.

## Fitting and Interpreting Linear Regression Values

The better the linear regression fits the data in comparison to the simple average, the closer the value of R2 is to 1.
```
R2 = 1 - [ RSS / TSS ]
```
## Training Sets versus Test Sets



## Root Mean Squared Error (RMSE) and Mean Absolute Error (MAE)

MAE = | Sample Error |

## Computing RMSE and MAE for simple numeric values

RMSE measures the predictive ability of the model
```
RMSE = √RSS
```
## Properties of the Normal Distribution

![Normal Distribution](https://en.wikipedia.org/wiki/Normal_distribution#/media/File:Empirical_Rule.png)

## Z Distribution and Standardizing Values

Answers the question "How far away is it from its' mean?"
```
Z = [ x̅ - μ ] / σ
```
Z values range from -3.4 to 3.4

## Central Limit Theorem

Given a sufficient sample size from a population with a finite level of variance, the mean of all samples from the same population will approximately equal the mean of the entire population.

As n gets larger, the distribution of of the difference between the sample average and its limit approximates the normal distribution with a mean of 0 and a variance of σ2.

> The distribution approaches normality regardless of the shape of the distribution of the individual

## Impact of sample standard deviation and sample size on SE

Standard Error - measures the average error resulting from using a sample to estimate the mean.

`Standard Error = [ σ (sample) / √n ]`

Standard Error is the Sample Standard Deviation divided by the sqrt of the sample size

1. As n increases the standard error decreases
2. As sample standard deviation increases, standard error decreases

## Differences between Normal versus Sampling Distribution

> The average of the averages

## Computing Confidence Intervals for various levels of significance (90%, 95%, 99%, etc.)
>[Excel File](https://hawksrockhurst-my.sharepoint.com:443/:x:/g/personal/powellpr_hawks_rockhurst_edu/Ee1lYpOdtMxMhOrWMMCpn7EBWoBYP_xr9ypzqvV5WUPSWg?email=prp1277%40gmail.com&e=kd6gQq)
1. The higher the confidence levels (all else equal), the greater the Margin of Error
2. As sample size increases, Margin of Error decreases

x̅ + 2(s) = 95%

x̅ - 2(s) = 95%

## Multiple Linear Regression



## Interpreting Confidence Intervals in a Multiple Linear Regression

Range from 0 - 1 (0 - 100%)



## Categorical (Dummy) Variables

Make sure to bin things correctly.



## Interpreting Dummy Variables

