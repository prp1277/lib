#-- Data Transformation with dplyr --#
# [Source](http://r4ds.had.co.nz/transform.html)
library(nycflights13)
library(tidyverse)
flights

# Tibbles - dataframes tweaked to work better in tidyverse

# Data Types:
# 1. int - Integer                         # 5. lgl - logical (Boolean)
# 2. dbl - Doubles (real numbers)          # 6. fctr - factor categorical variables with fixed values
# 3. chr - Character, vector or string     # 7. date - dates 
# 4. dttm - Date / Time 

# Key dplyr functions
#     (filter()) - pick observation by value
#     (arrange()) - reorder rows
#     (select()) - pick variables by name
#     (mutate()) - create new variable with functions of existing variables
#     (summarize()) - collapse values into single summary
#     (group_by()) - changes scope of each function

# fn (
#    arg1 = dataframe, 
#    arg2 = action to take, 
#    arg3 = result
# ) Where & = "and", | = "or", ! = "not"

attach(flights)
filter(flights, month == 11 | month == 12)
nov_dec <- filter(flights, month %in% c(11,12))

# De Morgan's law:
# !(x & y) = !x | !y = !x & !y = !( x | y ) 

df <- tibble(x = c(1, NA, 3))
filter(df, x > 1)
filter(df, is.na(x) | x > 1)
# Filter only includes rows where condition = TRUE

(names(df))
names(flights)
str(flights)

# Find flights that were 2 + hours late
L <- (sched_arr_time - arr_time)
filter(flights, L > 120)
filter(flights, arr_delay > 120 )
L

# Using arrange()
arrange(flights, year, month, day)
