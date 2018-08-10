library(httr)
library(jsonlite)

# Initialize Environment
# URL = https://api.iextrading.com/1.0/stock/aapl/stats

ua <- user_agent("https://api.iextrading.com/1.0/stock/")
ua

# <-- Start Function -->
iex_api <- function(path) {
  
  #url <- modify_url("https://api.iextrading.com/", path = "1.0/stock/aapl/stats")
  url <- modify_url("https://api.iextrading.com/", path = "1.0/stock/aapl/chart/5y")
  
  resp <- GET(url, ua)
  
  if(http_type(resp) !="application/json") {
    stop("API did not return a json file", call. = FALSE)
  } # Get the url
  
  parsed <- jsonlite::fromJSON(content(resp, "text"), simplifyVector = TRUE)
  
  if(status_code(resp) !=200) {
    stop(
      sprintf(
        "IEX API Request Failed [%s]\n%s\n<%s>",
        status_code(resp),
        parsed$message,
        parsed$documentation_url
      ),
      call. = FALSE
    )
  }
  
  structure(
    list(
      content = parsed, 
      path = "1.0/stock/aapl/chart/5y", 
      response = resp
      ),
    class = "iex_api"
    )
}

# <-- Print the Result -->
print.iex_api <- function(x, ...) {
  cat("<IEX", x$path, ">n", sep = "")
  str(x$content)
  invisible(x)
}

View(parsed)

new <- data.frame(parsed)

summary(parsed)

xtable::xtable(parsed)
