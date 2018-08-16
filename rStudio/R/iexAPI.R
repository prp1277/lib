library(httr)
library(tidyverse)

ua <- user_agent("https://api.iextrading.com")
url <- modify_url(url = "https://api.iextrading.com", path = "/1.0/ref-data/symbols")


target <- function(path) {
  resp <- GET(url, ua)
  if(http_type(resp) !="application/json"){
    stop("API did not return json format", call. = FALSE)
  }
  parsed <- jsonlite::fromJSON(content(resp, "text"), simplifyDataFrame = TRUE)
  
  if(status_code(resp) !=200){
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
      path ="/1.0/ref-data/symbols", 
      response = resp
      ),
    class="target"
    )
}

print.target <- function(x, ...){
  cat("<IEX", x$path, ">n", sep = "")
  str(x$content)
  invisible(x)
}

View(parsed)

new <- data.frame(parsed)

summary(parsed)

xtable::xtable(parsed)

enabled <- ifelse(parsed$isEnabled ==TRUE, 1, 0)

symbols <- (parsed$symbol)
symbols

write.csv(parsed, file = file.choose(new = T))
