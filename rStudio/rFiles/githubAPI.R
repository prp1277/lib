# Walking through API interaction using httr
# > [Source](https://github.com/r-lib/httr/blob/master/vignettes/quickstart.Rmd)
## Useful arguments:
### modify_url -      POST(modify_url("https://httpbin.org", path = "/post"))
### query arguments - POST("https://httpbin.org/post", query = list(foo = "bar"))
### headers -         POST("https://httpbin.org/post", add_headers(foo = "bar"))
### body as form -    POST("https://httpbin.org/post", body = list(foo = "bar"), encode = "form")
### body as json -    POST("https://httpbin.org/post", body = list(foo = "bar"), encode = "json")

library(httr)

ua <- user_agent("https://github.com/prp1277/lib")
ua

github_api <- function(path){
  url <- modify_url("https://api.github.com", path = path)
  
  # Import the response in json format or throw an error
  resp <- GET(url, ua)
  if (http_type(resp) !="application/json"){
    stop("API did not return json", call. = FALSE)
  }
  
  # Parse the response for json content. If there's a 200 error, tell why
  parsed <- jsonlite::fromJSON(content(resp, "text"), simplifyVector = FALSE)
  
  if (status_code(resp) != 200) {
    stop(
      sprintf(
        "GitHub API Request Failed [%s]\n%s\n<%s>",
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
      path = path,
      response = resp
    ),
    class = "github_api"
  )
} # End function(path)


print.github_api <- function(x, ...) {
  cat("<GitHub", x$path, ">\n", sep = "")
  str(x$content)
  invisible(x)
}

github_api("/users/prp1277")

rate_limit <- function() {
  req <- github_api("/rate_limit")
  core <- req$content$resources$core
  
  reset <- as.POSIXct(core$reset, origin = "1970-01-01")
  cat(core$remaining, " / ", core$limit,
      " (Resets at ", strftime(reset, "%H:%M:%S"), ")\n", sep = "")
}

rate_limit()