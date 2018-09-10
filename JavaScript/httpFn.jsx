var request = require("request");

var options = {
  method: "GET",
  url: "https://api.iextrading.com/1.0/stock/market/batch",
  qs: {
    last: "5",
    range: "1m",
    types: "peers,news,quote,chart",
    symbols: "MSFT,GOOG,FB,AMZN"
  },
  headers: { "accept-encoding": "application/json" }
};

request(options, function(error, response, body) {
  if (error) throw new Error(error);

  console.log(body);
});
