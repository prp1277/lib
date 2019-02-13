# [`Zip`](https://1drv.ms/u/s!AoLkNBOSNnKylLZXCrKCzPq8MAB3lg)` | `[`Video`](https://www.lynda.com/Ajax-tutorials/Parsing-JSON-data)

## Function

```js
$(document).ready(function(){
  $('#rep-lookup').submit(function(e){
    e.preventDefault();

    var $results = $('#rep-lookup-results'),
    zipcode = $('#txt-zip').val(),
    apiKey = 'kjasdfoiasdfhl2312asdfgwerfv';

    var requestURL = 'protocol.host.path.query?callback=?';
    // callback=? is crucial

    $.getJSON(requestURL, {
      'apiKey' : apiKey,
      'zipcode' : zipcode,
    }, function(data){
      console.log(data)
    });

    $results.html('Your reps using ' + zipcode + ' are:');
  })
})
```
