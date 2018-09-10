


```js
<!DOCTYPE html>
<html>
<head>
    <title>JSSample</title>
    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.0/jquery.min.js"></script>
</head>
<body>

<script type="text/javascript">
    $(function() {
        var params = {
            // Request parameters
            "apikey": "{string}",
            "city": "{string}",
            "companyname": "{string}",
            "ein": "{string}",
            "email": "{string}",
            "firstname": "{string}",
            "lastname": "{string}",
            "limit": "{number}",
            "offset": "{number}",
            "phone": "{string}",
            "postalcode": "{string}",
            "resourcetype": "Basic",
            "stateprovince": "{string}",
        };
      
        $.ajax({
            url: "https://api.infoconnect.com/v1/companies/?" + $.param(params),
            beforeSend: function(xhrObj){
                // Request headers
                xhrObj.setRequestHeader("Accept","application/json");
                xhrObj.setRequestHeader("apikey","{subscription key}");
            },
            type: "GET",
            // Request body
            data: "{body}",
        })
        .done(function(data) {
            alert("success");
        })
        .fail(function() {
            alert("error");
        });
    });
</script>
</body>
</html>


```