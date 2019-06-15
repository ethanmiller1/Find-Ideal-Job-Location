# Get Ideal Job Location 
**Version 1.0.0** 

> An Excel spreadsheet that determines what job locations would be ideal for you.

This speadsheet was designed to help you determine where to apply for jobs based on its proximity to existing placed of interest to you (family, church, etc.). It uses Google's Distance Matrix API to calculate the driving time between job sites and each place of interest.

## Build

1. Create an Excel Macro-Enabled Worksheet (Distance Calculator.xlsm)
1. `Developer` > `Visual Basic` > Right-click `VBAProject (Distance Calculator.xlms)` > `Insert` > `Module`
1. Copy and paste the code from `functions.vb` into the created `Module1` module.

Now you should have a function called GetDuration() in your Excel workbook that takes 3 parameters:

1. Starting Location
1. Destination
1. Your API Key

Example usage in cell:
``` excel
=GetDuration("Richardson","1501 N Country Club Rd, Garland, TX 75040",KEY)
```

### Get your free Google Distance Matrix API Key

1. Head down to [Google Cloud Platform](https://console.cloud.google.com)
1. Click "New Project" and then open your newly created project.
1. Open the hamburger menu and select `APIs & Services` > `Library` > `View All` (Under "Maps")
1. Select `Distance Matrix API` and click `Enable` on the following page.
1. Open the hamburger menu again and select `APIs & Services` > `Credentials`
1. Click on `Create credentials` > `API Key`

Note: you can only make 1 request to the web API per day without setting up a billing account. Google's API will still be free to use up to 100,000 requests per day, and there will be no autocharge after free trial ends.

### Activate your Key

1. On [Google Cloud Platform](https://console.cloud.google.com), click `Activate` on the top right.
1. Fill in your credit card information, and voila, your FREE API key is now activated.

## Usage

1. Determine which cities are ideal to apply for jobs in.

This step is to give you an idea of which cities you want as you're browsing job openings on indeed.com.

![](https://github.com/king-melchizedek/Find-Ideal-Job-Location/raw/master/demos/cities.png)

(Note: [traveltimeplatform.com](https://app.traveltimeplatform.com/search/0_lng=-97.04434&0_color=%23eb9f22&0_mode=driving&0_title=East%20Avenue%20J%2C%20Grand%20Prairie%2C%20TX%2C%20USA&0_lat=32.76720&1_lng=-96.61390&1_color=%2339d2e2&1_mode=driving&1_title=Northlake%20Baptist%20Church%2C%20Garland%2C%20TX%2C%20USA&1_lat=32.92730&2_lat=33.04136&2_lng=-96.56308&2_title=1310%20Leeward%20Ln%2C%20Wylie%2C%20TX%2C%20USA&2_mode=driving "I dare you to click me.") can help you eyeball a custom list of cities to test within this spreadsheet.)

2. Determine which companies take location priority.

Once you've found a potential company, enter it's location into the "Jobs" spreadsheet to evaluate it.

![](https://github.com/king-melchizedek/Find-Ideal-Job-Location/raw/master/demos/companies.png)

3. Give your locations and key a defined name in Excel.

![](https://github.com/king-melchizedek/Find-Ideal-Job-Location/raw/master/demos/metrics.png)

## How it works

Our function takes in 3 parameters and concatenates them to generate a URL that will return a JSON object. Assuming the following usage:

``` excel
=GetDuration("Richardson","Northlake Baptist Church",AIzaSyB06tMOemrwtlMcat2ZKLPYGdFJ--BNJ7c)
```

Our function would generate the following URL:

``` bash
https://maps.googleapis.com/maps/api/distancematrix/json?origins=richardson&destinations=northlake+baptist+church&mode=driving&language=en&key=AIzaSyB06tMOemrwtlMcat2ZKLPYGdFJ--BNJ7c
```

 Then it makes a server request to that URL, and grabs the following JSON object:

``` json
{
   "destination_addresses" : [ "1501 N Country Club Rd, Garland, TX 75040, USA" ],
   "origin_addresses" : [ "Richardson, TX, USA" ],
   "rows" : [
      {
         "elements" : [
            {
               "distance" : {
                  "text" : "12.2 km",
                  "value" : 12196
               },
               "duration" : {
                  "text" : "17 mins",
                  "value" : 1015
               },
               "status" : "OK"
            }
         ]
      }
   ],
   "status" : "OK"
}
```

Then it uses a regex command to find the keyword "duration", and searches the following text for the first occurrence of the keyword "value", isolating the numeric values directly following it into a capturing group, finally returning that group to the cell (in this case 1015).

### Error handling

If you use a key that has not been activated, it may return the following JSON upon visiting the URL:

``` json
{
   "destination_addresses" : [],
   "error_message" : "You have exceeded your daily request quota for this API. If you did not set a custom daily request quota, verify your project has an active billing account: http://g.co/dev/maps-no-account",
   "origin_addresses" : [],
   "rows" : [],
   "status" : "OVER_QUERY_LIMIT"
}
```

When that happens, our regex will fail to produce a value associated with the keyword "value" because it doesn't exist. It will simply return 0. Make sure you use an active key to prevent this issue.

If you use an invalid key, you will receive the following JSON and receive the same resulting 0:
``` json
{
   "destination_addresses" : [],
   "error_message" : "The provided API key is invalid.",
   "origin_addresses" : [],
   "rows" : [],
   "status" : "REQUEST_DENIED"
}
```