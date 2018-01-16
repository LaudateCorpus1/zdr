## About
This script is used to manage rotations within Zendesk.

This script builds off the work done here: https://support.zendesk.com/hc/en-us/community/posts/203458976-Round-robin-ticket-assignment

## Usage


## Configuration

Configuration is done through an additional sheet in your spreadsheet. The sheet must be named "Configuration". An example of the rows and columns expected is provided in `configuration.csv`. The rows and columns must match for this script to work (Ex: subdomain is always in cell B1).

By default, the timezone the script runs in is that of the owner of the file. However, the timezone may be configured under File > Project Properties ([more details](https://developers.google.com/apps-script/reference/base/session#getScriptTimeZone())). Our script assumes all time inputs are hours in ET and based on a 24 hour clock.


## References
Documentation about the Google Apps Sheets API can be found here: https://developers.google.com/apps-script/reference/spreadsheet/
Triggers https://developers.google.com/apps-script/guides/triggers/
