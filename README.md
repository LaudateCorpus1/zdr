## About
This script is used to manage rotations within Zendesk.

This script builds off the work done here: https://support.zendesk.com/hc/en-us/community/posts/203458976-Round-robin-ticket-assignment

## Usage

0. Open up the target spreadsheet and click Tools > Script Editor.
0. Copy and paste the script in there.
0. Replace 'SHEET_ID' with your spreadsheet id. This can be found from the URL https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit?ts=123#gid=456
0. Configure a trigger to run the script automatically with the following parameters: main, Time-driven, Minutes timer, Every 30 minutes.
0. Set up your spreadsheet sheets and columns based on the examples in `agents.csv` and `configuration.csv`

## Configuration

Configuration is done through an additional sheet in your spreadsheet. The sheet must be named "Configuration". An example of the rows and columns expected is provided in `configuration.csv`. The rows and columns must match for this script to work (Ex: subdomain is always in cell B1).

The following fields are configurable through the configuration spreadsheet:

- __subdomain:__ The domain of your zendesk account
- __username:__ The username of your zendesk account that we will be connecting as
- __token:__ The token we will use to connect to your zendesk account
- __working hours:__ Used to populate the working hours dropdown.
This should be the start and end of working hours in EST converted to UTC and 24 hour notation separated by a dash.
For example, 9am - 5pm EST would be represented as `14-22` and 12:00PM - 8:00PM EST would be 17-1`.
- __daylight savings time?:__ Determine if we should compensate for daylight savings time. If this flag is set incorrectly, all calculations will be off by one. Value should be `yes` or `no`. For an example converter, check out https://www.timeanddate.com/worldclock/converter.html?iso=20180116T140000&p1=179&p2=1440

## References
- Google Apps Sheets API: https://developers.google.com/apps-script/reference/spreadsheet/

- Triggers https://developers.google.com/apps-script/guides/triggers/
