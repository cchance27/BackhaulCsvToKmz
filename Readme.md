# Backhaul Csv To Kmz

This script was designed for internal use but as it seemed so helpful, I decided to share it

## Point-to-Point KMZ
This file generates a KMZ file to show wireless base station/site locations, along with the wireless point-to-point backhauls between those sites. 

## CSV Input

The input for the script expects 2 csv files. 
> **sites.csv**: This file is a basic list of sites and site locations

| Field Name | Description | Required |
|----|----|----|
|BBCode|Short-Code Abbreviation of site (eg. XMZI)|required|
|Site|Full Site Name (eg. XMZ Industries)|required|
|Latitude|Decimal latitude of the site|required|
|Longitude|Decimal longitude of the site|required|

> **links.csv**: This file lists all the actual point-to-point backhauls along with details about the links, including the ability to include miscellaneous information that will show in the KMZ when the link is clicked.

| Field Name | Description | Required |
|----|----|----|
|From|Site BBCode for start of link|required|
|To|Site BBCode for end of link|required|
|Channel Size|MHz this link uses|required|
_Channel TX|Transmit Channel|required|
|_Channel RX|Receive Channel|required|
|_Vendor|Backhaul Vendor|optional|
|_Anything|You can freely add columns starting with _ and they will be included as an entry in the link description table|optional|

## Execution
Execution of this script is super simple.
1.	Copy sites.csv.example to sites.csv and update it with your site information.
2.	Copy links.csv.example to links.csv and update it with your links.
3.	Execute generate-link-kmz.ps1 from a powershell terminal

## TODO
* **Channel Sizes** - Currently only supports 40mhz and 120mhz channel sizes for line thickness
* **XML Generation Improvement** - Don't like the way I'm doing the XML generation but it works for now and other ways are much more fussy.
* **Error Checking** - Seriously haven't put any yet
* **KMZ Compression** - Scripts called generate-link-kmz but i don't have it compressing the kml to kmz yet, minor but not done. 
