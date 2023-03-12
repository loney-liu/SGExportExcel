# SGExportExcel
## What is this python script can do
This is based on ShotGrid [Action Menu Item](https://developer.shotgridsoftware.com/581648bb/?title=Action+Menu+Items)(*AMI*) to help user export ShotGrid data to excel with thumbnail. [Export CSV](https://help.autodesk.com/view/SGSUB/ENU/?guid=SG_Tutorials_tu_import_bids_html) only export thumbnail url. There are some situation required thumbnail include in excel file.
## Why use AMI
When export data from ShotGrid Web Page, the data will be filtered by page filter. Export columns are selected. AMI can help to export selected columns and filtered data. It can save many times. This script will use [Custom Browser Protocol](https://developer.shotgridsoftware.com/af0c94ce/?title=Launching+Applications+Using+Custom+Browser+Protocols). The script can be used for http/https.
## Tested on 
python 3.10
## Required python package
- urllib
- Pillow
- openpyxl
- [shotgun python api](https://developer.shotgridsoftware.com/682204e9/?title=Python+API)
## Setup ShotGrid Action Menu Item
- This is work