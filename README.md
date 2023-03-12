# SGExportExcel
## What is this python script can do
This is based on ShotGrid [Action Menu Item](https://developer.shotgridsoftware.com/581648bb/?title=Action+Menu+Items)(*AMI*) to help user export ShotGrid data to excel with thumbnail. [Export CSV](https://help.autodesk.com/view/SGSUB/ENU/?guid=SG_Tutorials_tu_import_bids_html) only export thumbnail url. There are some situation required thumbnail include in excel file.
## Why use AMI
When export data from ShotGrid Web Page, the data will be filtered by page filter. Export columns are selected. AMI can help to export selected columns and filtered data. It can save many times. This script will use [Custom Browser Protocol](https://developer.shotgridsoftware.com/af0c94ce/?title=Launching+Applications+Using+Custom+Browser+Protocols). The script can be used for http/https.
## What does this script doesn't support
Doesn't support export pivot columns.
## Tested on 
python 3.10
## Required python package
- urllib

- Pillow

- openpyxl

- [shotgun python api](https://developer.shotgridsoftware.com/682204e9/?title=Python+API)
## Setup ShotGrid Action Menu Item
#### Create a ShotGrid [script user](https://developer.shotgridsoftware.com/b6636515/?title=API+Overview#script-keys)

#### Create [AMI](https://developer.shotgridsoftware.com/67695b40/?title=Custom+Action+Menu+Items)

#### Sample

![ami](https://user-images.githubusercontent.com/17845155/224524235-362ff215-062c-42ff-bb17-39f6c96e5b29.jpg)

- Title: Export Excel

- Entity Type; Shot (can be any entity which be exported)

- URL: sgami://export_excel/<script_user>/<script_key>

- Configure Menu Options: Include in "Add Entity" dropdown menu on Entity pages

- Export Excel Menu

![export_menu](https://user-images.githubusercontent.com/17845155/224524316-0795b99a-1fe9-4d0d-afd2-f8a84539517b.jpg)
