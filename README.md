# excel download

## How to Use
* git clone
* gradle bootRun
* access 'http://localhost:8080/download.xlsx'

## How to Use (ExcelBuiderUtil)
* Prepare headerMap (key is Integer from 0, value is String list(into header name))
* Prepare contentMap (key is Integer from 0, value is Object list)
* Prepare sheetList (String list)
* Mapping (key is String(headers, contents, sheets), value is the above things)
