_This package is __deprecated__ because it is not maintained any more. It still works, but for now saving objects to csv for further convert is much better way. If you want do something complicated with MS Excel sheet yuo may look to [xlsx-populate](https://github.com/dtjohnson/xlsx-populate) or [xlsx](https://github.com/sheetjs/js-xlsx)._

----

# excelize

Save array of similar objects to MS Excel sheet

It is just wrapper for `msexcel-builder`

## Install

```bash
npm install excelize
```

## Usage

```
excelize(obj, path, file, sheet, cb);
```

`obj` - array of objects with same keys
`path` - where to save file
`file` - name for new Excel file (xlsx)
`sheet` - name for sheet
`cb` - callback

## Example

```js
var excelize = require('excelize');

var testObj = [
    {name: 'Alice', age: '20'},
    {name: 'Bob', age: '21'},
    {name: 'Chuck', age: '22'}
];

excelize(testObj, './', 'sample.xlsx', 'sheet', function(err){
    if (err)
        throw err;
    else
        console.log('congratulations, your workbook created');
});
```

## License

MIT