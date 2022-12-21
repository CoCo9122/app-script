# app-script

const onEdit = (e) => {
  //編集されたシートの名前を取得
  var sheetName = e.source.getSheetName()

  //編集されたセルの行数を取得
  var row = e.range.getRow()

  //編集されたセルの列数を取得
  var col = e.range.getColumn()

  Logger.log("シート："+sheetName+"の行："+row+",列"+col+"が編集されました。")

  var objs = changeListETL()

  for (obj of objs){
    if (sheetName == obj.before.sheet){
      if (row == obj.before.row && col == obj.before.col){
        resetCheck(row,col,sheetName)
        setCheck(obj.after)
        break
      }
    }
  }
}

const setCheck = (obj) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()

  const sheet = ss.getSheetByName(obj.sheet)

  var range = sheet.getRange(obj.col+obj.row)
  // Browser.msgBox(obj.col+obj.row)
  range.setValue(obj.string)

  Logger.log("シート："+obj.sheet+"の行："+obj.row+",列"+obj.col+"に"+obj.string+"が追加されました。")
}

const changeListETL = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet()

  const sheet = ss.getSheetByName('変更検知一覧')

  var values = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues()

  var objs = values.reduce(
    (obj, [beforeSheet,beforeRow,beforeCol,afterSheet,afterRow,afterCol,string]) => (
      [...obj, {
        "before":{
          "sheet":beforeSheet,
          "row":beforeRow,
          "col":beforeCol
        },
        "after":{
          "sheet":afterSheet,
          "row":afterRow,
          "col":afterCol,
          "string":string
        }
      }]
    ),[]
  )

  return objs
}

const resetCheck = (row,col,sheetName) => {

  const ss = SpreadsheetApp.getActiveSpreadsheet()

  const sheet = ss.getSheetByName(sheetName)

  var range = sheet.getRange(parseInt(row), parseInt(col))

  range.setValue("FALSE")
  Logger.log("シート："+sheetName+"の行："+row+",列"+col+"がリセットされました。")

}
