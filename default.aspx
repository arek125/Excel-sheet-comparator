<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width,initial-scale=1.0">
    <script src="/sites/Debica/Compare%20Sheets/promise.min.js" type="text/javascript"></script>
    <link rel="stylesheet" type="text/css" href="/sites/Debica/Compare%20Sheets/pico.min.css"/>
    <script src="/sites/Debica/Compare%20Sheets/main.js" type="text/javascript"></script>
    <title>Excel sheet comparator</title>
  </head>
  <body>
    <main class="container">
        <h2>Excel sheet comparator</h2>
        <form>
          <div class="grid">
            <label for="excel1">Choose a first excel file:
              <input type="file" id="excel1" name="excel1" accept="application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required>
            </label>
            <label for="excel1sheetname">Sheet name of first excel file:
                <!-- <input type="text" id="excel1sheetname" name="excel1sheetname" value="Sheet1" required> -->
                <select name="excel1sheetname" id="excel1sheetname"></select>
            </label>
          </div>
          <div class="grid">
            <label for="excel2">Choose a second excel file:
              <input type="file" id="excel2" name="excel2" accept="application/vnd.ms-excel, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" required>
            </label>
            <label for="excel2sheetname">Sheet name of second excel file:
                <!-- <input type="text" id="excel2sheetname" name="excel2sheetname" value="Sheet1" required> -->
                <select name="excel2sheetname" id="excel2sheetname"></select>
            </label>
          </div>
          <label for="mainkey">Key column name:
            <input type="text" id="mainkey" name="mainkey" value="ID" required>
          </label>
          <button type="button" id="compere">Compare</button>
          <div id="output"></div>
        </form>
      </main>
  </body>
</html>
