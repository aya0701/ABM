function checkCompanyPhoneUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const phoneSheet = ss.getSheetByName('社用携帯マスター');
  const employeeSheet = ss.getSheetByName('社員マスター');

  if (!phoneSheet || !employeeSheet) {
    throw new Error(
      ' 「社用携帯マスター」または「社員マスター」シートが見つかりません。'
    );
  }

  const startRow = 2; //1行目は見出し

  const phoneLastRow = phoneSheet.getLastRow();
  const employeeLastRow = employeeSheet.getLastRow();

  if (employeeLastRow < startRow) {
    throw new Error('社員マスターに突合対象のデータがありません。');
  }

  const employeeRowCount = employeeLastRow - startRow + 1;
  const employeeNames = employeeSheet
    .getRange(startRow, 1, employeeRowCount, 1)
    .getDisplayValues();

  const employeeMap = new Map();

  employeeNames.forEach((row, index) => {
    const normalizedName = normalizeName_(row[0]);

    if (!normalizedName) {
      return
    }

    //同姓同名に備えて行番号を配列で保持

    if (!employeeMap.has(normalizedName)) {
      employeeMap.set(normalizedName, []);
    }

    employeeMap.get(normalizedName).push(index);
  });

  //初期値　社員マスターは突合結果0、デバイス数1
  const matchingResults = Array.from(
    { length: employeeRowCount },
    () => ['0']
  );

  const deviceCounts = Array.from(
    { length: employeeRowCount },
    () => ['1']
  );

  //社用携帯マスターにデータがある場合
  if (phoneLastRow >= startRow) {
    const phoneRowCount = phoneLastRow - startRow + 1;
    const phoneNameRange = phoneSheet.getRange(
      startRow,
      2,
      phoneRowCount,
      1,
    );

    const phoneNames = phoneNameRange.getDisplayValues();

    const backgroundcolors = phoneNames.map(row => {
      const normalizedName = normalizeName_(row[0]);


      if (
        normalizedName &&
        employeeMap.has(normalizedName)
      ){
        //同姓同名の場合は該当する全員をマッチ扱い
        employeeMap.get(normalizedName).forEach(employeeIndex => {
          matchingResults[employeeIndex][0] = '1';
          deviceCounts[employeeIndex][0] = '2'
        });

        return ['#f4cccc']; //うすピンク
      }

      return ['#ffffff']; //白
    });

    //社用携帯マスターB列の背景色を反映
    phoneNameRange.setBackgrounds(backgroundcolors);

  
  }

  //社員マスターE列・F列に反映
  employeeSheet
    .getRange(startRow, 5, employeeRowCount, 1)
    .setNumberFormat('0')
    .setValues(matchingResults);

  employeeSheet
    .getRange(startRow, 6, employeeRowCount, 1)
    .setNumberFormat('0')
    .setValues(deviceCounts);

  //デバイス数に応じてURLを割り当てる
  assignUrlsToEmployees_(
    employeeSheet,
    ss,
    employeeNames,
    deviceCounts,
    startRow
  );

  SpreadsheetApp.flush();

  ss.toast(
    '社用携帯マスターとの突合が完了しました。',
    '突合完了',
    5
  );
}

//氏名突合用に整備
function normalizeName_(name) {
  return String(name ?? '')
  .trim()
  .replace(/[\s　]+/g, '');
}

// CSVシートの未使用URLを社員マスターのURL1・URL2へ割り当てる
function assignUrlsToEmployees_(
  employeeSheet,
  ss,
  employeeNames,
  deviceCounts,
  startRow
) {
  const csvSheet = ss.getSheetByName('CSV');

  if (!csvSheet) {
    throw new Error('「CSV」シートが見つかりません。');
  }

  //ログシートを取得、なければ自動作成
  const logSheet = getOrCreateUrlLogSheet_(ss);

  //今回の実行で記録するログを一時保存する
  const logRows = [];

  // --------------------------------
  // 1. 各シートの見出しを取得
  // --------------------------------

  const employeeHeaders = employeeSheet
    .getRange(
      1,
      1,
      1,
      employeeSheet.getLastColumn()
    )
    .getDisplayValues()[0];

  const csvHeaders = csvSheet
    .getRange(
      1,
      1,
      1,
      csvSheet.getLastColumn()
    )
    .getDisplayValues()[0];

  // 社員マスターの列番号を探す
  const url1Column = findHeaderColumn_(
    employeeHeaders,
    'URL1'
  );

  const url2Column = findHeaderColumn_(
    employeeHeaders,
    'URL2'
  );

  // CSVシートの列番号を探す
  const csvUrlColumn = findHeaderColumn_(
    csvHeaders,
    'URL'
  );

  const unusedColumn = findHeaderColumn_(
    csvHeaders,
    '未使用'
  );

  const usedColumn = findHeaderColumn_(
    csvHeaders,
    '使用済み'
  );

  // --------------------------------
  // 2. 必要な見出しがあるか確認
  // --------------------------------

  if (url1Column === -1) {
    throw new Error(
      '社員マスターに「URL1」ヘッダーが見つかりません。'
    );
  }

  if (url2Column === -1) {
    throw new Error(
      '社員マスターに「URL2」ヘッダーが見つかりません。'
    );
  }

  if (csvUrlColumn === -1) {
    throw new Error(
      'CSVシートに「URL」ヘッダーが見つかりません。'
    );
  }

  if (unusedColumn === -1) {
    throw new Error(
      'CSVシートに「未使用」ヘッダーが見つかりません。'
    );
  }

  if (usedColumn === -1) {
    throw new Error(
      'CSVシートに「使用済み」ヘッダーが見つかりません。'
    );
  }

  // --------------------------------
  // 3. CSVシートのURL情報を取得
  // --------------------------------

  const csvLastRow = csvSheet.getLastRow();

  if (csvLastRow < startRow) {
    throw new Error(
      'CSVシートにURLデータがありません。'
    );
  }

  const csvRowCount = csvLastRow - startRow + 1;

  const csvUrlValues = csvSheet
    .getRange(
      startRow,
      csvUrlColumn,
      csvRowCount,
      1
    )
    .getDisplayValues();

  const unusedValues = csvSheet
    .getRange(
      startRow,
      unusedColumn,
      csvRowCount,
      1
    )
    .getDisplayValues();

  const usedValues = csvSheet
    .getRange(
      startRow,
      usedColumn,
      csvRowCount,
      1
    )
    .getDisplayValues();

  // --------------------------------
  // 4. 社員マスターの現在のURLを取得
  // --------------------------------

  const employeeRowCount = employeeNames.length;

  const url1Results = employeeSheet
    .getRange(
      startRow,
      url1Column,
      employeeRowCount,
      1
    )
    .getDisplayValues();

  const url2Results = employeeSheet
    .getRange(
      startRow,
      url2Column,
      employeeRowCount,
      1
    )
    .getDisplayValues();

  // --------------------------------
  // 5. CSV内のURL重複を確認
  // --------------------------------

  const csvUrlMap = new Map();

  csvUrlValues.forEach((row, index) => {
    const url = normalizeUrl_(row[0]);

    // URLが空白の行は無視
    if (!url) {
      return;
    }

    // 同じURLがCSV内に複数ある場合
    if (csvUrlMap.has(url)) {
      throw new Error(
        `CSVシート内でURLが重複しています：${url}`
      );
    }

    /*
     * URLをキーにして、
     * CSVデータ内の何番目にあるか記録する
     */
    csvUrlMap.set(url, index);
  });

  // --------------------------------
  // 6. 社員マスター内の既存URLを確認
  // --------------------------------

  const assignedUrls = new Set();

  employeeNames.forEach((row, index) => {
    const employeeName = normalizeName_(row[0]);

    // 社員名が空白の行は無視
    if (!employeeName) {
      return;
    }

    const currentUrl1 = normalizeUrl_(
      url1Results[index][0]
    );

    const currentUrl2 = normalizeUrl_(
      url2Results[index][0]
    );

    const currentUrls = [
      currentUrl1,
      currentUrl2
    ].filter(url => url !== '');

    currentUrls.forEach(url => {
      /*
       * 同じURLが、社員マスター内で
       * 既に別の場所にあった場合
       */
      if (assignedUrls.has(url)) {
        throw new Error(
          `社員マスター内でURLが重複しています：${url}`
        );
      }

      /*
       * 社員マスターに入っているURLが
       * CSVシートに存在しない場合
       */
      if (!csvUrlMap.has(url)) {
        throw new Error(
          `社員マスターのURLがCSVシートに存在しません：${url}`
        );
      }

      assignedUrls.add(url);

      /*
       * 社員マスターに既にあるURLは、
       * CSV側も配布済みとして扱う
       */
      const csvIndex = csvUrlMap.get(url);

      unusedValues[csvIndex][0] = '';
      usedValues[csvIndex][0] = url;
    });
  });

  // --------------------------------
  // 7. 配布可能なURL一覧を作る
  // --------------------------------

  const availableUrls = [];

  csvUrlValues.forEach((row, index) => {
    const originalUrl = normalizeUrl_(row[0]);

    let unusedUrl = normalizeUrl_(
      unusedValues[index][0]
    );

    const usedUrl = normalizeUrl_(
      usedValues[index][0]
    )

    //URL列が空白なら対象外
    if (!originalUrl) {
      return;
    }

    //使用済列にあるURLは配布済のため再利用不可
    if (usedUrl) {
      unusedValues[index][0] = '';
      return;
    }

    // 社員マスターに既にあるURLは除外
    if (assignedUrls.has(originalUrl)) {
      unusedValues[index][0] = '';
      usedValues[index][0] = originalUrl;
      return;
    }

    // URL列にURLがあるが未使用列が空白の場合、まだ使用されていなければ未使用列へ入れる
    if (!unusedUrl) {
      unusedValues[index][0] = originalUrl;
      unusedUrl = originalUrl;
    }

    //未使用列のURLとURL列のURLが同じものだけ配布候補とする
    if (unusedUrl === originalUrl) {
      availableUrls.push({
        url: originalUrl,
        csvIndex: index
      });

    }
      
  });

  // --------------------------------
  // 8. 新たに必要なURL数を計算
  // --------------------------------

  let requiredUrlCount = 0;

  employeeNames.forEach((row, index) => {
    const employeeName = normalizeName_(row[0]);

    if (!employeeName) {
      return;
    }

    const deviceCount = Number(
      deviceCounts[index][0]
    );

    const currentUrl1 = normalizeUrl_(
      url1Results[index][0]
    );

    const currentUrl2 = normalizeUrl_(
      url2Results[index][0]
    );

    // デバイス数が1の場合
    if (deviceCount === 1) {
      // URL1が空白の場合だけ新しいURLが必要
      if (!currentUrl1) {
        requiredUrlCount += 1;
      }
    }

    // デバイス数が2の場合
    if (deviceCount === 2) {
      // URL1が空白なら1件必要
      if (!currentUrl1) {
        requiredUrlCount += 1;
      }

      // URL2が空白なら1件必要
      if (!currentUrl2) {
        requiredUrlCount += 1;
      }
    }
  });

  // 配布できるURLが足りるか確認
  if (availableUrls.length < requiredUrlCount) {
    console.log(
      '未使用URLが不足しています。' +
      `必要数：${requiredUrlCount}件、` +
      `利用可能数：${availableUrls.length}件、` +
      `利用可能な分だけ割り当てます。`
    );
  }

  // --------------------------------
  // 9. URLを社員へ割り当てる
  // --------------------------------

  let availableUrlIndex = 0;

  employeeNames.forEach((row, index) => {
    const employeeName = normalizeName_(row[0]);

    if (!employeeName) {
      return;
    }

    const deviceCount = Number(
      deviceCounts[index][0]
    );

    let currentUrl1 = normalizeUrl_(
      url1Results[index][0]
    );

    let currentUrl2 = normalizeUrl_(
      url2Results[index][0]
    );

    // デバイス数が1の場合
    if (deviceCount === 1) {
      /*
       * URL1が空白の場合だけ、
       * 新しいURLを1件割り当てる
       */
      if (!currentUrl1) {
        //利用可能なURLがなくなった場合は社員には割り当てない
        if (availableUrlIndex >= availableUrls.length) {
          return;
        }

        const selectedUrl =
          availableUrls[availableUrlIndex];

        url1Results[index][0] = selectedUrl.url;

        // CSV側を配布済みに変更
        unusedValues[selectedUrl.csvIndex][0] = '';
        usedValues[selectedUrl.csvIndex][0] = selectedUrl.url;

        assignedUrls.add(selectedUrl.url);

        //ログへ記録
        logRows.push([
          new Date(),
          'URL割当',
          row[0],
          deviceCount,
          'URL1',
          selectedUrl.url,
          selectedUrl.csvIndex + startRow,
          '成功'
        ]);

        availableUrlIndex += 1;
      }

      /*
       * URL2に既存URLがある場合は、
       * 一度配布済みの可能性があるため
       * 勝手には削除しない
       */
      return;
    }

    // デバイス数が2の場合
    if (deviceCount === 2) {
      // URL1が空白なら割り当て


      if (!currentUrl1) {
        //URLが残っている場合だけURL1へ割り当てる
        if (availableUrlIndex < availableUrls.length) {
          const selectedUrl = 
            availableUrls[availableUrlIndex];

          url1Results[index][0] = selectedUrl.url;

          unusedValues[selectedUrl.csvIndex][0] = '';
          usedValues[selectedUrl.csvIndex][0] = selectedUrl.url;

          assignedUrls.add(selectedUrl.url);
          
          //ログへ記録
          logRows.push([
            new Date(),
            'URL割当',
            row[0],
            deviceCount,
            'URL1',
            selectedUrl.url,
            selectedUrl.csvIndex + startRow,
            '成功',
          ]);

          currentUrl1 = selectedUrl.url;

          availableUrlIndex += 1;
        }
      }
      // URL2が空白なら割り当て
      if (!currentUrl2) {
        //URLが残っている場合だけURL2へ割り当てる
        if (availableUrlIndex < availableUrls.length){
          const selectedUrl =
          availableUrls[availableUrlIndex];

          url2Results[index][0] = selectedUrl.url;

          unusedValues[selectedUrl.csvIndex][0] = '';
          usedValues[selectedUrl.csvIndex][0] = selectedUrl.url;

          assignedUrls.add(selectedUrl.url);

          //ログへ記録
          logRows.push([
            new Date(),
            'URL割当',
            row[0],
            deviceCount,
            'URL2',
            selectedUrl.url,
            selectedUrl.csvIndex + startRow,
            '成功'
          ]);

          currentUrl2 = selectedUrl.url;

          availableUrlIndex += 1;
        }
        
      }

      return;
    }

    /*
     * デバイス数が1・2以外の場合は
     * URLを割り当てない
     */
  });

  // --------------------------------
  // 10. 社員マスターへURLを反映
  // --------------------------------

  employeeSheet
    .getRange(
      startRow,
      url1Column,
      employeeRowCount,
      1
    )
    .setValues(url1Results);

  employeeSheet
    .getRange(
      startRow,
      url2Column,
      employeeRowCount,
      1
    )
    .setValues(url2Results);

  // --------------------------------
  // 11. CSVシートへ使用状況を反映
  // --------------------------------

  csvSheet
    .getRange(
      startRow,
      unusedColumn,
      csvRowCount,
      1
    )
    .setNumberFormat('@')
    .setValues(unusedValues);

    //今回URLを割り当てた場合だけログシートへ追記する
    if (logRows.length > 0) {
      logSheet
        .getRange(
          logSheet.getLastRow() + 1,
          1,
          logRows.length,
          logRows[0].length
        )
        .setValues(logRows);

        //実行日時の表示形式
        logSheet
          .getRange(
            logSheet.getLastRow() - logRows.length + 1,
            1,
            logRows.length,
            1
          )
          .setNumberFormat('yyyy/MM/dd/ HH:mm:ss');
    }

    console.log(
      `URLを${logRows.length}件割り当て、ログへ記録しました。`
    );

  csvSheet
    .getRange(
      startRow,
      usedColumn,
      csvRowCount,
      1
    )
    .setNumberFormat('@')
    .setValues(usedValues);
}


// ヘッダーから列番号を探す
function findHeaderColumn_(headers, targetHeader) {
  const normalizedTarget = 
    normalizeHeader_(targetHeader);

  const headerIndex = headers.findIndex(header => {
    return normalizeHeader_(header) === normalizedTarget;
  });

  if (headerIndex === -1) {
    return -1;
  }
  
  // 配列は0始まり、スプレッドシートの列は1始まり
  return headerIndex + 1;
}


//見出しを比較しやすい形に整える
function normalizeHeader_(header) {
  return String(header ?? '')
    .trim()
    .replace(/[\s　]+/g, '')
    .normalize('NFKC')
    .toUpperCase();
}

//URLの前後の空白を削除する
function normalizeUrl_(url) {
  return String(url ?? '')
    .trim();
}

//URL割当用のログシートを取得する、存在しない場合は自動で新規作成する
function getOrCreateUrlLogSheet_(ss) {
  const logSheetName = 'ログ';

  let logSheet = ss.getSheetByName(logSheetName);

  //ログシートがなければ作成
  if (!logSheet) {
    logSheet = ss.insertSheet(logSheetName);
  }

  //1行目が空白ならヘッダーを設定
  if (logSheet.getLastRow() === 0) {
    logSheet.getRange(1, 1, 1, 8).setValues([[
      '実行日時',
      '処理内容',
      '社員名',
      'デバイス数',
      '割当先',
      'URL',
      'CSV結果',
      '結果',
    ]]);

    logSheet.setFrozenRows(1);
  }

  return logSheet;
}



