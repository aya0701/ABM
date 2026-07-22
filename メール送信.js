function sendAssignedUrlsByGmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const employeeSheet = ss.getSheetByName('社員マスター');

  if (!employeeSheet) {
    throw new Error(
      '「社員マスター」シートが見つかりません。'
    );
  }

  const startRow = 2;
  const lastRow = employeeSheet.getLastRow();
  const lastColumn = employeeSheet.getLastColumn();

  if (lastRow < startRow) {
    throw new Error(
      '社員マスターにデータがありません。'
    );
  }

  // 1回のテストで送信する最大件数
  // 必要に応じて変更
  const MAX_SEND_COUNT = 5;

  // 固定列
  const nameColumn = 1;        // A列：氏名
  const emailColumn = 2;       // B列：メールアドレス
  const deviceCountColumn = 6; // F列：デバイス数

  // ヘッダー取得
  const headers = employeeSheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const url1Column = findHeaderColumn_(
    headers,
    'URL1'
  );

  const url2Column = findHeaderColumn_(
    headers,
    'URL2'
  );

  const testTargetColumn = findHeaderColumn_(
    headers,
    'テスト送信対象'
  );

  const sendStatusColumn = findHeaderColumn_(
    headers,
    '送信状況'
  );

  const sentAtColumn = findHeaderColumn_(
    headers,
    '送信日時'
  );

  const requiredHeaders = [
    ['URL1', url1Column],
    ['URL2', url2Column],
    ['テスト送信対象', testTargetColumn],
    ['送信状況', sendStatusColumn],
    ['送信日時', sentAtColumn]
  ];

  requiredHeaders.forEach(([headerName, column]) => {
    if (column === -1) {
      throw new Error(
        `社員マスターに「${headerName}」ヘッダーが見つかりません。`
      );
    }
  });

  const rowCount = lastRow - startRow + 1;

  const values = employeeSheet
    .getRange(
      startRow,
      1,
      rowCount,
      lastColumn
    )
    .getDisplayValues();

  const logSheet =
    getOrCreateMailLogSheet_(ss);

  const logRows = [];

  let sentCount = 0;
  let skippedCount = 0;

  for (
    let index = 0;
    index < values.length;
    index++
  ) {
    if (sentCount >= MAX_SEND_COUNT) {
      break;
    }

    const row = values[index];
    const sheetRow = startRow + index;

    const employeeName = String(
      row[nameColumn - 1] ?? ''
    ).trim();

    const employeeEmail = String(
      row[emailColumn - 1] ?? ''
    ).trim();

    const deviceCount = Number(
      row[deviceCountColumn - 1]
    );

    const url1 = normalizeUrl_(
      row[url1Column - 1]
    );

    const url2 = normalizeUrl_(
      row[url2Column - 1]
    );

    const testTarget = normalizeSendTarget_(
      row[testTargetColumn - 1]
    );

    const sendStatus = String(
      row[sendStatusColumn - 1] ?? ''
    ).trim();

    // 指定した社員だけ送る
    if (testTarget !== '1') {
      continue;
    }

    // 送信済みは再送しない
    if (sendStatus === '送信済み') {
      continue;
    }

    // 氏名チェック
    if (!employeeName) {
      setMailErrorStatus_(
        employeeSheet,
        sheetRow,
        sendStatusColumn,
        '送信不可：氏名なし'
      );

      skippedCount++;
      continue;
    }

    // メールアドレスチェック
    if (!isValidEmailAddress_(employeeEmail)) {
      setMailErrorStatus_(
        employeeSheet,
        sheetRow,
        sendStatusColumn,
        '送信不可：メールアドレス不正'
      );

      skippedCount++;
      continue;
    }

    // デバイス数チェック
    if (
      deviceCount !== 1 &&
      deviceCount !== 2
    ) {
      setMailErrorStatus_(
        employeeSheet,
        sheetRow,
        sendStatusColumn,
        '送信不可：デバイス数不正'
      );

      skippedCount++;
      continue;
    }

    //デバイス数1　URL1必要
    if (
      deviceCount === 1 &&
      !url1
    ) {
      setMailErrorStatus_(
        employeeSheet,
        sheetRow,
        sendStatusColumn,
        '送信不可：URL1未配布'
      );

      skippedCount++;
      continue;
    }

    //デバイス数2　URL1・URL2両方必要
    if (
      deviceCount === 2 &&
      (
        !url1 ||
        !url2
      )
    ) {
      setMailErrorStatus_(
        employeeSheet,
        sheetRow,
        sendStatusColumn,
        '送信不可：URL1またはURL2未配布'
      );

      skippedCount++;
      continue;
    }

    // URL1とURL2が同じ場合は送信しない
    if (
      deviceCount === 2 &&
      url1 === url2
    ) {
      setMailErrorStatus_(
        employeeSheet,
        sheetRow,
        sendStatusColumn,
        '送信不可：URL1とURL2が重複'
      );

      skippedCount++;
      continue;
    }

    const subject =
      '【テスト送信】【ご案内】業務アプリのダウンロードについて';

    const plainBody = createAppMailPlainBody_(
      employeeName,
      url1,
      url2
    );

    const htmlBody = createAppMailHtmlBody_(
      employeeName,
      url1,
      url2
    );

    try {
      GmailApp.sendEmail(
        employeeEmail,
        subject,
        plainBody,
        {
          htmlBody: htmlBody,
          name: 'システム窓口'
        }
      );

      const sentAt = new Date();

      employeeSheet
        .getRange(
          sheetRow,
          sendStatusColumn
        )
        .setValue('送信済み');

      employeeSheet
        .getRange(
          sheetRow,
          sentAtColumn
        )
        .setValue(sentAt)
        .setNumberFormat(
          'yyyy/MM/dd HH:mm:ss'
        );

      //テスト送信対象を0に戻す　再送したい場合は、再度1を入力する
      employeeSheet
        .getRange(
          sheetRow,
          testTargetColumn
        )
        .setValue('0');

      logRows.push([
        sentAt,
        'テスト送信',
        employeeName,
        employeeEmail,
        deviceCount,
        url1,
        url2,
        '成功',
        ''
      ]);

      sentCount++;

    } catch (error) {
      const errorMessage = String(
        error && error.message
          ? error.message
          : error
      );

      setMailErrorStatus_(
        employeeSheet,
        sheetRow,
        sendStatusColumn,
        `送信失敗：${errorMessage}`
      );

      logRows.push([
        new Date(),
        'テスト送信',
        employeeName,
        employeeEmail,
        deviceCount,
        url1,
        url2,
        '失敗',
        errorMessage
      ]);

      skippedCount++;
    }
  }

  // ログ書き込み
  if (logRows.length > 0) {
    const logStartRow =
      logSheet.getLastRow() + 1;

    logSheet
      .getRange(
        logStartRow,
        1,
        logRows.length,
        logRows[0].length
      )
      .setValues(logRows);

    logSheet
      .getRange(
        logStartRow,
        1,
        logRows.length,
        1
      )
      .setNumberFormat(
        'yyyy/MM/dd HH:mm:ss'
      );
  }

  SpreadsheetApp.flush();

  ss.toast(
    `送信：${sentCount}件／送信不可・失敗：${skippedCount}件`,
    'テスト送信完了',
    7
  );
}

function createAppMailPlainBody_(
  employeeName,
  url1,
  url2
) {
  const lines = [];

  lines.push(`${employeeName} 様`);
  lines.push('');
  lines.push('お疲れ様です。');
  lines.push('');
  lines.push(
    '下記URLから、業務アプリをダウンロードしてください。'
  );
  lines.push('');

  lines.push('【URL1】');
  lines.push(url1);

  if (url2) {
    lines.push('');
    lines.push('【URL2】');
    lines.push(url2);
  }

  lines.push('');

  if (url2) {
    lines.push(
      '社用携帯と私用携帯の2台ご利用の方は、' +
      'URL1とURL2をそれぞれの端末で開き、' +
      'アプリをダウンロードしてください。'
    );
  } else {
    lines.push(
      'URLを開き、アプリをダウンロードしてください。'
    );
  }

  lines.push('');
  lines.push('【注意事項】');
  lines.push(
    'URLが失効した場合や、正常にダウンロードできない場合は、' +
    'システム窓口までご連絡ください。'
  );
  lines.push('');
  lines.push('よろしくお願いいたします。');

  return lines.join('\n');
}

function createAppMailHtmlBody_(
  employeeName,
  url1,
  url2
) {
  const safeName = escapeHtml_(employeeName);
  const safeUrl1 = escapeHtml_(url1);
  const safeUrl2 = escapeHtml_(url2);

  let html = `
   <p>${safeName} 様</p>
   
   <p>お疲れ様です。</p>
   
   <p>
    下記URLから、業務アプリをダウンロードしてください。
   </p>
   
   <p>
     <strong>【URL1】</strong><br>
     <a href="${safeUrl1}">
      URL1からアプリをダウンロードする
     </a>
    </p>
  `;

  if (url2) {
    html += `
      <p>
        <strong>【URL2】</strong><br>
        <a href="${safeUrl2}">
          URL2からアプリをダウンロードする
        </a>
      </p>
      
      <p>
       社用携帯と私用携帯の2台ご利用の方は、
       URL1とURL2をそれぞれの端末で開き、
       アプリをダウンロードしてください。
      </p>
    `;
  } else {
    html += `
      <p>
       URL1を開き、アプリをダウンロードしてください。
      </p>
    `;
  }

  html += `
    <div style="
      padding: 12px;
      margin-top: 18px;
      background-color: #fce8e6;
      border: 1px solid #d93025;
     ">
       <strong>【注意事項】</strong><br>
       各URLは一度使用すると再利用できません。<br>
       URLが失効した場合や、
       正常にダウンロードできない場合は、
       システム窓口までご連絡ください。
    </div>

    <p>よろしくお願いいたします。</p>


  `;

  return html;
}

function normalizeSendTarget_(value) {
  const normalizedValue = String(value ?? '')
    .trim()
    .normalize('NFKC')
    .toUpperCase();

  if (
    normalizedValue === '1' ||
    normalizedValue === 'TRUE' ||
    normalizedValue === '〇' ||
    normalizedValue === '○' ||
    normalizedValue === '送信'
  ) {
    return '1';
  }

  return '0'; 
}

function setMailErrorStatus_(
  employeeSheet,
  sheetRow,
  sendStatusColumn,
  message
) {
  employeeSheet
    .getRange(
      sheetRow,
      sendStatusColumn
    )
    .setValue(message);
}

function escapeHtml_(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

//メアドの形式を簡易チェックする
function isValidEmailAddress_(emailAddress) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(
    String(emailAddress ?? '').trim()
  );
}

function getOrCreateMailLogSheet_(ss) {
  const sheetName = 'メール送信ログ';

  let logSheet =
    ss.getSheetByName(sheetName);

  if (!logSheet) {
    logSheet =
      ss.insertSheet(sheetName);
  }

  if (logSheet.getLastRow() === 0) {
    logSheet
      .getRange(1, 1, 1, 9)
      .setValues([[
        '送信日時',
        '送信種別',
        '社員名',
        'メールアドレス',
        'デバイス数',
        'URL1',
        'URL2',
        '結果',
        'エラー内容'
      ]]);

    logSheet.setFrozenRows(1);
  }

  return logSheet;
}






