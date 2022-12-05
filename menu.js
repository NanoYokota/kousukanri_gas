function addTasks()
{
  const functionLabel = "addTasks";
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function starts.` );
    console.time( functionLabel );
  }
  const labelsByCategory = labelListSh.getLabelValuesByCategory();
  for ( let i = 0; i < categories.length; i++ ) {
    const category = categories[ i ];
    const categoryName = category[ 0 ];
    const labels = labelsByCategory[ categoryName ];
    if ( debug ) {
      log( functionLabel, category, { label: "category", lineTwo: true });
      log( functionLabel, labels, { label: "labels", lineTwo: true });
    }
    addCategoryTasks( categoryName , labels );
  }
  if ( debug ) {
    console.timeEnd( functionLabel );
    console.log( `[DEBUG: ${ functionLabel }] Function ended.` );
  }
}

function addProjects()
{
  const functionLabel = "addProjects";
  if ( debug ) {
    log( functionLabel, "Function starts." );
  }
  const projectsByCategory = listSh.getProjectValuesByCategory();
  if ( debug ) {
    log( functionLabel, projectsByCategory, { label: "projectsByCategory", lineTwo: true } );
  }
  for ( let i = 0; i < categories.length; i++ ) {
    const category = categories[ i ];
    const categoryName = category[ 0 ];
    if ( debug ) {
      log( functionLabel, categoryName, { label: "categoryName", } );
    }
    if ( !projectsByCategory[ categoryName ] ) {
      if ( debug ) {
        log( functionLabel, "Projects NOT found in projectsByCategory." );
      }
      continue;
    }
    const projects = [ projectsByCategory[ categoryName ].values ];
    if ( debug ) {
      console.log( log( functionLabel, projects, { label: "projects", lineTwo: true } ) );
    }
    const shInfo = sheetsAnalysisByCat[ categoryName ];
    if ( debug ) {
      console.log( log( functionLabel, shInfo, { output: "return", label: "shInfo", lineTwo: true } ) );
      Logger.log( shInfo );
    }
    shInfo.putProjects( projects );
  }
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function ended.` );
  }
}

function putMonthTasks()
{
  const functionLabel = "putMonthTasks";
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function starts.` );
  }
  const sh = SpreadsheetApp.getActiveSheet();
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Sheet Name: ${ sh.getName() }`);
  }
  const shName = sh.getName();
  const shInfo = new MonthSheetInfo( shName );
  shInfo.getTotalsRangeForClear().clearContent();
  const tasks = shInfo.getTaskLabelValues();
  const tasksTrimmed = [ ...new Set( tasks.map( JSON.stringify ) ) ].map( JSON.parse );
  const numTrimmedData = tasksTrimmed.length;
  if ( !numTrimmedData || numTrimmedData <= 0 ) {
    console.error( log( functionLabel, "Rows num is invalid.", { type: "error", output: "return" } ) );
    log( functionLabel, numTrimmedData, { type: "error", label: "numTrimmedData" } );
  }
  if ( !shInfo.total.numbers.inputCol || shInfo.total.numbers.inputCol <= 0 ) {
    console.error( log( functionLabel, "Columns num is invalid.", { type: "error", output: "return" } ) );
    log( functionLabel, shInfo.total.numbers.inputCol, { type: "error", label: "shInfo.total.numbers.inputCol" } );
  }
  const totalRange = sh.getRange(
    shInfo.total.rows.firstInput,
    shInfo.total.cols.firstInput,
    numTrimmedData,
    shInfo.total.numbers.inputCol
  );
  totalRange.setValues( tasksTrimmed );
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function ended.` );
  }
}

function calcTotal()
{
  const functionLabel = "calcTotal";
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function starts.` );
    console.time( functionLabel );
  }
  // 月のシート名からインスタンスを取り出せる配列を作成。
  let baseShsMonth = {};
  if ( yearNow == yearStr ) {
    if ( debug ) {
      Logger.log( `[DEBUG: baseShsMonth] This year is first year.` );
    }
    for ( let i = monthNow; i >= monthStr; i-- ) {
      const sheetName = monthShName( yearStr, i );
      baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
      if ( debug ) {
        console.log( `[DEBUG: baseShsMonth] baseShsMonth[ sheetName ] ↓` );
        Logger.log( baseShsMonth[ sheetName ] );
      }
    }
  } else {
    if ( debug ) {
      Logger.log( `[DEBUG: baseShsMonth] This year is after first year.` );
    }
    // 初年度分の12月まで
    for ( let i = 12; i >= monthStr; i-- ) {
      const sheetName = monthShName( yearStr, i );
      baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
      if ( debug ) {
        console.log( `[DEBUG: baseShsMonth] baseShsMonth[ sheetName ] ↓` );
        Logger.log( baseShsMonth[ sheetName ] );
      }
    }
    // 初年度の次の年から去年までの12月分
    for ( let i = yearStr + 1; i < yearNow; i++ ) {
      for ( let j = 1; j < 12; j++ ) {
        const sheetName = monthShName( i, j );
        baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
        if ( debug ) {
          console.log( `[DEBUG: baseShsMonth] baseShsMonth[ sheetName ] ↓` );
          Logger.log( baseShsMonth[ sheetName ] );
        }
      }
    }
    // 今年の今月まで
    for ( let i = 1; i <= monthNow; i++ ) {
      const sheetName = monthShName( yearNow, i );
      baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
      if ( debug ) {
        console.log( `[DEBUG: baseShsMonth] baseShsMonth[ sheetName ] ↓` );
        Logger.log( baseShsMonth[ sheetName ] );
      }
    }
  }
  const sheetsByMonth = baseShsMonth;
  // カテゴリーごとに集計を計算して集計シートへ出力
  for ( let k = 0; k < categories.length; k++ ) { // カテゴリー
    const category = categories[ k ];
    const categoryName = category[ 0 ];
    const categoryLabel = category[ 1 ];
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] category: ${ category }` );
    }
    const shInfo = sheetsAnalysisByCat[ categoryName ];
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] shInfo.sheetName: ${ shInfo.sheetName }`);
    }
    const projectsNum = shInfo.project.numbers.raw
      ? shInfo.project.numbers.raw
      : shInfo.getProjectsNumber();
    const labelsNum = shInfo.labels.numbers.raw
      ? shInfo.labels.numbers.raw
      : shInfo.getRawLabelsNumber();
    if ( projectsNum < 1 || labelsNum < 1 ) {
      if ( debug ) {
        log( functionLabel, "This category has no valid tasks, projects or totals." );
      }
      continue;
    }
    if ( debug ) {
      log( functionLabel, `This category has ${ projectsNum } project(s).` );
      log( functionLabel, `This category has ${ labelsNum } label(s).` );
    }
    let totalValues = shInfo.getTotalValuesWithTaskLabels();
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] totalValues` );
      console.log( totalValues );
    }
    let indexPrj, arrayIndexPrj = {};
    const projects = shInfo.getRawProjectValues()[ 0 ];
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] projects ↓` );
      console.log( projects );
    }
    projects.forEach( ( project, i ) => {
      arrayIndexPrj[ project ] = i + shInfo.total.cols.firstInput - 1;
    } );
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] arrayIndexPrj ↓` );
      console.log( arrayIndexPrj );
    }
    // 月ごとのシートから集計を取得し配列にマージ。
    for ( key in sheetsByMonth ) { // 月ごとのシート
      if ( debug ) {
        console.log( `[DEBUG: ${ functionLabel }] key: ${ key }` );
      }
      const monthSh = sheetsByMonth[ key ];
      if ( !monthSh ) {
        const options = { type: "error", };
        log( functionLabel, 'monthSh is null.', options );
      }
      const totalsNum = monthSh.total.rows.numRaw
        ? monthSh.total.rows.numRaw
        : monthSh.getTotalRows().numRaw;
      if ( totalsNum <= 0 ) {
        if ( debug ) {
          console.log( `[DEBUG: ${ functionLabel }] This month has no totals.` );
        }
        continue;
      }
      const shTotals = monthSh.getRawTotalValues();
      if ( debug ) {
        console.log( `[DEBUG: ${ functionLabel }] shTotals↓` );
        console.log( shTotals );
      }
      for ( let i = 0; i < shTotals.length; i++ ) { // 対象のシートの集計
        const total = shTotals[ i ];
        if ( debug ) {
          console.log( `[DEBUG: ${ functionLabel }] This loop total of month sheet ↓` );
          console.log( total );
        }
        const prjTotal = total[ 0 ];
        const labelTotal = total[ 1 ];
        const timeTotal = total[ 2 ];
        if ( labelTotal.indexOf( categoryLabel ) < 0 ) {
          if ( debug ) {
            console.log( `[DEBUG: ${ functionLabel }] Task ${ labelTotal } has been skipped.` );
          }
          continue;
        }
        for ( let j = 0; j < totalValues.length; j++ ) { // 集計―シートの集計
          const totalValue = totalValues[ j ];
          if ( debug ) {
            console.log( `[DEBUG: ${ functionLabel }] totalValue before this loop process ↓` );
            console.log( totalValue );
          }
          const labelTotalValue = totalValue[ 2 ];
          if ( labelTotalValue == labelTotal ) { // 同じタスクが見つかった場合
            indexPrj = arrayIndexPrj[ prjTotal ];
            if ( debug ) {
              console.log( `[DEBUG: ${ functionLabel }] Same task ${ labelTotal } detected.` );
              console.log( `[DEBUG: ${ functionLabel }] Worked hour: ${ timeTotal / MINUTES_PER_HOUR }` );
            }
            totalValue[ indexPrj ] += timeTotal / MINUTES_PER_HOUR;
            break;
          } // end if ( labelTotalValue == labelTotal )
        } // end for ( let j = 0; j < totalValues.length; j++ )
      } // end for ( let i = 0; i < shTotals.length; i++ )
      if ( debug ) {
        console.log( `[DEBUG: ${ functionLabel }] totalValues after one month sheet process ↓` );
        console.log( totalValues );
      }
    } // end for ( key in sheetsByMonth )
    // 集計シートへ出力
    shInfo.putTotals( totalValues );
  } // end for ( let k = 0; k < categories.length; k++ )
  if ( debug ) {
    console.timeEnd( functionLabel );
    console.log( `[DEBUG: ${ functionLabel }] Function ended.` );
  }
}

function addCategories()
{
  const functionLabel = "addCategories";
  const lastShPos = thisSs.getSheets().length + 1 - 4;
  for ( let i = 0; i < categories.length; i++ ) {
    const category = categories[ i ];
    const newShName = `集計_${ category[ 0 ] }`;
    if ( thisSs.getSheetByName( newShName ) ) {
      if ( debug ) {
        console.log( `[DEBUG: ${ functionLabel }] Sheet ${ newShName } already exists.` );
      }
      continue;
    }
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] newShName: ${ newShName }` );
    }
    copySheet( formatAnalysisSh.sheet, newShName, lastShPos );
    sheetsAnalysisByCat[ category[ 0 ] ].init();
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] sheetsAnalysisByCat ↓` );
      console.log( sheetsAnalysisByCat );
    }
  }
}
