function copySheet( sheet, newSheetName, position )
{
  const newSheet = sheet.copyTo( thisSs );
  newSheet.setName( newSheetName );
  thisSs.setActiveSheet( newSheet );
  thisSs.moveActiveSheet( position );
}

function monthShName( year, month )
{
  return `${ year }年${ month }月`;
}

function squeezeArray( array = [], value )
{
  const functionLabel = "squeezeArray";
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function starts.` );
    console.time( functionLabel );
  }
  let set = new Set( array );
  set.delete( value );
  if ( debug ) {
    console.timeEnd( functionLabel );
    console.log( `[DEBUG: ${ functionLabel }] Function ended.` );
  }
  return Array.from( set );
}

function monthSheets()
{
  const functionLabel = "monthSheets";
  let baseShsMonth = {};
  if ( yearNow == yearStr ) {
    if ( debug ) {
      Logger.log( `[DEBUG: ${ functionLabel }] This year is first year.` );
    }
    for ( let i = monthNow; i >= monthStr; i-- ) {
      const sheetName = monthShName( yearStr, i );
      baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
      if ( debug ) {
        console.log( `[DEBUG: ${ functionLabel }] baseShsMonth[ sheetName ] ↓` );
        Logger.log( baseShsMonth[ sheetName ] );
      }
    }
  } else {
    if ( debug ) {
      Logger.log( `[DEBUG: ${ functionLabel }] This year is after first year.` );
    }
    // 初年度分の12月まで
    for ( let i = 12; i >= monthStr; i-- ) {
      const sheetName = monthShName( yearStr, i );
      baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
      if ( debug ) {
        console.log( `[DEBUG: ${ functionLabel }] baseShsMonth[ sheetName ] ↓` );
        Logger.log( baseShsMonth[ sheetName ] );
      }
    }
    // 初年度の次の年から去年までの12月分
    for ( let i = yearStr + 1; i < yearNow; i++ ) {
      for ( let j = 1; j < 12; j++ ) {
        const sheetName = monthShName( i, j );
        baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
        if ( debug ) {
          console.log( `[DEBUG: ${ functionLabel }] baseShsMonth[ sheetName ] ↓` );
          Logger.log( baseShsMonth[ sheetName ] );
        }
      }
    }
    // 今年の今月まで
    for ( let i = 1; i <= monthNow; i++ ) {
      const sheetName = monthShName( yearNow, i );
      baseShsMonth[ sheetName ] = new MonthSheetInfo( sheetName );
      if ( debug ) {
        console.log( `[DEBUG: ${ functionLabel }] baseShsMonth[ sheetName ] ↓` );
        Logger.log( baseShsMonth[ sheetName ] );
      }
    }
  }
  if ( debug ) {
    const options = { label: "baseShsMonth", lineTwo: true, }
    log( functionLabel, baseShsMonth, options );
  }
  return baseShsMonth;
}

function calcAllTotals()
{
  const functionLabel = "calcAllTotals";
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function starts.` );
    console.time( functionLabel );
  }
  // カテゴリーごとに集計を計算して集計シートへ出力
  for ( let k = 0; k < categories.length; k++ ) { // カテゴリー
    const category = categories[ k ];
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] category: ${ category }` );
    }
    const categoryName = category[ 0 ];
    const categoryLabel = category[ 1 ];
    const shInfo = new AnalysisSheetInfo( categoryName, categoryLabel );
    if ( debug ) {
      console.log( `[DEBUG: ${ functionLabel }] shInfo.sheetName: ${ shInfo.sheetName }`);
    }
    if ( !shInfo.projectExists() || !shInfo.labelExists() ) {
      if ( debug ) {
        log( functionLabel, "This category has no labels or projects." );
        continue;
      }
    }
    // 月ごとのシートから集計を取得し配列にマージ。
    shInfo.calcAllMonth();
  } // end for ( let k = 0; k < categories.length; k++ )
  if ( debug ) {
    console.timeEnd( functionLabel );
    console.log( `[DEBUG: ${ functionLabel }] Function ended.` );
  }
}