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
    if ( debug ) {
      log( functionLabel, category, { label: "category", lineTwo: true } );
    }
    const categoryName = category[ 0 ];
    const categoryLabel = category[ 1 ];
    const labels = labelsByCategory[ categoryName ];
    if ( debug ) {
      log( functionLabel, labels, { label: "labels", lineTwo: true } );
    }
    const shInfo = new AnalysisSheetInfo( categoryName, categoryLabel );
    shInfo.putLabels( labels );
    shInfo.clearTotals();
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
    const categoryLabel = category[ 1 ];
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
    const shInfo = new AnalysisSheetInfo( categoryName, categoryLabel );
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
  }
}

function calcCategoryTotals()
{
  const functionLabel = "calcCategoryTotals";
  const shName = SpreadsheetApp.getActiveSheet().getName();
  const categoryName = shName.replace( "集計_", "" );
  const categoryLabel = listSh.getCategoryLabel( categoryName );
  if ( debug ) {
    log( functionLabel, categoryName, { label: "categoryName" } );
  }
  const shInfo = new AnalysisSheetInfo( categoryName, categoryLabel );
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] shInfo.sheetName: ${ shInfo.sheetName }`);
  }
  if ( !shInfo.projectExists() || !shInfo.labelExists() ) {
    if ( debug ) {
      log( functionLabel, "This category has no labels or projects." );
    }
    return;
  }
  shInfo.calcAllMonth();
}
