function calcTotalMonthly()
{
  addCategories();
  addProjects();
  addTasks();
  calcTotal();
}

function makeNewMonthSh()
{
  let sheetName;
  if ( !isReleased ) {
    sheetName = monthShName( "2099", "99" );
  } else {
    sheetName = monthShName( yearNow, monthNow );
  }
  const thisMonthShName = sheetName;
  if ( debug ) {
    Logger.log( `[DEBUG:${ Function.name }] thisMonthShName: ${ thisMonthShName }` );
  }
  copySheet( formatSh.sheet, thisMonthShName, 1 );
  const newSh = new MonthSheetInfo( thisMonthShName );
  const newMonthData = [ [ yearNow, monthNow ] ];
  const rangeMonth = newSh.getRawMonthInputRange();
  if ( debug ) {
    Logger.log( `[DEBUG:${ Function.name }] Values of rangeMonth: ${ rangeMonth.getValues() }` );
  }
  rangeMonth.setValues( newMonthData );
}