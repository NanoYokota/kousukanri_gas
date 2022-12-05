function addCategoryTasks( category, labels )
{
  const functionLabel = "addCategoryTasks";
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] Function starts.` );
    console.time( functionLabel );
  }
  const shInfo = sheetsAnalysisByCat[ category ];
  shInfo.putLabels( labels );
  let emptyTotals = [];
  let emptyTotalsRow = [];
  const numLabels = shInfo.getRawLabelsNumber();
  for ( let i = 0; i < shInfo.getProjectsNumber(); i++ ) {
    emptyTotalsRow.push( 0 );
  }
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] emptyTotalsRow.length: ${ emptyTotalsRow.length }` );
  }
  for ( let i = 0; i < numLabels; i++ ) {
    emptyTotals.push( emptyTotalsRow );
  }
  if ( debug ) {
    console.log( `[DEBUG: ${ functionLabel }] emptyTotals.length: ${ emptyTotals.length }` );
  }
  shInfo.putRawTotals( emptyTotals );
  if ( debug ) {
    console.timeEnd( functionLabel );
    console.log( `[DEBUG: ${ functionLabel }] Function ended.` );
  }
}

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