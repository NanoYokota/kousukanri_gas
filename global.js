const transpose = a=> a[0].map((_, c) => a.map(r => r[c]));

// メタ情報
const now = new Date();
const yearNow = now.getFullYear();
const monthNow = now.getMonth() + 1;
const dateNow = now.getDate();
const dayNow = now.getDay();
const shNameNow = monthShName( yearNow, monthNow );

const MINUTES_PER_HOUR = 60;
const HOURS_PER_DAY = 8;

// 「フォーマット」シート
const formatSh = new FormatSheetInfo( "フォーマット" );

// 「LIST」シート
const listSh = new ListSheetInfo( "LIST" );

// LABEL LIST」シート
const labelListSh = new LabelListSheet( "LABEL LIST" );

// 月ごとのシート
const yearStr = 2022;
const monthStr = 7;
const numSheetsInStrYear_month = 12 - yearStr + 1;

// 「フォーマット_集計」シート
const formatAnalysisSh = new FormatAnalysisSheetInfo( "フォーマット_集計" );
if ( debug ) {
  console.log( `formatAnalysisSh ↓` );
  Logger.log( formatAnalysisSh );
}

// 「集計」シート
const categories = listSh.getRawCategoryValues();
const sheetsAnalysis = categories.map( ( cat ) => {
  return thisSs.getSheetByName( `集計_${ cat[ 0 ] }` );
} );

/* 
 * Name: sheetsAnalysisByCat
 * Structure:
 * {
 *   projectName: {
 *     category: string,
 *     sheet: Sheet Object,
 *     sheetName: string,
 *     lastCol: number,
 *     lastRow: {
 *       category: number,
 *     },
 *     num: {
 *       project: number,
 *       category: number,
 *     },
 *     projects: Array,
 *     tasks: Array,
 *   },
 * }
 */
let baseShsAnalysis = {};
categories.forEach( ( category ) => {
  const catName = category[ 0 ];
  if ( debug ) {
    Logger.log( `[DEBUG: baseShsAnalysis] category: ${ catName }` );
  }
  const sheetName = `集計_${ catName }`;
  baseShsAnalysis[ catName ] = new AnalysisSheetInfo( sheetName, catName );
  if ( debug ) {
    console.log( `[DEBUG: baseShsAnalysis] baseShsAnalysis[ catName ] ↓` );
    Logger.log( baseShsAnalysis[ catName ] );
  }
} );
const sheetsAnalysisByCat = baseShsAnalysis;