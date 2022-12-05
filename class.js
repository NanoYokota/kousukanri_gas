const template = () => {
  return { rows: {}, cols: {}, ranges: {}, values: {}, numbers: {}, };
};

class SheetInfo {
  constructor( sheetName ) {
    this.sheetName = sheetName;
    this.className = "SheetInfo";
    this.sheet = thisSs.getSheetByName( this.sheetName );
    this.input = template();
    this.row = {
      last: {},
      input: {},
    };
    this.col = {
      input: {},
    };
    this.num = {};
    this.range = {};
  }

  getLastColumn() {
    this.input.cols.last = this.sheet ? this.sheet.getLastColumn() : null;
    return this.input.cols.last;
  }
}

class FormatSheetInfo extends SheetInfo {
  constructor( sheetName ) {
    super( sheetName );
    this.className = "FormatSheetInfo";
    this.input.cols.first = 1;
    this.input.rows.first = 2;
    this.sheetMonth = template();
    this.sheetMonth.cols.firstInput = 1;
    this.sheetMonth.cols.yearInput = 1;
    this.sheetMonth.cols.monthInput = 2;
    this.sheetMonth.cols.lastInput = this.sheetMonth.cols.monthInput;
    this.sheetMonth.rows.labelInput = 1;
    this.sheetMonth.rows.firstInput = 2;
    this.task = template();
    this.task.cols.firstInput = 1;
    this.task.cols.dateInput = this.task.cols.firstInput;
    this.task.cols.projectInput = 2;
    this.task.cols.taskInput = 3;
    this.task.cols.detailInput = 4;
    this.task.cols.timeInput = 5;
    this.task.cols.lastInput = this.task.cols.timeInput;
    this.task.cols.num = this.task.cols.lastInput - this.task.cols.firstInput + 1;
    this.task.rows.labelInput = 3;
    this.task.rows.firstInput = 4;
    this.total = template();
    this.total.rows.labelInput = 3;
    this.total.rows.firstInput = 4;
    this.total.cols.firstInput = this.task.cols.lastInput + 2;
    this.total.cols.projectInput = this.total.cols.firstInput;
    this.total.cols.taskInput = this.total.cols.projectInput + 1;
    this.total.cols.sumInput = this.total.cols.taskInput + 1;
    this.total.cols.lastInput = this.total.cols.sumInput;
    this.total.cols.num = this.total.cols.lastInput - this.total.cols.firstInput + 1;
    this.total.numbers.inputCol = this.total.cols.taskInput - this.total.cols.projectInput + 1;
  }

  getSheetMonthCols() {
    return this.sheetMonth.cols;
  }

  getTaskCols() {
    return this.task.cols;
  }

  getTotalCols() {
    return this.total.cols;
  }

  getTaskRows() {
    this.task.rows.nowInput = this.sheet
      .getRange( this.task.rows.labelInput, this.task.cols.firstInput )
      .getNextDataCell( SpreadsheetApp.Direction.DOWN )
      .getRow();
    this.task.rows.numRaw = this.task.rows.nowInput - this.task.rows.labelInput;
    return this.task.rows;
  }

  getRawMonthInputRange() {
    this.sheetMonth.ranges.raw = this.sheet.getRange(
      this.sheetMonth.rows.firstInput,
      this.sheetMonth.cols.firstInput,
      1,
      2
    );
    return this.sheetMonth.ranges.raw;
  }
}

class MonthSheetInfo extends FormatSheetInfo {
  constructor( sheetName ) {
    super( sheetName );
    this.className = "MonthSheetInfo";
    this.col.input.total = {};
    this.num.col = {};
    this.month = template();
  }

  initParams() {
    const functionLabel = "MonthSheetInfo.initParams()";
    try {
      this.col.input.task.date = this.task.cols.firstInput;
      this.col.input.task.project = this.col.input.task.date + 1;
      this.col.input.task.name = this.col.input.task.project + 1;
      this.col.input.task.detail = this.col.input.task.name + 1;
      this.col.input.task.manMinute = this.col.input.task.detail + 1;
      this.col.input.task.last = this.col.input.task.manMinute;
      this.total.cols.firstInput = this.col.input.task.last + 2;
      this.col.input.total.project = this.total.cols.firstInput;
      this.col.input.total.task = this.col.input.total.project + 1;
      this.col.input.total.sum = this.col.input.total.task + 1;
      this.col.input.total.last = this.col.input.total.sum;
      this.num.col.task = this.col.input.task.last - this.input.cols.first + 1;
      this.num.col.total = this.col.input.total.last - this.total.cols.firstInput + 1;
      const lastRowTasks = this.sheet
        ? this.sheet
          .getRange( this.task.rows.firstInput - 1, this.col.input.task.date )
          .getNextDataCell( SpreadsheetApp.Direction.DOWN )
          .getRow()
        : null;
      this.row.last.task = lastRowTasks > 0 && lastRowTasks < 500
        ? lastRowTasks
        : this.task.rows.firstInput - 1;
      const numTasks = this.row.last.task - this.task.rows.firstInput + 1;
      this.num.task = numTasks > 0 && numTasks < 500 ? numTasks : 0;
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  initRanges() {
    this.initTaskRange();
    this.initTotalRange();
  }

  initTaskRange() {
    const functionLabel = "MonthSheetInfo.initTaskRange()";
    try {
      this.range.task = this.num.task > 0
        ? this.sheet.getRange(
            this.task.rows.firstInput,
            this.col.input.task.date,
            this.num.task,
            this.num.col.task
          )
        : null;
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  getTotalRows() {
    const funcName = `${ this.className }.getTotalRows`;
    const lastRowTotals = this.sheet
      .getRange( this.task.rows.labelInput, this.total.cols.firstInput )
      .getNextDataCell( SpreadsheetApp.Direction.DOWN )
      .getRow();
    if ( !lastRowTotals ) {
      throw `[ERROR: ${ funcName }] Failed to get now input row of total.`;
    }
    this.total.rows.nowInput = lastRowTotals > 0 && lastRowTotals < 500
      ? lastRowTotals : this.task.rows.labelInput;
    const numTotals = this.total.rows.nowInput - this.task.rows.labelInput;
    this.total.rows.numRaw = numTotals > 0 && numTotals < 500 ? numTotals : 0;
    return this.total.rows;
  }

  getRawTotalsRange() {
    const functionLabel = "MonthSheetInfo.getRawTotalsRange()";
    const rows = this.total.rows.numRaw
      ? this.total.rows
      : this.getTotalRows();
    if ( debug ) {
      const options = { label: "rows" };
      log( functionLabel, rows, options )
    }
    if ( !this.total.cols.num || this.total.cols.num <= 0 ) {
      const options = { type: 'error', label: "this.total.cols.num", output: "return" };
      console.error( log( functionLabel, this.total.cols.num, options ) );
      log( functionLabel, "Invalid column number.", { type: "error" } );
    } else if ( !rows.numRaw || rows.numRaw <= 0 ) {
      const options = { label: "rows.numRaw" };
      log( functionLabel, rows.numRaw, options );
      this.total.ranges.raw = this.sheet.getRange(
        rows.firstInput,
        this.total.cols.firstInput
      );
    } else {
      this.total.ranges.raw = this.sheet.getRange(
        rows.firstInput,
        this.total.cols.firstInput,
        rows.numRaw,
        this.total.cols.num
      );
    }
    return this.total.ranges.raw;
  }

  getRawTasksRange() {
    const rows = this.getTaskRows();
    const cols = this.getTaskCols();
    this.task.ranges.raw = this.sheet.getRange(
      rows.firstInput,
      cols.firstInput,
      rows.numRaw,
      cols.num
    );
    return this.task.ranges.raw;
  }

  getTotalsRangeForClear() {
    const rows = this.getTaskRows();
    const cols = this.getTotalCols();
    this.total.ranges.clear = this.sheet.getRange(
      this.total.rows.firstInput,
      cols.firstInput,
      rows.numRaw,
      2
    );
    return this.total.ranges.clear;
  }

  getTaskLabelRange() {
    const rows = this.getTaskRows();
    this.task.ranges.label = this.sheet.getRange(
      this.task.rows.firstInput,
      this.task.cols.projectInput,
      rows.numRaw,
      2
    );
    return this.task.ranges.label;
  }

  getRawTaskValues() {
    const range = this.getRawTasksRange();
    this.task.values.raw = range.getValues();
    return this.task.values.raw;
  }

  getTaskLabelValues() {
    const range = this.getTaskLabelRange();
    this.task.values.label = range.getValues();
    return this.task.values.label;
  }

  getRawTotalValues() {
    const range = this.getRawTotalsRange();
    if ( !range ) {
      log( funcName, "Invalid range.", { type: "error" } );
    }
    this.total.values.raw = range.getValues();
    return this.total.values.raw;
  }
}

class FormatAnalysisSheetInfo extends SheetInfo {
  constructor( sheetName ) {
    super( sheetName );
    this.className = "FormatAnalysisSheetInfo";
    this.row.input = {
      first: 3,
    };
    this.col.input = {
      first: 4,
    }
    this.col.category = 1;
    this.col.task = 2;
    this.col.taskLabel = 3;
    this.input = template();
    this.input.rows.label = 2;
    this.input.rows.first = 3;
    this.input.cols.first = 1;
    this.labels = template();
    this.labels.cols.firstInput = 1;
    this.labels.cols.category = this.labels.cols.firstInput;
    this.labels.cols.task = this.labels.cols.category + 1;
    this.labels.cols.sum = this.labels.cols.task + 1;
    this.labels.cols.lastInput = this.labels.cols.sum;
    this.labels.cols.num = this.labels.cols.lastInput - this.labels.cols.firstInput + 1;
    this.project = template();
    this.project.rows.input = this.input.rows.label;
    this.project.cols.firstInput = this.labels.cols.lastInput + 1;
    this.total = template();
    this.total.rows.firstInput = this.input.rows.first;
    this.total.cols.firstInput = this.labels.cols.lastInput + 1;
  }
}

class AnalysisSheetInfo extends FormatAnalysisSheetInfo {
  constructor( sheetName, category ) {
    super( sheetName );
    this.className = "AnalysisSheetInfo";
    this.category = category;
  }

  setRawProjectsNumber( projectsNum ) {
    const funcName = this.className + ".setRawProjectsNumber";
    if ( !projectsNum || projectsNum <= 0 ) {
      log( funcName, "Invalid number of projects.", { type: "error", } );
    }
    this.project.numbers.raw = projectsNum;
    return this.project.numbers.raw;
  }

  getNowLabelsRow() {
    const nowInput = this.sheet
      .getRange( this.input.rows.label, this.labels.cols.firstInput )
      .getNextDataCell( SpreadsheetApp.Direction.DOWN )
      .getRow();
    if ( !nowInput || nowInput <= 0 || nowInput > 500 ) {
      console.error( log( funcName, nowInput, { label: "nowInput", output: "return", type: "error" } ) );
      log( funcName, "The nowInput is invalid.", { type: "error" } );
    }
    this.labels.rows.nowInput = nowInput;
    return this.labels.rows.nowInput;
  }

  getNowProjectCol() {
    const funcName = this.className + ".getNowProjectCol";
    const nowCol = this.sheet
      .getRange( this.input.rows.label, this.labels.cols.lastInput )
      .getNextDataCell( SpreadsheetApp.Direction.NEXT )
      .getColumn();
    if ( !nowCol || nowCol <= this.project.cols.firstInput || nowCol > this.getLastColumn() ) {
      if ( debug ) {
        log( funcName, nowCol, { label: "nowCol" } );
      }
      this.project.cols.nowInput = this.project.cols.firstInput - 1;
    } else {
      this.project.cols.nowInput = nowCol;
    }
    return this.project.cols.nowInput;
  }

  getRawLabelsNumber() {
    const nowRow = this.getNowLabelsRow();
    this.labels.numbers.raw = nowRow - this.input.rows.label;
    return this.labels.numbers.raw;
  }

  getProjectsNumber() {
    this.project.numbers.raw = this.getNowProjectCol() - this.labels.cols.num;
    return this.project.numbers.raw;
  }

  getRawTotalsRange( projects = null ) {
    const functionLabel = "AnalysisSheetInfo.getRawTotalsRange()";
    const rowNum = this.getRawLabelsNumber();
    let colNum;
    if ( !projects ) {
      colNum = this.getProjectsNumber();
    } else {
      colNum = projects[ 0 ].length;
    }
    this.total.ranges.raw = rowNum > 0 && colNum > 0
      ? this.sheet.getRange(
          this.total.rows.firstInput,
          this.total.cols.firstInput,
          rowNum,
          colNum
        )
      : null;
    return this.total.ranges.raw;
  }

  getTotalsRangeWithTaskLabels( projects = null ) {
    const functionLabel = "AnalysisSheetInfo.getTotalsRangeWithTaskLabels()";
    const rowNum = this.getRawLabelsNumber();
    let colNum;
    if ( !projects ) {
      colNum = this.labels.cols.num + this.getProjectsNumber();
    } else {
      colNum = this.labels.cols.num + projects[ 0 ].length;
    }
    this.total.ranges.raw = rowNum > 0 && colNum > 0
      ? this.sheet.getRange(
          this.total.rows.firstInput,
          this.input.cols.first,
          rowNum,
          colNum
        )
      : null;
    return this.total.ranges.raw;
  }

  getRawProjectsRange() {
    const functionLabel = "AnalysisSheetInfo.getRawProjectsRange()";
    const projectNum = this.project.numbers.raw
      ? this.project.numbers.raw
      : this.getProjectsNumber();
    if ( !projectNum || projectNum <= 0 ) {
      console.error( log( functionLabel, projectNum, { label: "projectNum", type: "error", output: "return" } ) );
      log( functionLabel, "Invalid projectNum.", { type: "error" } );
    }
    this.project.ranges.raw = this.sheet.getRange(
      this.project.rows.input,
      this.project.cols.firstInput,
      1,
      projectNum
    );
    return this.project.ranges.raw;
  }

  getLabelClearRange() {
    const lastRow = this.sheet.getLastRow();
    const rowNum = lastRow - this.input.rows.label;
    if ( !rowNum || rowNum <= 0 ) {
      log( funcName, rowNum, { label: "rowNum", type: "error" } );
    }
    if ( !this.labels.cols.num || this.labels.cols.num <= 0 ) {
      log( funcName, this.labels.cols.num, { label: "this.labels.cols.num", type: "error" } );
    }
    return this.sheet.getRange(
      this.input.rows.first,
      this.labels.cols.category,
      rowNum,
      this.labels.cols.num
    );
  }

  getClearTotalsRange() {
    const funcName = this.className + ".getClearTotalsRange";
    const rowNum = this.sheet.getLastRow() - this.total.rows.firstInput;
    const colNum = this.sheet.getLastColumn() - this.total.cols.firstInput + 1;
    if ( !rowNum || rowNum <= 0 || !colNum || colNum <= 0 ) {
      log( funcName, "rowNum or colNum is 0 or null." );
      this.total.ranges.clear = this.sheet.getRange(
        this.total.rows.firstInput,
        this.total.cols.firstInput
      );
    } else {
      this.total.ranges.clear = this.sheet.getRange(
        this.total.rows.firstInput,
        this.total.cols.firstInput,
        rowNum,
        colNum
      );
    }
    return this.total.ranges.clear;
  }

  getRawProjectValues() {
    const functionLabel = "AnalysisSheetInfo.getRawProjectValues()";
    const range = this.getRawProjectsRange();
    if ( !range ) {
      log( functionLabel, "No range.", { type: "error" } );
    }
    this.project.values.raw = range.getValues();
    return this.project.values.raw;
  }

  initTasks() {
    const functionLabel = "AnalysisSheetInfo.initTasks()";
    this.initRangeTask();
    try {
      let tasks = [];
      if ( this.num.category > 0 && this.num.category <= 100 ) {
        tasks = this.range.task
          ? this.range.task.getValues()
          : null;
      }
      this.tasks = tasks;
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  getTotalValuesWithTaskLabels() {
    const functionLabel = "AnalysisSheetInfo.initParams()";
    const range = this.getTotalsRangeWithTaskLabels();
    if ( !range ) {
      log( functionLabel, "The range is invalid.", { type: "error" } );
    }
    this.total.values.withTaskLabels = range.getValues();
    return this.total.values.withTaskLabels;
  }

  putLabels( labels ) {
    const funcName = this.className + ".putLabels";
    this.getLabelClearRange().clearContent();
    if ( !labels.length || labels.length <= 0 ) {
      console.error( log( funcName, labels, { type : "error", output : "return", label: "labels" } ) );
      console.error( labels );
      log( funcName, "The labels must not be empty.", { type : "error", } );
    }
    if ( !this.labels.cols.num || this.labels.cols.num <= 0 ) {
      console.error( log( funcName, this.labels.cols.num, { type : "error", output : "return", label: "this.labels.cols.num" } ) );
      log( funcName, "The this.labels.cols.num must be more than 0.", { type : "error", } );
    }
    try {
      this.sheet.getRange(
        this.input.rows.first,
        this.labels.cols.category,
        labels.length,
        this.labels.cols.num
      ).setValues( labels );
    } catch ( e ) {
      log( funcName, e, { type : "error", } );
    }
  }

  putProjects( projects ) {
    const funcName = this.className + ".putProjects";
    const clearColNum = this.sheet.getLastColumn() - this.project.cols.firstInput + 1;
    let clearRange;
    if ( !clearColNum || clearColNum <= 0 ) {
      log( funcName, clearColNum, { type: "info", } )
      log( funcName, "The clearColNum is 0.", { type : "info", });
      clearRange = this.sheet.getRange(
        this.project.rows.input,
        this.project.cols.firstInput
      );
    } else {
      clearRange = this.sheet.getRange(
        this.project.rows.input,
        this.project.cols.firstInput,
        1,
        clearColNum
      );
    }
    const projectsNum = this.setRawProjectsNumber( projects[ 0 ].length );
    try {
      clearRange.clearContent();
      this.sheet.getRange(
        this.project.rows.input,
        this.project.cols.firstInput,
        1,
        projectsNum
      ).setValues( projects );
    } catch ( e ) {
      log( funcName, e, { type: 'error' } );
    }
  }

  putRawTotals( totals ) {
    const functionLabel = "AnalysisSheetInfo.putTotals()";
    this.clearTotals();
    const range = this.getRawTotalsRange();
    this.total.values.raw = totals;
    try {
      range.setValues( this.total.values.raw );
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  putTotals( totals ) {
    const functionLabel = "AnalysisSheetInfo.putTotals()";
    this.clearTotals();
    const range = this.getTotalsRangeWithTaskLabels();
    this.total.values.raw = totals;
    try {
      range.setValues( this.total.values.raw );
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  clearTotals() {
    const funcName = this.className + ".clearTotals";
    const range = this.getClearTotalsRange();
    try {
      range.clearContent();
    } catch ( e ) {
      log( funcName, e, { label: "error", } );
    }
  }
}

class ListSheetInfo extends SheetInfo {
  constructor( sheetName ) {
    super( sheetName );
    this.className = this.constructor.name;
    this.input = template();
    this.project = template();
    this.task = template();
    this.category = template();
    this.col.project = 1;
    this.col.projectCategory = 2;
    this.col.category = 4;
    this.col.categoryLabel = this.col.category + 1;
    this.task.rows.firstInput = 1;
    this.input.rows.first = 1;
    this.project.cols.firstInput = 1;
    this.project.cols.name = this.project.cols.firstInput;
    this.project.cols.category = this.project.cols.name + 1;
    this.project.cols.lastInput = this.project.cols.category;
    this.project.cols.num = this.project.cols.lastInput - this.project.cols.firstInput + 1;
    this.category.rows.firstInput = 1;
    this.category.cols.firstInput = 4;
    this.category.cols.nameInput = this.category.cols.firstInput;
    this.category.cols.labelInput = this.category.cols.nameInput + 1;
    this.category.cols.lastInput = this.category.cols.labelInput;
    this.category.cols.num = this.category.cols.lastInput - this.category.cols.firstInput + 1;
  }

  getProjectNowRow() {
    const nowRow = this.sheet
      .getRange( this.input.rows.first, this.project.cols.firstInput )
      .getNextDataCell( SpreadsheetApp.Direction.DOWN )
      .getRow();
    if ( !nowRow || nowRow <= 0 || nowRow > 500 ) {
      console.error( log( funcName, nowRow, { label: 'nowRow', output: "return" } ) );
      log( funcName, "The nowRow is invalid.", { type: 'error', } );
    }
    this.project.rows.nowInput = nowRow;
    return this.project.rows.nowInput;
  }

  getCategoryNowRow() {
    const nowRow = this.sheet
      .getRange( this.category.rows.firstInput, this.category.cols.firstInput )
      .getNextDataCell( SpreadsheetApp.Direction.DOWN )
      .getRow();
    if ( nowRow > 0 && nowRow < 100 ) {
      this.category.rows.nowInput = nowRow;
    } else {
      this.category.rows.nowInput = 0;
    }
    return this.category.rows.nowInput;
  }

  getRawProjectsNum() {
    const nowRow = this.getProjectNowRow();
    this.project.numbers.rawValues = nowRow - this.input.rows.first + 1;
    return this.project.numbers.rawValues;
  }

  getRawCategoriesNum() {
    const nowRow = this.getCategoryNowRow();
    this.category.numbers.rawValues = nowRow - this.category.rows.firstInput + 1;
    return this.category.numbers.rawValues;
  }

  initParams() {
    const functionLabel = "ListSheetInfo.initParams()";
    try {
      this.row.last.project = this.sheet
        .getRange( this.task.rows.firstInput, this.col.project )
        .getNextDataCell( SpreadsheetApp.Direction.DOWN )
        .getRow();
      this.row.last.category = this.sheet
        .getRange( this.task.rows.firstInput, this.col.category )
        .getNextDataCell( SpreadsheetApp.Direction.DOWN )
        .getRow();
      this.num.project = this.row.last.project;
      this.num.category = this.row.last.category;
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  initRangeProject() {
    const functionLabel = "ListSheetInfo.initRangeProject()";
    try {
      this.range.project = this.sheet.getRange(
        this.task.rows.firstInput,
        this.col.project,
        this.num.project,
        2
      );
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  getRawCategoriesRange() {
    const functionLabel = this.className + ".getRawCategoriesRange()";
    const categoryNum = this.getRawCategoriesNum();
    if ( categoryNum > 0 && this.category.cols.num > 0 ) {
      this.category.ranges.raw = this.sheet.getRange(
        this.category.rows.firstInput,
        this.category.cols.firstInput,
        categoryNum,
        this.category.cols.num
      );
    } else {
      this.category.ranges.raw = null;
    }
    return this.category.ranges.raw;
  }

  getRawProjectsRange() {
    const projectNum = this.getRawProjectsNum();
    if ( !projectNum || projectNum <= 0 ) {
      console.error( log( funcName, projectNum, { label: "projectNum", output: "return" } ) );
      log( funcName, "The projectNum is invalid.", { type : "error", } );
    }
    if ( !this.project.cols.num || this.project.cols.num <= 0 ) {
      console.error( log( funcName, this.project.cols.num, { label: "this.project.cols.num", output: "return" } ) );
      log( funcName, "this.project.cols.num is invalid.", { type : "error", } );
    }
    this.project.ranges.raw = this.sheet.getRange(
      this.input.rows.first,
      this.project.cols.firstInput,
      projectNum,
      this.project.cols.num
    );
    return this.project.ranges.raw;
  }

  getRawProjectValues() {
    const functionLabel = "ListSheetInfo.getRawProjectValues()";
    const range = this.getRawProjectsRange();
    if ( !range ) {
      log( functionLabel, "The range is invalid.", { type : "error", } );
    }
    this.project.values.raw = range.getValues();
    return this.project.values.raw;
  }

  /* 
  * Name: listSh.project.values.byCategory
  * Structure:
  * {
  *   category: {
  *     values: Array of projects,
  *     label: String of category name, 
  *   }
  * }
  */
  getProjectValuesByCategory() {
    const functionLabel = this.className + ".getProjectsByCategory";
    let projects = {};
    this.getRawProjectValues().forEach( ( project ) => {
      if ( !projects[ project[ 1 ] ] ) {
        projects[ project[ 1 ] ] = {
          label: project[ 1 ],
          values: [],
        }
      }
      projects[ project[ 1 ] ].values.push( project[ 0 ] );
    } );
    this.project.values.byCategory = projects;
    return this.project.values.byCategory;
  }

  getRawCategoryValues() {
    const functionLabel = `${ this.className }.getRawCategoryValues`;
    const range = this.getRawCategoriesRange();
    if ( !range ) {
      log( functionLabel, "The range is null.", { type: "error", } );
    }
    this.category.values.raw = range.getValues();
    return this.category.values.raw;
  }
}

class LabelListSheet extends SheetInfo {
  constructor( sheetName ) {
    super( sheetName );
    this.className = "LabelListSheet";
    this.col.category = 1;
    this.col.name = 3;
    this.col.label = 5;
    this.label = template();
    this.label.rows.labelInput = 1;
    this.label.rows.firstInput = 2;
    this.label.cols.firstInput = 1;
    this.label.cols.categoryNameInput = this.label.cols.firstInput;
    this.label.cols.categoryLabelInput = this.label.cols.categoryNameInput + 1;
    this.label.cols.taskNameInput = this.label.cols.categoryLabelInput + 1;
    this.label.cols.taskLabelInput = this.label.cols.taskNameInput + 1;
    this.label.cols.sumInput = this.label.cols.taskLabelInput + 1;
    this.label.cols.lastInput = this.label.cols.sumInput;
    this.label.cols.num = this.label.cols.lastInput - this.label.cols.firstInput + 1;
  }

  initParams() {
    const functionLabel = "LabelListSheet.initParams()";
    try {
      this.col.nameLabel = this.col.name + 1;
      this.col.categoryLabel = this.col.category + 1;
      this.row.last.category = this.sheet
        .getRange( 1, this.col.category )
        .getNextDataCell( SpreadsheetApp.Direction.DOWN )
        .getRow();
      this.row.last.name = this.sheet
        .getRange( 1, this.col.name )
        .getNextDataCell( SpreadsheetApp.Direction.DOWN )
        .getRow();
      this.num.category = this.row.last.category - this.label.rows.firstInput + 1;
      this.num.name = this.row.last.name - this.label.rows.firstInput + 1;
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  initRanges() {
    const functionLabel = "LabelListSheet.initRanges()";
    try {
      this.range.category = this.sheet.getRange(
        this.label.rows.firstInput,
        this.col.category,
        this.num.category,
        1
      );
      this.range.name = this.sheet.getRange(
        this.label.rows.firstInput,
        this.col.name,
        this.num.name,
        1
      );
      this.range.label = this.sheet.getRange(
        this.label.rows.firstInput,
        this.col.label,
        this.num.name,
        1
      );
      this.range.taskList = this.sheet.getRange(
        this.label.rows.firstInput,
        this.col.category,
        this.num.category,
        5
      );
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  getNowLabelRow() {
    const funcName = this.className + ".getNowLabelRow";
    const nowRow = this.sheet
      .getRange( this.label.rows.labelInput, this.label.cols.firstInput )
      .getNextDataCell( SpreadsheetApp.Direction.DOWN )
      .getRow();
    if ( debug ) {
      log( funcName, nowRow, { label: "nowRow" } );
    }
    if ( nowRow > 0 && nowRow <= this.sheet.getLastRow() ) {
      this.label.rows.nowInput = nowRow;
    } else {
      this.label.rows.nowInput = this.label.rows.labelInput;
    }
    return this.label.rows.nowInput;
  }

  getLabelCols () {
    const funcName = this.className + ".getLabelCols";
    if ( debug ) {
      log( funcName, this.label.cols.firstInput, { label: "this.label.cols.firstInput" } );
      log( funcName, this.label.cols.sumInput, { label: "this.label.cols.sumInput" } );
      log( funcName, this.label.cols.lastInput, { label: "this.label.cols.lastInput" } );
      log( funcName, this.label.cols.firstInput, { label: "this.label.cols.firstInput" } );
    }
    return this.label.cols;
  }

  getRawLabelsNumber() {
    this.label.numbers.raw = this.getNowLabelRow() - this.label.rows.labelInput;
    return this.label.numbers.raw;
  }

  getRawLabelsRange() {
    const funcName = this.className + ".getRawLabelsRange";
    const labelsNum = this.getRawLabelsNumber();
    const cols = this.getLabelCols();
    if ( labelsNum <= 0 || cols.num <= 0 ) {
      console.error( log( funcName, labelsNum, { type: "error", output: "return", label: "labelsNum" } ) );
      console.error( log( funcName, cols.num, { type: "error", output: "return", label: "cols.num" } ) );
      return null;
    }
    if ( debug ) {
      log( funcName, cols.num, { label: "cols.num" } );
    }
    this.label.ranges.raw = this.sheet.getRange(
      this.label.rows.firstInput,
      this.label.cols.firstInput,
      labelsNum,
      cols.num
    );
    return this.label.ranges.raw;
  }

  initCategories() {
    const functionLabel = "LabelListSheet.initCategories()";
    try {
      this.categories = this.range.category.getValues();
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  initTasks() {
    const functionLabel = "LabelListSheet.initTasks()";
    try {
      this.names = this.range.name.getValues();
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  initLabels() {
    const functionLabel = "LabelListSheet.initLabels()";
    try {
      this.labels = this.range.label.getValues();
    } catch ( e ) {
      console.error( `[ERROR: ${ functionLabel }] ${ e.message }` );
    }
  }

  getRawLabelValues() {
    const functionLabel = "LabelListSheet.getRawLabelValues()";
    const range = this.getRawLabelsRange();
    if ( !range ) {
      log( functionLabel, "Invalid range", { type: "error", } );
    }
    this.label.values.raw = range.getValues();
    return this.label.values.raw;
  } 

  getLabelValuesByCategory() {
    const functionLabel = this.className + ".getLabelValuesByCategory()";
    let labels = {};
    this.getRawLabelValues().forEach( ( label ) => {
      const category = label[ 0 ];
      const task = label[ 2 ];
      const sum = label[ 4 ];
      if ( !labels[ category ] ) {
        labels[ category ] = []; // タスクを入れていくための枠を追加
      }
      labels[ category ].push( [ category, task, sum ] );
    } );
    this.label.values.byCategory = labels;
    return this.label.values.byCategory;
  }
}