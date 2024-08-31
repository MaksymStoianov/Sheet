/**
 * MIT License
 * 
 * Copyright (c) 2023 Maksym Stoianov
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */



/**
 * @class               Sheet
 * @namespace           Sheet
 * @version             1.3.0
 * @author              Maksym Stoianov <stoianov.maksym@gmail.com>
 * @license             MIT
 * @tutorial            https://maksymstoianov.com/
 * @see                 [GitHub](https://github.com/MaksymStoianov/Sheet)
 */
class Sheet {

  /**
   * Пазбирает строку `A1Notation`, возможно с преобразованием получаемого в процессе разбора значения.
   * 
   * #### Example 1
   * ```javascript
   * const range = SpreadsheetApp.getActiveRange();
   * const a1Notation = range.getA1Notation();
   * const result = Sheet.parseA1Notation(a1Notation);
   * 
   * console.log(result);
   * ```
   * 
   * #### Example 2
   * ```javascript
   * console.log(Sheet.parseA1Notation('A1:AZ10'));
   * console.log(Sheet.parseA1Notation('B5'));
   * console.log(Sheet.parseA1Notation('5:15'));
   * console.log(Sheet.parseA1Notation('M:X'));
   * console.log(Sheet.parseA1Notation('B2:B2'));
   * console.log(Sheet.parseA1Notation('B'));
   * console.log(Sheet.parseA1Notation('5'));
   * console.log(Sheet.parseA1Notation('15:5'));
   * console.log(Sheet.parseA1Notation('15:M5'));
   * ```
   * @param {string} input Разбираемая строка `A1Notation`.
   * @param {function} [reviver] Если параметр является функцией, определяет преобразование полученного в процессе разбора значения, прежде, чем оно будет возвращено вызывающей стороне.
   * @return {Object}
   */
  static parseA1Notation(input, reviver) {
    if (!arguments.length)
      throw new Error(`The parameters () do not match the signature for ${this.name}.`);

    if (!this.RegExp.A1NOTATION.test(input))
      throw new SyntaxError(`Разбираемая строка не является правильным A1Notation.`);

    input = input
      .replace(/:$/, '')
      .trim();

    const match = input.match(this.RegExp.A1NOTATION);

    const range = {
      "a1Notation": null,
      "isCell": null
    };

    const hasColon = /:/.test(input);


    /**
     * @param {*} input 
     * @return {Integer}
     */
    const _toInteger = input => {
      if (input === null || input === undefined)
        return null;

      const parsed = parseInt(input, 10);

      if (isNaN(parsed))
        return null;

      return parsed;
    };


    range.startRowPosition = (_toInteger(match.groups.startRowPosition) ?? 1);
    range.startRowIndex = (Number.isInteger(range.startRowPosition) ? range.startRowPosition - 1 : null);

    range.startColumnLabel = (input => typeof input === 'string' && input.length ? input : 'A')(match.groups.startColumnLabel);
    range.startColumnPosition = (input => typeof input === 'string' && input.length ? this.getColumnPositionByLabel(input) : null)(range.startColumnLabel);
    range.startColumnIndex = (input => input ? input - 1 : null)(range.startColumnPosition);

    range.endRowPosition = _toInteger(match.groups.endRowPosition);
    range.endRowIndex = (Number.isInteger(range.endRowPosition) ? range.endRowPosition - 1 : null);

    range.endColumnLabel = (input => typeof input === 'string' && input.length ? input : null)(match.groups.endColumnLabel);
    range.endColumnPosition = (input => typeof input === 'string' && input.length ? this.getColumnPositionByLabel(input) : null)(range.endColumnLabel);
    range.endColumnIndex = (input => input ? input - 1 : null)(range.endColumnPosition);

    range.isCell = (
      !hasColon ||
      (
        range.startRowIndex === range.endRowIndex &&
        range.startColumnIndex === range.endColumnIndex
      )
    );

    if (range.isCell) {
      range.numRows = 1;
      range.numColumns = 1;

      range.endRowIndex = range.startRowIndex;
      range.endRowPosition = range.startRowPosition;

      range.endColumnLabel = range.startColumnLabel;
      range.endColumnIndex = range.startColumnIndex;
      range.endColumnPosition = range.startColumnPosition;
    } else {
      if (Number.isInteger(range.startRowIndex) && Number.isInteger(range.endRowIndex)) {
        // Если строки указаны в обратном порядке, меняем их местами
        if (range.startRowIndex > range.endRowIndex) {
          [range.startRowPosition, range.endRowPosition] = [range.endRowPosition, range.startRowPosition];
          [range.startRowIndex, range.endRowIndex] = [range.endRowIndex, range.startRowIndex];
        }

        range.numRows = range.endRowIndex - range.startRowIndex + 1;
      } else {
        range.numRows = null;
      }

      if (Number.isInteger(range.startColumnIndex) && Number.isInteger(range.endColumnIndex)) {
        // Если столбцы указаны в обратном порядке, меняем их местами
        if (range.startColumnIndex > range.endColumnIndex) {
          [range.startColumnLabel, range.endColumnLabel] = [range.endColumnLabel, range.startColumnLabel];
          [range.startColumnPosition, range.endColumnPosition] = [range.endColumnPosition, range.startColumnPosition];
          [range.startColumnIndex, range.endColumnIndex] = [range.endColumnIndex, range.startColumnIndex];
        }

        range.numColumns = range.endColumnIndex - range.startColumnIndex + 1;
      } else {
        range.numColumns = null;
      }
    }


    // Формирование корректного a1Notation
    if (range.isCell) {
      range.a1Notation = `${range.startColumnLabel}${range.startRowPosition}`;
    } else if (range.endColumnLabel && range.endRowPosition) {
      range.a1Notation = `${range.startColumnLabel}${range.startRowPosition}:${range.endColumnLabel}${range.endRowPosition}`;
    } else if (range.endColumnLabel) {
      range.a1Notation = `${range.startColumnLabel}:${range.endColumnLabel}`;
    } else if (range.endRowPosition && !input.startsWith(range.startColumnLabel)) {
      range.a1Notation = `${range.startRowPosition}:${range.endRowPosition}`;
    } else if (range.endRowPosition) {
      range.a1Notation = `${range.startColumnLabel}${range.startRowPosition}:${range.startColumnLabel}${range.endRowPosition}`;
    } else {
      range.a1Notation = `${range.startColumnLabel}${range.startRowPosition}`;
    }

    if ((input => input[0] == input[1])(range.a1Notation.split(':'))) {
      range.isCell = true;

      if (range.startColumnLabel) {
        range.endColumnLabel = range.startColumnLabel;
        range.endColumnPosition = range.startColumnPosition;
        range.endColumnIndex = range.startColumnIndex;
      } else if (range.endColumnLabel) {
        range.startColumnLabel = range.endColumnLabel;
        range.startColumnPosition = range.endColumnPosition;
        range.startColumnIndex = range.endColumnIndex;
      }

      if (range.startRowIndex) {
        range.endRowPosition = range.startRowPosition;
        range.endRowIndex = range.startRowIndex;
      } else if (range.endRowIndex) {
        range.startRowPosition = range.endRowPosition;
        range.startRowIndex = range.endRowIndex;
      }

      range.a1Notation = `${range.startColumnLabel}${range.startRowPosition}`;
    }


    // Рекурсивная функция для применения reviver к каждому свойству объекта
    function applyReviver(obj, key, reviver) {
      if ((typeof obj === 'object' && obj !== null)) {
        for (const prop in obj) {
          if (!obj.hasOwnProperty(prop)) continue;

          const value = obj[prop];
          const revivedValue = applyReviver(value, prop, reviver);

          if (revivedValue === undefined) {
            delete obj[prop];
          } else {
            obj[prop] = revivedValue;
          }
        }
      }

      // Вызов reviver на текущем уровне
      return reviver.call(this, key, obj);
    }

    // Применение reviver, если он определён
    if (typeof reviver === 'function') {
      range = applyReviver(range, '', reviver);
    }

    return range;
  }



  /**
   * Преобразует метку столбца (букву или комбинацию букв) в номер столбца.
   * 
   * #### Example 1
   * ```javascript
   * console.log(Sheet.getColumnPositionByLabel('A'));   // Вывод: 1
   * console.log(Sheet.getColumnPositionByLabel('AA'));  // Вывод: 27
   * console.log(Sheet.getColumnPositionByLabel('AZ'));  // Вывод: 52
   * ```
   * @param {string} columnLabel Метка столбца (например, 'A', 'B', ..., 'AA').
   * @return {Integer} Номер столбца.
   */
  static getColumnPositionByLabel(columnLabel) {
    let column = 0;
    const length = columnLabel.length;

    for (let i = 0; i < length; i++) {
      column += (columnLabel.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }

    return column;
  }



  /**
   * Преобразует номер столбца в метку столбца (букву или комбинацию букв).
   * 
   * #### Example 1
   * ```javascript
   * console.log(Sheet.getColumnLabelByPosition(1));   // Вывод: A
   * console.log(Sheet.getColumnLabelByPosition(27));  // Вывод: AA
   * console.log(Sheet.getColumnLabelByPosition(52));  // Вывод: AZ
   * ```
   * @param {Integer} columnPosition Позиция столбца, начиная с 1.
   * @return {string} Метка столбца, соответствующая указанной позиции.
   */
  static getColumnLabelByPosition(columnPosition) {
    let columnLabel = '';

    while (columnPosition > 0) {
      const modulo = (columnPosition - 1) % 26;
      columnLabel = String.fromCharCode(65 + modulo) + columnLabel;
      columnPosition = Math.floor((columnPosition - 1) / 26);
    }

    return columnLabel;
  }



  /**
   * Создает и возвращает экземпляр класса [`Sheet`](#).
   */
  static newSheet(...args) {
    return Reflect.construct(this, args);
  }



  /**
   * Создает и возвращает экземпляр класса [`Cell`](#).
   */
  static newCell(...args) {
    return Reflect.construct(this.Cell, args);
  }



  /**
   * Проверяет, является ли заданное значение объектом типа [`Spreadsheet`](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet).
   * @param {*} input Значение для проверки.
   * @return {boolean}
   */
  static isSpreadsheet(input) {
    if (!arguments.length)
      throw new Error(`The parameters () don't match any method signature for ${this.name}.isSpreadsheet.`);

    return (
      input === Object(input) &&
      input?.toString() === 'Spreadsheet'
    );
  }



  /**
   * Проверяет, является ли заданное значение объектом типа [`Sheet`](https://developers.google.com/apps-script/reference/spreadsheet/sheet).
   * @param {*} input Значение для проверки.
   * @return {boolean}
   */
  static isSheet(input) {
    if (!arguments.length)
      throw new Error(`The parameters () don't match any method signature for ${this.name}.isSheet.`);

    return (
      input === Object(input) &&
      input?.toString() === 'Sheet'
    );
  }



  /**
   * Проверяет, является ли заданное значение объектом типа [`Sheet`](#).
   * @param {*} input Значение для проверки.
   * @return {boolean}
   */
  static isSheetLike(input) {
    if (!arguments.length)
      throw new Error(`The parameters () don't match any method signature for ${this.name}.isSheetLike.`);

    return (input instanceof this);
  }



  /**
   * Проверяет, является ли заданное значение объектом типа [`Range`](https://developers.google.com/apps-script/reference/spreadsheet/range).
   * @param {*} input Значение для проверки.
   * @return {boolean}
   */
  static isRange(input) {
    if (!arguments.length)
      throw new Error(`The parameters () don't match any method signature for ${this.name}.isRange.`);

    return (
      input === Object(input) &&
      input?.toString() === 'Range'
    );
  }



  /**
   * @overload
   * @param {Sheet} sheet Объект `Sheet` для записи данных.
   */
  /**
   * @overload
   * @param {Sheet} sheet Объект `Sheet` для записи данных.
   * @param {Array} fields Поля схемы по умолчанию.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const activeSheet = SpreadsheetApp.getActiveSheet();
   * const sheet = new Sheet(activeSheet);
   * 
   * console.log(sheet);
   * ```
   * @param {SpreadsheetApp.Sheet} sheet Экземпляр класса [`Sheet`](https://developers.google.com/apps-script/reference/spreadsheet/sheet).
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const activeSheet = SpreadsheetApp.getActiveSheet();
   * const fields = [ 'id', 'name', 'email' ];
   * const sheet = new Sheet(activeSheet, fields);
   * 
   * console.log(sheet);
   * ```
   * @param {SpreadsheetApp.Sheet} sheet Экземпляр класса [`Sheet`](https://developers.google.com/apps-script/reference/spreadsheet/sheet).
   * @param {Array} fields Поля схемы по умолчанию.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const sheetName = 'Sheet Name';
   * const sheet = new Sheet(sheetName);
   * 
   * console.log(sheet);
   * ```
   * @param {string} sheetName Имя листа в текущей активной электронной таблице.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const sheetName = 'Sheet Name';
   * const fields = [ 'id', 'name', 'email' ];
   * const sheet = new Sheet(sheetName, fields);
   * 
   * console.log(sheet);
   * ```
   * @param {string} sheetName Имя листа в текущей активной электронной таблице.
   * @param {Array} fields Поля схемы по умолчанию.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const sheetId = 0;
   * const sheet = new Sheet(sheetId);
   * 
   * console.log(sheet);
   * ```
   * @param {Integer} sheetId Id листа в текущей активной электронной таблице.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const sheetId = 0;
   * const fields = [ 'id', 'name', 'email' ];
   * const sheet = new Sheet(sheetId, fields);
   * 
   * console.log(sheet);
   * ```
   * @param {Integer} sheetId Id листа в текущей активной электронной таблице.
   * @param {Array} fields Поля схемы по умолчанию.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
   * const sheetName = 'Sheet Name';
   * const sheet = new Sheet(spreadsheet, sheetName);
   * 
   * console.log(sheet);
   * ```
   * @param {SpreadsheetApp.Spreadsheet} spreadsheet Экземпляр класса [`Spreadsheet`](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet).
   * @param {string} sheetName Имя листа в электронной таблице.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
   * const sheetName = 'Sheet Name';
   * const fields = [ 'id', 'name', 'email' ];
   * const sheet = new Sheet(spreadsheet, sheetName, fields);
   * 
   * console.log(sheet);
   * ```
   * @param {SpreadsheetApp.Spreadsheet} spreadsheet Экземпляр класса [`Spreadsheet`](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet).
   * @param {string} sheetName Имя листа в электронной таблице.
   * @param {Array} fields Поля схемы по умолчанию.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
   * const sheetId = 0;
   * const sheet = new Sheet(spreadsheet, sheetId);
   * 
   * console.log(sheet);
   * ```
   * @param {SpreadsheetApp.Spreadsheet} spreadsheet Экземпляр класса [`Spreadsheet`](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet).
   * @param {Integer} sheetId Id листа в электронной таблице.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
   * const sheetId = 0;
   * const fields = [ 'id', 'name', 'email' ];
   * const sheet = new Sheet(spreadsheet, sheetId, fields);
   * 
   * console.log(sheet);
   * ```
   * @param {SpreadsheetApp.Spreadsheet} spreadsheet Экземпляр класса [`Spreadsheet`](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet).
   * @param {Integer} sheetId Id листа в электронной таблице.
   * @param {Array} fields Поля схемы по умолчанию.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheetId = 'spreadsheet-id';
   * const sheetName = 'Sheet Name';
   * const sheet = new Sheet(spreadsheetId, sheetName);
   * 
   * console.log(sheet);
   * ```
   * @param {string} spreadsheetId Уникальный идентификатор электронной таблицы.
   * @param {string} sheetName Имя листа в электронной таблице.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheetId = 'spreadsheet-id';
   * const sheetName = 'Sheet Name';
   * const fields = [ 'id', 'name', 'email' ];
   * const sheet = new Sheet(spreadsheetId, sheetName, fields);
   * 
   * console.log(sheet);
   * ```
   * @param {string} spreadsheetId Уникальный идентификатор электронной таблицы.
   * @param {string} sheetName Имя листа в электронной таблице.
   * @param {Array} fields Поля схемы по умолчанию.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheetId = 'spreadsheet-id';
   * const sheetId = 0;
   * const sheet = new Sheet(spreadsheetId, sheetId);
   * 
   * console.log(sheet);
   * ```
   * @param {string} spreadsheetId Уникальный идентификатор электронной таблицы.
   * @param {Integer} sheetId Id листа в электронной таблице.
   */
  /**
   * @overload
   * #### Example
   * ```javascript
   * const spreadsheetId = 'spreadsheet-id';
   * const sheetId = 0;
   * const fields = [ 'id', 'name', 'email' ];
   * const sheet = new Sheet(spreadsheetId, sheetId, fields);
   * 
   * console.log(sheet);
   * ```
   * @param {string} spreadsheetId Уникальный идентификатор электронной таблицы.
   * @param {Integer} sheetId Id листа в электронной таблице.
   * @param {Array} fields Поля схемы по умолчанию.
   */
  constructor(...args) {
    if (!args.length)
      throw new Error(`The parameters () do not match the signature for ${this.constructor.name}.`);


    /**
     * @private
     * @readonly
     * @type {SpreadsheetApp.Sheet}
     */
    this._sheet = null;


    /**
     * @return {SpreadsheetApp.Spreadsheet}
     */
    const _getActiveSpreadsheet = () => SpreadsheetApp.getActiveSpreadsheet();


    /**
     * @param {SpreadsheetApp.Spreadsheet} spreadsheet
     * @param {string} sheetName 
     * @return {SpreadsheetApp.Sheet}
     */
    const _getSheetByName = (spreadsheet, sheetName) => (
      spreadsheet.getSheetByName(sheetName) ??
      spreadsheet.insertSheet(sheetName)
    );


    /**
     * @param {SpreadsheetApp.Spreadsheet} spreadsheet
     * @param {Iteger} sheetId 
     * @return {SpreadsheetApp.Sheet}
     */
    const _getSheetById = (spreadsheet, sheetId) => {
      const sheets = spreadsheet.getSheets();

      for (const sheet of sheets) {
        if (sheet.getSheetId() === sheetId) {
          return sheet;
        }
      }

      return null;
    };


    /**
     * @param {*} input 
     * @return {boolean}
     */
    const _isValidSpreadsheetId = input => (
      typeof input === 'string' &&
      input.length > 10
    );


    /**
     * @param {*} input 
     * @return {boolean}
     */
    const _isValidSheetName = input => (
      typeof input === 'string' &&
      input.length
    );


    /**
     * @param {*} input 
     * @return {boolean}
     */
    const _isValidSheetId = input => (
      Number.isInteger(input) &&
      input >= 0
    );


    /**
     * @param {*} input 
     * @return {boolean}
     */
    const _isFields = input => Array.isArray(input);


    /**
     * @type {Array}
     */
    let fields;


    /**
     * Case 1
     * @param {Sheet} sheet
     */
    if (args.length === 1 && (this.constructor.isSheetLike(args[0]) && this.constructor.isSheet(args[0]._sheet))) {
      this._sheet = args[0]._sheet;
    }


    /**
     * Case 2
     * @param {Sheet} sheet
     * @param {Array} fields
     */
    else if (args.length === 2 && (this.constructor.isSheetLike(args[0]) && this.constructor.isSheet(args[0]._sheet)) && _isFields(args[1])) {
      this._sheet = args[0];
      fields = args[1];
    }


    /**
     * Case 3
     * @param {SpreadsheetApp.Sheet} sheet
     */
    else if (args.length === 1 && this.constructor.isSheet(args[0])) {
      this._sheet = args[0];
    }


    /**
     * Case 4
     * @param {SpreadsheetApp.Sheet} sheet
     * @param {Array} fields
     */
    else if (args.length === 2 && (this.constructor.isSheet(args[0]) && _isFields(args[1]))) {
      this._sheet = args[0];
      fields = args[1];
    }


    /**
     * Case 5
     * @param {string} sheetName
     */
    else if (args.length === 1 && _isValidSheetName(args[0])) {
      this._sheet = _getSheetByName(_getActiveSpreadsheet(), args[0]);
    }


    /**
     * Case 6
     * @param {string} sheetName
     * @param {Array} fields
     */
    else if (args.length === 2 && _isValidSheetName(args[0]) && _isFields(args[1])) {
      this._sheet = _getSheetByName(_getActiveSpreadsheet(), args[0]);
      fields = args[1];
    }


    /**
     * Case 7
     * @param {Integer} sheetId
     */
    else if (args.length === 1 && _isValidSheetId(args[0])) {
      this._sheet = _getSheetById(_getActiveSpreadsheet(), args[0]);
    }


    /**
     * Case 8
     * @param {Integer} sheetId
     * @param {Array} fields
     */
    else if (args.length === 2 && _isValidSheetId(args[0]) && _isFields(args[1])) {
      this._sheet = _getSheetById(_getActiveSpreadsheet(), args[0]);
      fields = args[1];
    }


    /**
     * Case 9
     * @param {SpreadsheetApp.Spreadsheet} spreadsheet
     * @param {string} sheetName
     */
    else if (args.length === 2 && (this.constructor.isSpreadsheet(args[0]) && _isValidSheetName(args[1]))) {
      this._sheet = _getSheetByName(args[0], args[1]);
    }


    /**
     * Case 10
     * @param {SpreadsheetApp.Spreadsheet} spreadsheet
     * @param {string} sheetName
     * @param {Array} fields
     */
    else if (args.length === 3 && this.constructor.isSpreadsheet(args[0]) && _isValidSheetName(args[1]) && _isFields(args[2])) {
      this._sheet = _getSheetByName(args[0], args[1]);
      fields = args[2];
    }


    /**
     * Case 11
     * @param {SpreadsheetApp.Spreadsheet} spreadsheet
     * @param {Integer} sheetId
     */
    else if (args.length === 2 && (this.constructor.isSpreadsheet(args[0]) && _isValidSheetId(args[1]))) {
      this._sheet = _getSheetById(args[0], args[1]);
    }


    /**
     * Case 12
     * @param {SpreadsheetApp.Spreadsheet} spreadsheet
     * @param {Integer} sheetId
     * @param {Array} fields
     */
    else if (args.length === 3 && this.constructor.isSpreadsheet(args[0]) && _isValidSheetId(args[1]) && _isFields(args[2])) {
      this._sheet = _getSheetById(args[0], args[1]);
      fields = args[2];
    }


    /**
     * Case 13
     * @param {string} spreadsheetId
     * @param {string} sheetName
     */
    else if (args.length === 2 && (_isValidSpreadsheetId(args[0]) && _isValidSheetName(args[1]))) {
      this._sheet = _getSheetByName(SpreadsheetApp.openById(args[0]), args[1]);
    }


    /**
     * Case 14
     * @param {string} spreadsheetId
     * @param {string} sheetName
     * @param {Array} fields
     */
    else if (args.length === 3 && _isValidSpreadsheetId(args[0]) && _isValidSheetName(args[1]) && _isFields(args[2])) {
      this._sheet = _getSheetByName(SpreadsheetApp.openById(args[0]), args[1]);
      fields = args[2];
    }


    /**
     * Case 15
     * @param {string} spreadsheetId
     * @param {Integer} sheetId
     */
    else if (args.length === 2 && (_isValidSpreadsheetId(args[0]) && _isValidSheetId(args[1]))) {
      this._sheet = _getSheetById(SpreadsheetApp.openById(args[0]), args[1]);
    }


    /**
     * Case 16
     * @param {string} spreadsheetId
     * @param {Integer} sheetId
     * @param {Array} fields
     */
    else if (args.length === 3 && _isValidSpreadsheetId(args[0]) && _isValidSheetId(args[1]) && _isFields(args[2])) {
      this._sheet = _getSheetById(SpreadsheetApp.openById(args[0]), args[1]);
      fields = args[2];
    }


    else throw new Error('Invalid arguments: Unable to determine the correct overload.');


    if (!this.constructor.isSheet(this._sheet))
      throw new Error(`Invalid argument "sheet".`);


    if (fields && !_isFields(fields))
      throw new Error(`Invalid argument "fields".`);


    try {
      if (String(SheetSchema ?? '')) {
        // Получить схему листа

        /**
         * @type {SheetSchema.Schema}
         */
        this._sheet.schema = (
          (fields ? SheetSchema?.newSchema(fields) : SheetSchema?.getSchemaBySheet(this._sheet)) ??
          null
        );
      }
    } catch (error) {
      console.warn(error.stack);
    }


    /**
     * @type {number}
     */
    this._sheet.id = (this._sheet?.getSheetId() ?? null);


    /**
     * @type {string}
     */
    this._sheet.name = (this._sheet?.getName() ?? null);


    /**
     * @readonly
     * @private
     * @type {Proxy}
     */
    this._proxy = new Proxy(this, {

      /**
       * @param {Object} target 
       * @param {string} prop 
       * @param {Object} receiver
       * @return {*}
       */
      get(target, prop, receiver) {
        if (prop === 'inspect') {
          return null;
        }

        if (prop == '_proxy') {
          return receiver;
        }

        if (typeof prop === 'symbol' || ['_sheet'].includes(prop)) {
          return target[prop];
        }

        if (typeof target[prop] === 'function') {
          return (...args) => target[prop](...args);
        }

        if (typeof target._sheet[prop] === 'function') {
          return (...args) => target._sheet[prop](...args);
        }

        return (
          target[prop] ??
          target._sheet[prop] ??
          null
        );
      },

    });


    for (const key of Object.getOwnPropertyNames(this)) {
      if (key.startsWith('_')) {
        // Скрыть свойство
        Object.defineProperty(this, key, {
          "configurable": true,
          "enumerable": false,
          "writable": true
        });
      }
    }

    return this._proxy;
  }



  /**
   * Возвращает схему текущего листа электронной таблицы или `null`.
   * @return {SheetSchema.Schema}
   */
  getSchema() {
    try {
      if (SheetSchema && !this._sheet?.schema) {
        this._sheet.schema = (SheetSchema.getSchemaBySheet(this._sheet) ?? null);

        return this._sheet.schema;
      }
    } catch (error) {
      throw error;
    } finally {
      return null;
    }
  }



  /**
   * Устанавливает схему в текущий лист электронной таблицы.
   */
  /**
   * @overload
   * @param {SpreadsheetApp.Sheet} sheet Экземпляр класса [`Sheet`](https://developers.google.com/apps-script/reference/spreadsheet/sheet).
   * @param {SheetSchema.Schema} schema Экземпляр класса [`Schema`](#).
   * @return {SheetSchema.Schema}
   */
  /**
   * @overload
   * @param {SpreadsheetApp.Sheet} sheet Экземпляр класса [`Sheet`](https://developers.google.com/apps-script/reference/spreadsheet/sheet).
   * @param {SheetSchema.Field[]} fields Массив полей.
   * @return {SheetSchema.Schema}
   */
  insertSchema(schema) {
    try {
      if (SheetSchema) {
        this._sheet.schema = SheetSchema.insertSchema(this._sheet, schema);

        return this._sheet.schema;
      }
    } catch (error) {
      throw error;
    } finally {
      return false;
    }
  }



  /**
   * Удаляет схему из текущего листа электронной таблицы.
   * @return {boolean}
   */
  removeSchema() {
    try {
      if (SheetSchema) {
        SheetSchema.removeSchema(this._sheet);
        this._sheet.schema = null;
        return true;
      }
    } catch (error) {
      throw error;
    } finally {
      return false;
    }
  }



  /**
   * Возвращает прямоугольную сетку или объект значений для диапазона со значениями.
   * 
   * #### Example 1
   * ```javascript
   * const sheet = new Sheet('Sheet Name');
   * 
   * const values = sheet.getValues({
   *   displayValues: false,
   *   includeFrozenRows: true,
   *   includeFrozenCols: true,
   *   output: 'OBJECT_VALUES',
   *   rowNaming: 'ROW_POSITION',
   *   colNaming: 'COLUMN_POSITION',
   * });
   * 
   * console.log(values);
   * ```
   * @param {Object} [options = {}] 
   * @param {boolean} [options.displayValues = false] 
   * @param {boolean} [options.includeFrozenRows = false] 
   * @param {boolean} [options.includeFrozenCols = false] 
   * @param {boolean} [options.includeRowsHiddenByFilter = false] 
   * @param {boolean} [options.includeRowsHiddenByFilterView = false] 
   * @param {boolean} [options.includeRowsHiddenByUser = false] 
   * @param {boolean} [options.includeColumnsHiddenByUser = false] 
   * @param {string} [options.output = 'ARRAY'] Формат вывода.
   * Возможные значения:
   * - `ARRAY`: данные будут представлены в виде двумерного массива.
   * - `OBJECT`: данные будут представлены в виде объекта с объектами ячеек.
   * - `OBJECT_VALUES`: данные будут представлены в виде объекта со значениями ячеек.
   * @param {string} [options.rowNaming = 'ROW_POSITION'] Схема именования столбцов.
   * Возможные значения:
   * - `ROW_POSITION`: строки обозначаются по положению. Например: `Row1`, `Row2`, ..., `RowN`.
   * - `ROW_INDEX`: строки обозначаются по индексу. Например: `Row0`, `Row1`, ..., `RowN`.
   * - `POSITION`: строки обозначаются по положению. Например: `1`, `2`, ..., `N`.
   * - `INDEX`: строки обозначаются по индексу. Например: `0`, `1`, ..., `N`.
   * @param {string} [options.colNaming = 'COLUMN_POSITION'] Схема именования столбцов.
   * Возможные значения:
   * - `COLUMN_POSITION`: столбцы обозначаются по положению. Например: `Col1`, `Col2`, ..., `ColN`.
   * - `COLUMN_INDEX`: столбцы обозначаются по индексу. Например: `Col0`, `Col1`, ..., `ColN`.
   * - `COLUMN_LABEL`: столбцы обозначаются по буквам. Например: `ColA`, `ColB`, ..., `ColAA`.
   * - `FIELD_NAME`: столбцы обозначаются по названиям полей схемы (если доступны). Например: `name`, `size`, ..., `date`.
   * - `POSITION`: столбцы обозначаются по положению. Например: `1`, `2`, ..., `N`.
   * - `INDEX`: столбцы обозначаются по индексу. Например: `0`, `1`, ..., `N`.
   * - `LABEL`: столбцы обозначаются по буквам. Например: `A`, `B`, ..., `AA`.
   * @return {(Array|Object)}
   */
  getValues(options = {}) {
    const sheet = this._sheet;

    const frozenRows = sheet?.getFrozenRows();

    if (!Number.isInteger(frozenRows))
      throw new TypeError(`Frozen rows is not an integer.`);

    const lastRow = sheet?.getLastRow();

    if (!Number.isInteger(lastRow))
      throw new TypeError(`Last row is not an integer.`);

    let rowPosition = 1;
    let numRows = lastRow;

    if (options.includeFrozenRows !== true && frozenRows > 0) {
      rowPosition = frozenRows + 1;
      numRows = lastRow - frozenRows;

      if (numRows <= 0)
        throw new Error(`После замороженных строк нет данных.`);
    }


    const frozenCols = sheet?.getFrozenColumns();

    if (!Number.isInteger(frozenCols))
      throw new TypeError(`Frozen columns is not an integer.`);

    const lastCol = sheet?.getLastColumn();

    if (!Number.isInteger(lastCol))
      throw new TypeError(`Last column is not an integer.`);

    let colPosition = 1;
    let numCols = lastCol;

    if (options.includeFrozenCols !== true && frozenCols > 0) {
      colPosition = frozenCols + 1;
      numCols = lastCol - frozenCols;

      if (numCols <= 0)
        throw new Error(`После замороженных столбцов нет данных.`);
    }

    let values = [];

    if (numRows > 0 && numCols > 0) {
      const range = sheet?.getRange(rowPosition, colPosition, numRows, numCols);

      if (options.displayValues === true) {
        values = range?.getDisplayValues();
      } else {
        values = range?.getValues();
      }
    }

    if (!(Array.isArray(values) && (!values.length || values.every(Array.isArray))))
      throw new TypeError(`Values are invalid or improperly formatted.`);

    if (!['OBJECT', 'OBJECT_VALUES'].includes(options.output)) {
      return values;
    }


    const filter = sheet.getFilter();
    const rowsHidden = {};
    const colsHidden = {};
    const rows = {};

    console.log('filter:', filter);

    for (const [i, rowValues] of values.entries()) {
      const rowIndex = i + (options.includeFrozenRows !== true ? frozenRows : 0);
      const rowPosition = rowIndex + 1;


      // Найти отфильтрованные строки?
      if (options.includeRowsHiddenByFilter !== true && filter) {
        if (!rowsHidden[rowPosition]) {
          rowsHidden[rowPosition] = sheet.isRowHiddenByFilter(rowPosition);
        }
      }


      // Найти скрытые пользователем строки?
      if (options.includeRowsHiddenByUser !== true) {
        if (!rowsHidden[rowPosition]) {
          rowsHidden[rowPosition] = sheet.isRowHiddenByUser(rowPosition);
        }
      }


      // Обойти скрытые строки
      if (rowsHidden[rowPosition] === true) {
        continue;
      }


      const cols = {};

      for (const [i, cellValue] of rowValues.entries()) {
        const colIndex = i + (options.includeFrozenCols !== true ? frozenCols : 0);
        const colPosition = colIndex + 1;
        let colName;

        // Найти скрытые пользователем столбцы?
        if (options.includeColumnsHiddenByUser !== true) {
          if (!colsHidden[colPosition]) {
            colsHidden[colPosition] = sheet.isColumnHiddenByUser(colPosition);
          }
        }

        // Обойти скрытые столбцы
        if (colsHidden[colPosition] === true) {
          continue;
        }

        if (options.colNaming === 'FIELD_NAME') {
          colName = (sheet?.schema?.getFieldByIndex(colIndex)?._values?.name ?? null);
        }

        else if (options.colNaming === 'INDEX') {
          colName = `${colIndex}`;
        }

        else if (options.colNaming === 'POSITION') {
          colName = `${colPosition}`;
        }

        else if (options.colNaming === 'COLUMN_LABEL') {
          const columnLabel = Sheet.getColumnLabelByPosition(colPosition);

          colName = (columnLabel ? `Col${columnLabel}` : null);
        }

        else if (options.colNaming === 'COLUMN_INDEX') {
          colName = `Col${colIndex}`;
        }

        if (options.colNaming === 'COLUMN_POSITION' || !colName) {
          colName = `Col${colPosition}`;
        }

        const cell = Sheet.newCell(rowIndex, colIndex, cellValue);

        cols[colName] = (options.output === 'OBJECT_VALUES' ? cell.value : cell);
      }


      let rowName;

      if (options.rowNaming === 'INDEX') {
        rowName = `${rowIndex}`;
      }

      else if (options.rowNaming === 'POSITION') {
        rowName = `${rowPosition}`;
      }

      else if (options.rowNaming === 'ROW_INDEX') {
        rowName = `Row${rowIndex}`;
      }

      if (options.rowNaming === 'ROW_POSITION' || !rowName) {
        rowName = `Row${rowPosition}`;
      }

      rows[rowName] = cols;
    }

    return rows;
  }



  /**
   * Удаляет строки по условию.
   * 
   * #### Example 1
   * ```javascript
   * const sheet = new Sheet('Sheet Name');
   * 
   * // Чётные строки
   * sheet.deleteRows((values, rowIndex) => rowIndex % 2 === 0);
   * ```
   * 
   * #### Example 2
   * ```javascript
   * const sheet = new Sheet('Sheet Name');
   * 
   * // Нечётные строки и столбец 3 равен true
   * sheet.deleteRows((values, rowIndex) => rowIndex % 2 === 1 && values.Col3 === true, {
   *   displayValues: false,
   *   includeFrozenRows: true,
   *   includeFrozenCols: true,
   *   output: 'OBJECT_VALUES',
   *   colNaming: 'COLUMN_POSITION',
   * });
   * ```
   * @param {function} callback Функция, которая будет вызвана для каждой строки.
   * Если функция возвращает `true`, то строка остаётся, если `false`, то удаляется.
   * @param {Object} [options = {}] Опции получения данных. Наследуется от `getValues()`.
   */
  /**
   * Удаляет несколько строк, начиная с заданной позиции строки.
   * @overload
   * @param {Integer} rowPosition	Позиция первой удаляемой строки.
   * @param {Integer} howMany Количество строк, которые необходимо удалить.
   */
  deleteRows(...args) {
    const sheet = this._sheet;
    const frozenRows = sheet?.getFrozenRows();

    if (!Number.isInteger(frozenRows))
      throw new TypeError(`Frozen rows is not an integer.`);


    /**
     * @param {function} callback 
     * @param {Object} options 
     * @return {void}
     */
    const _deleteRowsByConditional = (callback, options = {}) => {
      options.rowNaming = 'INDEX';
      let values = this.getValues(options);

      // Преобразование: Объекта в Массив объектов.
      if (['OBJECT', 'OBJECT_VALUES'].includes(options.output)) {
        const result = [];

        for (const rowIndex in values) {
          const i = rowIndex - (options.includeFrozenRows !== true ? frozenRows : 0);
          result[i] = values[rowIndex];
        }

        values = result;
      }


      // Поиск строк для удаления

      const length = values.length - 1;
      let rowsToDelete = new Map();
      let startRowPosition = null;
      let numRows = 0;

      for (const [i, rowValues] of values.entries()) {
        const rowIndex = i + (options.includeFrozenRows !== true ? frozenRows : 0);
        const rowPosition = rowIndex + 1;

        const isTrue = callback.apply(null, [rowValues, rowPosition]);

        if (isTrue) {
          if (startRowPosition === null) {
            startRowPosition = rowPosition;
          }

          numRows += 1;
        }

        if ((!isTrue && numRows > 0) || i === length) {
          // Если последовательность строк для удаления закончилась, сохраняем ее
          rowsToDelete.set(startRowPosition, numRows);
          startRowPosition = null;
          numRows = 0;
        }
      }

      // Удаление строк в обратном порядке, чтобы избежать смещения
      if (rowsToDelete.size > 0) {
        // Преобразуем Map в массив и сортируем его в обратном порядке
        rowsToDelete = [...rowsToDelete].reverse();
        const length = rowsToDelete.length - 1;

        for (let i = 0; i <= length; i++) {
          const [startRow, numRows] = rowsToDelete[i];

          // Если это последняя группа строк для удаления
          if (i === length) {
            const lastRowToDelete = startRow + numRows - 1;
            const maxRows = sheet.getMaxRows();

            // Проверка: убедиться, что останется хотя бы одна строка после закрепленных строк
            if (maxRows - lastRowToDelete < 1) {
              // Добавление новой строки, если удаление последней оставит таблицу пустой
              sheet.insertRowAfter(maxRows);

              console.info('Inserted one row to ensure at least one row remains after the frozen rows.');
            }
          }

          sheet.deleteRows(startRow, numRows);

          console.info(`Deleted ${numRows} row(s) starting from row: ${startRow}`);
        }
      }

      return void 0;
    };


    /**
     * Case 1
     * @param {Integer} rowPosition
     * @param {Integer} howMany
     */
    if (args.length === 2 && (Number.isInteger(args[0]) && Number.isInteger(args[1]))) {
      return sheet.deleteRows(...args);
    }


    /**
     * Case 2
     * @param {function} callback
     */
    else if (args.length === 1 && (typeof args[0] === 'function')) {
      return _deleteRowsByConditional(...args);
    }


    /**
     * Case 3
     * @param {function} callback
     * @param {Object} options
     */
    else if (args.length === 2 && (typeof args[0] === 'function' && typeof args[1] === 'object' || (args[1] == null || args[1] == undefined))) {
      return _deleteRowsByConditional(...args);
    }


    else throw new Error('Invalid arguments: Unable to determine the correct overload.');
  }



  /**
   * @todo
   * Удаляет столбцы по условию.
   * 
   * #### Example 1
   * ```javascript
   * const sheet = new Sheet('Sheet Name');
   * 
   * // Чётные столбцы
   * sheet.deleteColumns((values, colIndex) => colIndex % 2 === 0);
   * ```
   * 
   * #### Example 2
   * ```javascript
   * const sheet = new Sheet('Sheet Name');
   * 
   * // Нечётные столбцы и ячейка в строке 3 равна true
   * sheet.deleteColumns((values, colIndex) => colIndex % 2 === 1 && values.Row3 === true, {
   *   displayValues: false,
   *   includeFrozenRows: true,
   *   includeFrozenCols: true,
   *   output: 'OBJECT_VALUES',
   *   colNaming: 'COLUMN_POSITION',
   * });
   * ```
   * @param {function} callback Функция, которая будет вызвана для каждого столбца.
   * Если функция возвращает `true`, то столбец остаётся, если `false`, то удаляется.
   * @param {Object} [options = {}] Опции получения данных. Наследуется от `getValues()`.
   */
  /**
   * Удаляет несколько столбцов, начиная с заданной позиции столбца.
   * @overload
   * @param {Integer} colPosition	Позиция первого удаляемого столбца.
   * @param {Integer} howMany Количество столбцов, которые необходимо удалить.
   */
  deleteColumns(...args) {
    const sheet = this._sheet;
    const frozenCols = sheet?.getFrozenColumns();

    if (!Number.isInteger(frozenCols))
      throw new TypeError(`Frozen columns is not an integer.`);


    /**
     * @param {function} callback 
     * @param {Object} options 
     * @return {void}
     */
    const _deleteColumnsByConditional = (callback, options = {}) => {
      // options.colNaming = 'INDEX';
      // let values = this.getValues(options);

      // TODO
      throw new Error(`Метод ${this.constructor.name}.deleteColumns еще в разработке!`);

      return void 0;
    };


    /**
     * Case 1
     * @param {Integer} colPosition
     * @param {Integer} howMany
     */
    if (args.length === 2 && (Number.isInteger(args[0]) && Number.isInteger(args[1]))) {
      return sheet.deleteColumns(...args);
    }


    /**
     * Case 2
     * @param {function} callback
     */
    else if (args.length === 1 && (typeof args[0] === 'function')) {
      return _deleteColumnsByConditional(...args);
    }


    /**
     * Case 3
     * @param {function} callback
     * @param {Object} options
     */
    else if (args.length === 2 && (typeof args[0] === 'function' && typeof args[1] === 'object' || (args[1] == null || args[1] == undefined))) {
      return _deleteColumnsByConditional(...args);
    }


    else throw new Error('Invalid arguments: Unable to determine the correct overload.');
  }



  /**
   * Добавляет столбцы справа текущей области данных на листе.
   * Если содержимое ячейки начинается с `=`, оно интерпретируется как формула.
   * 
   * #### Example
   * ```javascript
   * const sheet = new Sheet('Sheet Name');
   * 
   * sheet.appendColumns([
   *   ["1-1", "1-2", "1-3"],
   *   ["2-1", "2-2", "2-3"],
   *   ["3-1", "3-2", "3-3"]
   * ]);
   * ```
   * @param {Array} colsContents
   */
  appendColumns(colsContents) {
    if (!Array.isArray(colsContents))
      throw new Error();

    if (!colsContents.every(item => Array.isArray(item)))
      throw new Error();

    const numRows = colsContents.length;

    if (!(colsContents.length > 0))
      throw new Error();

    const numColumns = colsContents[0].length;

    if (!(numColumns > 0))
      throw new Error();

    const lastCol = this._sheet?.getLastColumn();
    const columnPosition = lastCol + 1;

    const range = this.getRange(1, columnPosition, numRows, numColumns);

    range.setValues(values);

    return this;
  }



  /**
   * Добавляет строки внизу текущей области данных на листе.
   * Если содержимое ячейки начинается с `=`, оно интерпретируется как формула.
   * 
   * #### Example
   * ```javascript
   * const sheet = new Sheet('Sheet Name');
   * 
   * sheet.appendRows([
   *   ["1-1", "1-2", "1-3"],
   *   ["2-1", "2-2", "2-3"],
   *   ["3-1", "3-2", "3-3"]
   * ]);
   * ```
   * @param {Array} rowsContents 
   */
  appendRows(rowsContents) {
    if (!Array.isArray(rowsContents))
      throw new Error();

    if (!rowsContents.every(item => Array.isArray(item)))
      throw new Error();

    const numRows = rowsContents.length;

    if (!(rowsContents.length > 0))
      throw new Error();

    const numColumns = rowsContents[0].length;

    if (!(numColumns > 0))
      throw new Error();

    const lastRow = this._sheet?.getLastRow();
    const rowPosition = lastRow + 1;

    const range = this.getRange(rowPosition, 1, numRows, numColumns);

    range.setValues(values);

    return this;
  }



  /**
   * Вызывается при преобразовании объекта в соответствующее примитивное значение.
   * @param {string} hint Строковый аргумент, который передаёт желаемый тип примитива: `string`, `number` или `default`.
   * @return {string}
   */
  [Symbol.toPrimitive](hint) {
    if (hint !== 'string') {
      return null;
    }

    return this.constructor.name;
  }



  /**
   * Возвращает значение текущего объекта.
   * @return {string}
   */
  valueOf() {
    return (this[Symbol.toPrimitive] ? this[Symbol.toPrimitive]() : this.constructor.name);
  }



  /**
   * Геттер для получения строки, представляющей тег объекта.
   * @return {string} Имя класса текущего объекта, используемое в `Object.prototype.toString`.
   */
  get [Symbol.toStringTag]() {
    return this.constructor.name;
  }



  /**
   * Возвращает строку, представляющую объект.
   * @return {string}
   */
  toString() {
    return (this[Symbol.toPrimitive] ? this[Symbol.toPrimitive]('string') : this.constructor.name);
  }

}





/**
 * Конструктор 'Cell' - представляет собой объект для работы с ячейкой листа.
 * @class               Cell
 * @memberof            Sheet
 * @version             1.2.0
 */
Sheet.Cell = class Cell {

  /**
   * @param {Integer} rowIndex 
   * @param {Integer} colIndex 
   * @param {*} value 
   */
  constructor(rowIndex, colIndex, value) {
    this.rowIndex = rowIndex;
    this.colIndex = colIndex;
    this.value = value;
  }



  /**
   * Вызывается при преобразовании объекта в соответствующее примитивное значение.
   * @param {string} hint Строковый аргумент, который передаёт желаемый тип примитива: `string`, `number` или `default`.
   * @return {string}
   */
  [Symbol.toPrimitive](hint) {
    return (this.value ?? null);
  }



  /**
   * Возвращает значение текущего объекта.
   * @return {string}
   */
  valueOf() {
    return (this.value ?? null);
  }



  /**
   * Геттер для получения строки, представляющей тег объекта.
   * @return {string} Имя класса текущего объекта, используемое в `Object.prototype.toString`.
   */
  get [Symbol.toStringTag]() {
    return this.constructor.name;
  }



  /**
   * Возвращает строку, представляющую объект.
   * @return {string}
   */
  toString() {
    return (this[Symbol.toPrimitive] ? this[Symbol.toPrimitive]('string') : this.constructor.name);
  }

};





/**
 * Регулярные выражения.
 * @readonly
 * @enum {RegExp}
 */
Sheet.RegExp = {};

Sheet.RegExp.A1NOTATION = /^(?:(?<startColumnLabel>[A-Z]*)(?<startRowPosition>\d*)?(?::(?<endColumnLabel>[A-Z]*)(?<endRowPosition>\d*)?)?)$/i;
