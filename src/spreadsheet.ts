/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * VANILLA SPREADSHEET ENGINE
 * 
 * A zero-dependency spreadsheet application featuring:
 * - Decoupled in-memory state management (Map-based)
 * - Recursive Descent Parser for formula evaluation
 * - Directed Acyclic Graph (DAG) for reactive dependency tracking
 * - BFS-based recalculation with cycle detection
 * - Native column/row resizing via CSS variables
 * - Multi-sheet support
 * - SpreadsheetML (XML 2003) Export
 */

import * as XLSX from 'xlsx';
import * as d3 from 'd3';
import { createIcons, FilePlus, Upload, Download, FileText, Code, Undo2, Redo2, Copy, ClipboardPaste, ClipboardList, Paintbrush, Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, AlignStartVertical, AlignCenterVertical, AlignEndVertical, WrapText, Combine, Rows, Columns, Search, PlusSquare, Edit3, Link, Table, LineChart, BarChart, Sigma, SortAsc, SortDesc, Filter, CheckCircle, Lock, Unlock, Grid, LocateFixed, Keyboard, HelpCircle, PaintBucket, Type, X, Plus, Printer, Square, ListPlus, LayoutGrid, Maximize } from 'lucide';

// --- TYPES & INTERFACES ---

interface CellFormat {
    bold: boolean;
    italic: boolean;
    underline: boolean;
    backgroundColor: string;
    color: string;
    textAlign: 'left' | 'center' | 'right';
    verticalAlign: 'top' | 'middle' | 'bottom';
    type: 'text' | 'number' | 'currency' | 'percentage' | 'date';
    wrapText?: boolean;
    borderTop?: string;
    borderBottom?: string;
    borderLeft?: string;
    borderRight?: string;
}

interface ConditionalRule {
    operator: 'gt' | 'lt' | 'eq' | 'contains';
    value: string | number;
    format: Partial<CellFormat>;
}

interface DataValidation {
    type: 'list' | 'number' | 'date' | 'checkbox';
    allowEmpty: boolean;
    criteria?: any; // For list, it's string[]
    min?: number;
    max?: number;
}

interface CellData {
    id: string;
    formula: string;
    rawValue: string;
    computedValue: any;
    format: CellFormat;
    conditionalRules: ConditionalRule[];
    precedents: Set<string>;
    dependents: Set<string>;
    comment?: string;
    validation?: DataValidation;
}

interface NamedRange {
    name: string;
    range: string; // "Sheet1!A1:B10"
    id: string;
}

type SheetState = Map<string, CellData>;

interface Table {
    id: string;
    range: string; // e.g. "A1:C10"
    hasHeader: boolean;
    style: 'blue' | 'green' | 'orange';
    filters: Record<number, string[]>; // relative colIndex -> allowedValues
    sort?: { colIndex: number, direction: 'asc' | 'desc' }; // relative colIndex
}

interface ChartData {
    id: string;
    type: 'bar' | 'line' | 'pie' | 'scatter';
    range: string;
    title: string;
    x: number;
    y: number;
    width: number;
    height: number;
}

interface SheetMetadata {
    frozenRows: number;
    frozenCols: number;
    tables: Table[];
    charts: ChartData[];
    namedRanges?: NamedRange[];
    theme?: 'light' | 'dark' | 'sepia' | 'ocean';
}

interface WorkbookState {
    sheets: Map<string, SheetState>;
    activeSheet: string;
    metadata: Map<string, SheetMetadata>;
}

// --- CONSTANTS ---

let DEFAULT_ROWS = 100;
let DEFAULT_COLS = 26;

function getColName(n: number): string {
    let name = '';
    while (n >= 0) {
        name = String.fromCharCode((n % 26) + 65) + name;
        n = Math.floor(n / 26) - 1;
    }
    return name;
}

function colToIdx(name: string): number {
    let idx = 0;
    for (let i = 0; i < name.length; i++) {
        idx = idx * 26 + (name.charCodeAt(i) - 64);
    }
    return idx - 1;
}

const COL_NAMES = Array.from({ length: 100 }, (_, i) => getColName(i));

const FUNCTION_DETAILS: Record<string, { sig: string; desc: string }> = {
    'SUM': { sig: 'SUM(value1, [value2, ...])', desc: 'Adds all the numbers in a range of cells.' },
    'AVERAGE': { sig: 'AVERAGE(value1, [value2, ...])', desc: 'Returns the average (arithmetic mean) of the arguments.' },
    'MIN': { sig: 'MIN(value1, [value2, ...])', desc: 'Returns the minimum value in a list of arguments.' },
    'MAX': { sig: 'MAX(value1, [value2, ...])', desc: 'Returns the maximum value in a list of arguments.' },
    'COUNT': { sig: 'COUNT(value1, [value2, ...])', desc: 'Counts the number of cells that contain numbers.' },
    'IF': { sig: 'IF(logical_expression, value_if_true, value_if_false)', desc: 'Returns one value if a logical expression is `TRUE` and another if it is `FALSE`.' },
    'AND': { sig: 'AND(logical_expression1, [logical_expression2, ...])', desc: 'Returns `TRUE` if all the provided arguments are logically true, and `FALSE` if any of the provided arguments are logically false.' },
    'OR': { sig: 'OR(logical_expression1, [logical_expression2, ...])', desc: 'Returns `TRUE` if any of the provided arguments are logically true, and `FALSE` if all of the provided arguments are logically false.' },
    'VLOOKUP': { sig: 'VLOOKUP(search_key, range, index, [is_sorted])', desc: 'Vertical lookup. Searches down the first column of a range for a key and returns the value of a specified cell in the row found.' },
    'SPARKLINE': { sig: 'SPARKLINE(data, [options])', desc: 'Creates a miniature chart contained within a single cell.' },
    'CONCATENATE': { sig: 'CONCATENATE(string1, [string2, ...])', desc: 'Appends strings to one another.' },
    'UPPER': { sig: 'UPPER(text)', desc: 'Converts a specified string to uppercase.' },
    'LOWER': { sig: 'LOWER(text)', desc: 'Converts a specified string to lowercase.' },
    'TRIM': { sig: 'TRIM(text)', desc: 'Removes leading, trailing, and repeated spaces in text.' },
    'NOW': { sig: 'NOW()', desc: 'Returns the current date and time as a date value.' },
    'TODAY': { sig: 'TODAY()', desc: 'Returns the current date as a date value.' }
};

const FUNCTIONS = Object.keys(FUNCTION_DETAILS);

// --- STATE MANAGEMENT ---

class SpreadsheetStore {
    private state: WorkbookState;
    private undoStack: WorkbookState[] = [];
    private redoStack: WorkbookState[] = [];
    private maxHistory = 50;

    constructor() {
        this.state = {
            sheets: new Map(),
            activeSheet: 'Sheet1',
            metadata: new Map()
        };
        this.addSheet('Sheet1');
    }

    saveHistory() {
        this.undoStack.push(this.cloneState(this.state));
        if (this.undoStack.length > this.maxHistory) {
            this.undoStack.shift();
        }
        this.redoStack = []; // Clear redo on new action
    }

    undo() {
        if (this.undoStack.length === 0) return false;
        this.redoStack.push(this.cloneState(this.state));
        this.state = this.undoStack.pop()!;
        return true;
    }

    redo() {
        if (this.redoStack.length === 0) return false;
        this.undoStack.push(this.cloneState(this.state));
        this.state = this.redoStack.pop()!;
        return true;
    }

    private cloneState(state: WorkbookState): WorkbookState {
        const newState: WorkbookState = {
            sheets: new Map(),
            activeSheet: state.activeSheet,
            metadata: new Map()
        };
        state.sheets.forEach((sheet, name) => {
            const newSheet: SheetState = new Map();
            sheet.forEach((cell, id) => {
                newSheet.set(id, {
                    ...cell,
                    precedents: new Set(cell.precedents),
                    dependents: new Set(cell.dependents),
                    format: { ...cell.format },
                    conditionalRules: cell.conditionalRules.map(r => ({
                        ...r,
                        format: { ...r.format }
                    }))
                });
            });
            newState.sheets.set(name, newSheet);
        });
        state.metadata.forEach((meta, name) => {
            newState.metadata.set(name, { 
                ...meta,
                tables: meta.tables.map(t => ({
                    ...t,
                    filters: { ...t.filters },
                    sort: t.sort ? { ...t.sort } : undefined
                })),
                charts: meta.charts.map(c => ({ ...c }))
            });
        });
        return newState;
    }

    addSheet(name?: string) {
        let sheetName = name || `Sheet${this.state.sheets.size + 1}`;
        let counter = 1;
        while (this.state.sheets.has(sheetName)) {
            sheetName = `${name || 'Sheet'}${this.state.sheets.size + counter}`;
            counter++;
        }
        
        this.state.sheets.set(sheetName, new Map());
        this.state.metadata.set(sheetName, { frozenRows: 0, frozenCols: 0, tables: [], charts: [] });
        if (!this.state.activeSheet) this.state.activeSheet = sheetName;
        return sheetName;
    }

    renameSheet(oldName: string, newName: string) {
        if (oldName === newName) return true;
        if (this.state.sheets.has(newName)) {
            alert('A sheet with that name already exists.');
            return false;
        }

        this.saveHistory();

        const sheetData = this.state.sheets.get(oldName)!;
        this.state.sheets.delete(oldName);
        this.state.sheets.set(newName, sheetData);

        const metadata = this.state.metadata.get(oldName)!;
        this.state.metadata.delete(oldName);
        this.state.metadata.set(newName, metadata);

        if (this.state.activeSheet === oldName) {
            this.state.activeSheet = newName;
        }

        // Update formulas in all sheets
        this.state.sheets.forEach((sheet) => {
            sheet.forEach(cell => {
                if (cell.formula) {
                    cell.formula = this.replaceSheetNameInFormula(cell.formula, oldName, newName);
                }
            });
        });

        return true;
    }

    private replaceSheetNameInFormula(formula: string, oldName: string, newName: string): string {
        // Handle both quoted and unquoted sheet names
        const oldQuoted = oldName.includes(' ') ? `'${oldName}'` : oldName;
        const newQuoted = newName.includes(' ') ? `'${newName}'` : newName;

        // Regex to find sheet references: 'Sheet Name'!A1 or SheetName!A1
        // We need to be careful not to replace parts of other sheet names
        const escapedOld = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const escapedOldQuoted = oldQuoted.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        
        const regex = new RegExp(`(${escapedOldQuoted}|${escapedOld})!`, 'g');
        return formula.replace(regex, `${newQuoted}!`);
    }

    getActiveSheetName() {
        return this.state.activeSheet;
    }

    getSheetNames() {
        return Array.from(this.state.sheets.keys());
    }

    setActiveSheet(name: string) {
        if (this.state.sheets.has(name)) {
            this.state.activeSheet = name;
        }
    }

    getSheets() {
        return Array.from(this.state.sheets.keys());
    }

    getDimensions() {
        return { rows: DEFAULT_ROWS, cols: DEFAULT_COLS };
    }

    insertRow(sheetName: string, rowIndex: number, count: number = 1) {
        this.saveHistory();
        const sheet = this.state.sheets.get(sheetName)!;
        const newSheet: SheetState = new Map();

        sheet.forEach((cell, id) => {
            const col = id.match(/[A-Z]+/)?.[0]!;
            const row = parseInt(id.match(/[0-9]+/)?.[0]!);
            
            if (row >= rowIndex) {
                const newId = `${col}${row + count}`;
                newSheet.set(newId, { ...cell, id: newId });
            } else {
                newSheet.set(id, cell);
            }
        });

        this.state.sheets.set(sheetName, newSheet);
        this.updateAllFormulas(sheetName, 'row', rowIndex, count);
        DEFAULT_ROWS += count;
    }

    insertCol(sheetName: string, colIndex: number, count: number = 1) {
        this.saveHistory();
        const sheet = this.state.sheets.get(sheetName)!;
        const newSheet: SheetState = new Map();

        sheet.forEach((cell, id) => {
            const col = id.match(/[A-Z]+/)?.[0]!;
            const row = parseInt(id.match(/[0-9]+/)?.[0]!);
            const cIdx = colToIdx(col);
            
            if (cIdx >= colIndex) {
                const newCol = getColName(cIdx + count);
                const newId = `${newCol}${row}`;
                newSheet.set(newId, { ...cell, id: newId });
            } else {
                newSheet.set(id, cell);
            }
        });

        this.state.sheets.set(sheetName, newSheet);
        this.updateAllFormulas(sheetName, 'col', colIndex, count);
        DEFAULT_COLS += count;
        while (DEFAULT_COLS > COL_NAMES.length) {
            COL_NAMES.push(getColName(COL_NAMES.length));
        }
    }

    private updateAllFormulas(targetSheetName: string, type: 'row' | 'col', index: number, amount: number) {
        this.state.sheets.forEach((sheet, sheetName) => {
            sheet.forEach(cell => {
                if (cell.formula) {
                    cell.formula = this.shiftFormula(cell.formula, sheetName, targetSheetName, type, index, amount);
                }
            });
        });
    }

    private shiftFormula(formula: string, formulaSheet: string, targetSheet: string, type: 'row' | 'col', index: number, amount: number): string {
        const refRegex = /(([A-Za-z0-9]+)!)?(\$?[A-Z]+)(\$?[0-9]+)/g;
        return formula.replace(refRegex, (match, fullSheetPart, sheetName, colPart, rowPart) => {
            const refSheet = sheetName || formulaSheet;
            if (refSheet !== targetSheet) return match;

            const isColAbs = colPart.includes('$');
            const isRowAbs = rowPart.includes('$');
            const col = colPart.replace('$', '');
            const row = parseInt(rowPart.replace('$', ''));

            if (type === 'row') {
                if (!isRowAbs && row >= index) {
                    return (fullSheetPart || '') + colPart + (row + amount);
                }
            } else {
                if (!isColAbs) {
                    const cIdx = colToIdx(col);
                    if (cIdx >= index) {
                        return (fullSheetPart || '') + getColName(cIdx + amount) + rowPart;
                    }
                }
            }
            return match;
        });
    }

    getCell(sheetName: string, cellId: string): CellData {
        const sheet = this.state.sheets.get(sheetName);
        if (!sheet) throw new Error(`Sheet ${sheetName} not found`);

        if (!sheet.has(cellId)) {
            sheet.set(cellId, {
                id: cellId,
                formula: '',
                rawValue: '',
                computedValue: '',
                format: {
                    bold: false,
                    italic: false,
                    underline: false,
                    backgroundColor: '#ffffff',
                    color: '#000000',
                    textAlign: 'left',
                    verticalAlign: 'middle',
                    type: 'text',
                    borderTop: '',
                    borderBottom: '',
                    borderLeft: '',
                    borderRight: ''
                },
                conditionalRules: [],
                precedents: new Set(),
                dependents: new Set()
            });
        }
        return sheet.get(cellId)!;
    }

    updateCell(sheetName: string, cellId: string, updates: Partial<CellData>) {
        const cell = this.getCell(sheetName, cellId);
        Object.assign(cell, updates);
    }

    getWorkbookData() {
        return this.state.sheets;
    }

    getMetadata(sheetName: string): SheetMetadata {
        if (!this.state.metadata.has(sheetName)) {
            this.state.metadata.set(sheetName, { frozenRows: 0, frozenCols: 0, tables: [], charts: [] });
        }
        return this.state.metadata.get(sheetName)!;
    }

    setMetadata(sheetName: string, updates: Partial<SheetMetadata>) {
        const meta = this.getMetadata(sheetName);
        Object.assign(meta, updates);
    }

    serialize(): string {
        const data = {
            activeSheet: this.state.activeSheet,
            sheets: Array.from(this.state.sheets.entries()).map(([name, sheet]) => [
                name,
                Array.from(sheet.entries()).map(([id, cell]) => [
                    id,
                    {
                        ...cell,
                        precedents: Array.from(cell.precedents),
                        dependents: Array.from(cell.dependents)
                    }
                ])
            ]),
            metadata: Array.from(this.state.metadata.entries()),
            dimensions: { rows: DEFAULT_ROWS, cols: DEFAULT_COLS }
        };
        return JSON.stringify(data);
    }

    deserialize(json: string) {
        try {
            const data = JSON.parse(json);
            if (!data.sheets || !data.activeSheet) return false;

            this.state.activeSheet = data.activeSheet;
            this.state.sheets = new Map(data.sheets.map(([name, sheetArr]: any) => [
                name,
                new Map(sheetArr.map(([id, cell]: any) => [
                    id,
                    {
                        ...cell,
                        precedents: new Set(cell.precedents),
                        dependents: new Set(cell.dependents)
                    }
                ]))
            ]));
            this.state.metadata = new Map(data.metadata.map(([name, meta]: any) => [
                name,
                {
                    ...meta,
                    charts: meta.charts || []
                }
            ]));
            
            if (data.dimensions) {
                DEFAULT_ROWS = data.dimensions.rows;
                DEFAULT_COLS = data.dimensions.cols;
                // Ensure COL_NAMES is large enough
                while (COL_NAMES.length < DEFAULT_COLS) {
                    COL_NAMES.push(getColName(COL_NAMES.length));
                }
            }

            this.undoStack = [];
            this.redoStack = [];
            return true;
        } catch (e) {
            console.error('Failed to deserialize state', e);
            return false;
        }
    }
}

// --- FORMULA PARSER ---

class FormulaParser {
    private store: SpreadsheetStore;
    private currentSheet: string;

    constructor(store: SpreadsheetStore, currentSheet: string) {
        this.store = store;
        this.currentSheet = currentSheet;
    }

    parse(formula: string): any {
        if (!formula.startsWith('=')) return formula;
        
        const expression = formula.substring(1).toUpperCase();
        try {
            const tokens = this.tokenize(expression);
            const result = this.evaluate(tokens);
            return result;
        } catch (e) {
            console.error('Formula Error:', e);
            return '#ERROR!';
        }
    }

    private tokenize(str: string): string[] {
        const tokens: string[] = [];
        let i = 0;
        while (i < str.length) {
            const char = str[i];
            if (/\s/.test(char)) { i++; continue; }
            
            if (char === "'") {
                let ref = "'";
                i++;
                while (i < str.length && str[i] !== "'") {
                    ref += str[i++];
                }
                if (i < str.length) ref += str[i++]; // add the closing '
                while (i < str.length && /[A-Z0-9!:]/.test(str[i])) {
                    ref += str[i++];
                }
                tokens.push(ref);
            } else if (/[A-Z]/.test(char)) {
                let ref = '';
                while (i < str.length && /[A-Z0-9!:]/.test(str[i])) {
                    ref += str[i++];
                }
                tokens.push(ref);
            } else if (/[0-9.]/.test(char)) {
                let num = '';
                while (i < str.length && /[0-9.]/.test(str[i])) {
                    num += str[i++];
                }
                tokens.push(num);
            } else if ('+-*/^(),'.includes(char)) {
                tokens.push(char);
                i++;
            } else if (char === '"') {
                let s = '';
                i++;
                while (i < str.length && str[i] !== '"') {
                    s += str[i++];
                }
                tokens.push(`"${s}"`);
                i++;
            } else {
                i++;
            }
        }
        return tokens;
    }

    private evaluate(tokens: string[]): any {
        let pos = 0;

        const parseExpression = (): any => {
            let left = parseTerm();
            while (pos < tokens.length && (tokens[pos] === '+' || tokens[pos] === '-')) {
                const op = tokens[pos++];
                const right = parseTerm();
                if (op === '+') left = Number(left) + Number(right);
                else left = Number(left) - Number(right);
            }
            return left;
        };

        const parseTerm = (): any => {
            let left = parseFactor();
            while (pos < tokens.length && (tokens[pos] === '*' || tokens[pos] === '/')) {
                const op = tokens[pos++];
                const right = parseFactor();
                if (op === '*') left = Number(left) * Number(right);
                else left = Number(left) / Number(right);
            }
            return left;
        };

        const parseFactor = (): any => {
            const token = tokens[pos++];
            if (token === '(') {
                const val = parseExpression();
                pos++; // skip ')'
                return val;
            }
            if (token.startsWith('"')) return token.slice(1, -1);
            if (!isNaN(Number(token))) return Number(token);
            
            // Range detection
            if (token.includes(':') && !token.startsWith("'")) return token;
            
            // Function or Reference
            if (pos < tokens.length && tokens[pos] === '(') {
                const funcName = token;
                pos++; // skip '('
                const args: any[] = [];
                if (tokens[pos] !== ')') {
                    while (true) {
                        args.push(parseExpression());
                        if (tokens[pos] === ',') pos++;
                        else break;
                    }
                }
                pos++; // skip ')'
                return this.executeFunction(funcName, args);
            }

            return this.resolveReference(token);
        };

        return parseExpression();
    }

    private resolveReference(ref: string): any {
        // Try named range first
        const metadata = this.store.getMetadata(this.currentSheet);
        const namedRange = metadata.namedRanges?.find(nr => nr.name.toUpperCase() === ref.toUpperCase());
        if (namedRange) {
            if (namedRange.range.includes(':')) {
                return this.resolveRange(namedRange.range);
            } else {
                return this.resolveReference(namedRange.range);
            }
        }

        let sheetName = this.currentSheet;
        let cellId = ref;

        if (ref.includes('!')) {
            const parts = ref.split('!');
            sheetName = parts[0];
            cellId = parts[1];
            if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
                sheetName = sheetName.slice(1, -1);
            }
        }

        const cell = this.store.getCell(sheetName, cellId);
        return cell.computedValue;
    }

    private executeFunction(name: string, args: any[]): any {
        const flatArgs = this.flattenArgs(args);
        switch (name) {
            case 'SUM': return flatArgs.reduce((a, b) => Number(a) + (isNaN(Number(b)) ? 0 : Number(b)), 0);
            case 'AVERAGE': {
                const nums = flatArgs.map(Number).filter(n => !isNaN(n));
                return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
            }
            case 'MIN': return Math.min(...flatArgs.map(Number).filter(n => !isNaN(n)));
            case 'MAX': return Math.max(...flatArgs.map(Number).filter(n => !isNaN(n)));
            case 'COUNT': return flatArgs.filter(v => !isNaN(Number(v)) && v !== '' && v !== null).length;
            case 'IF': return args[0] ? args[1] : args[2];
            case 'IFS': {
                for (let i = 0; i < args.length; i += 2) {
                    if (args[i]) return args[i + 1];
                }
                return '#N/A';
            }
            case 'AND': return args.every(Boolean);
            case 'OR': return args.some(Boolean);
            case 'VLOOKUP': return this.vlookup(args);
            case 'XLOOKUP': return this.xlookup(args);
            case 'MATCH': return this.match(args);
            case 'INDEX': return this.index(args);
            case 'DATE': return new Date(args[0], args[1] - 1, args[2]).toLocaleDateString();
            case 'TIME': return `${args[0]}:${args[1]}:${args[2]}`;
            case 'NETWORKDAYS': return this.networkdays(args[0], args[1], args[2]);
            case 'CONCATENATE': return flatArgs.join('');
            case 'UPPER': return String(flatArgs[0]).toUpperCase();
            case 'LOWER': return String(flatArgs[0]).toLowerCase();
            case 'TRIM': return String(flatArgs[0]).trim();
            case 'NOW': return new Date().toLocaleString();
            case 'TODAY': return new Date().toLocaleDateString();
            case 'SPARKLINE': {
                const data = this.flattenArgs([args[0]]);
                return { type: 'sparkline', data: data.map(Number).filter(n => !isNaN(n)) };
            }
            default: return '#NAME?';
        }
    }

    private flattenArgs(args: any[]): any[] {
        const result: any[] = [];
        args.forEach(arg => {
            if (Array.isArray(arg)) {
                result.push(...arg);
            } else if (typeof arg === 'string' && arg.includes(':')) {
                result.push(...this.resolveRange(arg));
            } else {
                result.push(arg);
            }
        });
        return result;
    }

    private resolveRange(range: string): any[] {
        let sheetPrefix = '';
        let rangePart = range;
        if (range.includes('!')) {
            const parts = range.split('!');
            sheetPrefix = parts[0] + '!';
            rangePart = parts[1];
        }

        const [start, end] = rangePart.split(':');
        const startCol = start.match(/[A-Z]+/)?.[0]!;
        const startRow = parseInt(start.match(/[0-9]+/)?.[0]!);
        const endCol = end.match(/[A-Z]+/)?.[0]!;
        const endRow = parseInt(end.match(/[0-9]+/)?.[0]!);

        const startColIdx = COL_NAMES.indexOf(startCol);
        const endColIdx = COL_NAMES.indexOf(endCol);

        const values: any[] = [];
        for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
            for (let c = Math.min(startColIdx, endColIdx); c <= Math.max(startColIdx, endColIdx); c++) {
                values.push(this.resolveReference(`${sheetPrefix}${COL_NAMES[c]}${r}`));
            }
        }
        return values;
    }

    private vlookup(args: any[]): any {
        const [lookupValue, rangeStr, colIndex, exactMatch] = args;
        if (!rangeStr || typeof rangeStr !== 'string' || !rangeStr.includes(':')) return '#VALUE!';
        
        const grid = this.getRangeGrid(rangeStr);
        if (!grid) return '#VALUE!';

        const colIdx = Number(colIndex) - 1;
        if (colIdx < 0 || colIdx >= grid[0].length) return '#REF!';

        for (let r = 0; r < grid.length; r++) {
            if (grid[r][0] == lookupValue) {
                return grid[r][colIdx];
            }
        }
        return '#N/A';
    }

    private xlookup(args: any[]): any {
        const [lookupValue, lookupRange, returnRange, ifNotFound] = args;
        const lookupArray = this.flattenArgs([lookupRange]);
        const returnArray = this.flattenArgs([returnRange]);

        const idx = lookupArray.indexOf(lookupValue);
        if (idx !== -1) return returnArray[idx];
        return ifNotFound || '#N/A';
    }

    private match(args: any[]): any {
        const [lookupValue, rangeStr, matchType] = args;
        const arr = this.flattenArgs([rangeStr]);
        const idx = arr.indexOf(lookupValue);
        return idx !== -1 ? idx + 1 : '#N/A';
    }

    private index(args: any[]): any {
        const [rangeStr, rowNum, colNum] = args;
        const grid = this.getRangeGrid(rangeStr);
        if (!grid) return '#VALUE!';
        
        const r = Number(rowNum) - 1;
        const c = (colNum ? Number(colNum) : 1) - 1;
        
        if (r < 0 || r >= grid.length || c < 0 || c >= grid[0].length) return '#REF!';
        return grid[r][c];
    }

    private networkdays(start: any, end: any, holidays: any[] = []): number {
        const d1 = new Date(start);
        const d2 = new Date(end);
        let count = 0;
        const cur = new Date(d1);
        while (cur <= d2) {
            const day = cur.getDay();
            if (day !== 0 && day !== 6) count++;
            cur.setDate(cur.getDate() + 1);
        }
        return count;
    }

    private getRangeGrid(range: string): any[][] | null {
        // Try named range
        const metadata = this.store.getMetadata(this.currentSheet);
        const namedRange = metadata.namedRanges?.find(nr => nr.name.toUpperCase() === range.toUpperCase());
        if (namedRange) range = namedRange.range;

        let sheetPrefix = '';
        let rangePart = range;
        if (range.includes('!')) {
            const parts = range.split('!');
            sheetPrefix = parts[0] + '!';
            rangePart = parts[1];
        }

        const [start, end] = rangePart.split(':');
        const startCol = start.match(/[A-Z]+/)?.[0]!;
        const startRow = parseInt(start.match(/[0-9]+/)?.[0]!);
        const endCol = end.match(/[A-Z]+/)?.[0]!;
        const endRow = parseInt(end.match(/[0-9]+/)?.[0]!);

        const startColIdx = COL_NAMES.indexOf(startCol);
        const endColIdx = COL_NAMES.indexOf(endCol);

        const grid: any[][] = [];
        for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
            const row: any[] = [];
            for (let c = Math.min(startColIdx, endColIdx); c <= Math.max(startColIdx, endColIdx); c++) {
                row.push(this.resolveReference(`${sheetPrefix}${COL_NAMES[c]}${r}`));
            }
            grid.push(row);
        }
        return grid;
    }

    public extractReferences(formula: string): string[] {
        if (!formula.startsWith('=')) return [];
        const tokens = this.tokenize(formula.substring(1).toUpperCase());
        const refs = new Set<string>();
        
        tokens.forEach(t => {
            if (t.includes(':')) {
                const [start, end] = t.split(':');
                const startCol = start.match(/[A-Z]+/)?.[0]!;
                const startRow = parseInt(start.match(/[0-9]+/)?.[0]!);
                const endCol = end.match(/[A-Z]+/)?.[0]!;
                const endRow = parseInt(end.match(/[0-9]+/)?.[0]!);

                const startColIdx = COL_NAMES.indexOf(startCol);
                const endColIdx = COL_NAMES.indexOf(endCol);

                for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
                    for (let c = Math.min(startColIdx, endColIdx); c <= Math.max(startColIdx, endColIdx); c++) {
                        refs.add(`${COL_NAMES[c]}${r}`);
                    }
                }
            } else if (t.includes('!')) {
                refs.add(t);
            } else if (/[A-Z]+[0-9]+/.test(t)) {
                refs.add(t);
            }
        });
        return Array.from(refs);
    }
}

// --- ENGINE & UI ---

export class SpreadsheetEngine {
    private store: SpreadsheetStore;
    private container: HTMLElement;
    private activeCell: string | null = null;
    private selectedCells: Set<string> = new Set();
    private selectionAnchor: string | null = null;
    private isSelecting: boolean = false;
    private lastSelectionWasRange: boolean = false;
    private isEditing: boolean = false;
    private isFormatPainterActive: boolean = false;
    private formatPainterSource: CellFormat | null = null;
    private formulaBar: HTMLInputElement;
    private cellAddress: HTMLElement;
    private grid: HTMLElement;
    private gridContainer: HTMLElement;
    private tabBar: HTMLElement;
    private selectionIndicator: HTMLElement | null = null;
    private suggestionBox: HTMLElement | null = null;
    private previewBox: HTMLElement | null = null;
    private suggestionIndex: number = -1;

    constructor(containerId: string) {
        this.store = new SpreadsheetStore();
        this.container = document.getElementById(containerId)!;
        this.renderLayout();
        
        this.formulaBar = this.container.querySelector('.formula-input') as HTMLInputElement;
        this.cellAddress = this.container.querySelector('.cell-address') as HTMLElement;
        this.grid = this.container.querySelector('.spreadsheet-grid') as HTMLElement;
        this.gridContainer = this.container.querySelector('.grid-container') as HTMLElement;
        this.tabBar = this.container.querySelector('.tab-bar') as HTMLElement;
        this.selectionIndicator = this.container.querySelector('#selection-indicator') as HTMLElement;
        this.suggestionBox = this.container.querySelector('#formula-suggestions') as HTMLElement;
        this.previewBox = this.container.querySelector('#formula-preview') as HTMLElement;

        this.initGrid();
        this.initTabs();
        this.initEventListeners();
        this.checkAutosave();
        this.initAutosave();

        createIcons({
            icons: {
                FilePlus, Upload, Download, FileText, Code, Undo2, Redo2, Copy, ClipboardPaste, ClipboardList, Paintbrush, Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, AlignStartVertical, AlignCenterVertical, AlignEndVertical, WrapText, Combine, Rows, Columns, Search, PlusSquare, Edit3, Link, Table, LineChart, BarChart, Sigma, SortAsc, SortDesc, Filter, CheckCircle, Lock, Unlock, Grid, LocateFixed, Keyboard, HelpCircle, PaintBucket, Type, X, Plus, Printer, Square, ListPlus, LayoutGrid, Maximize
            }
        });

        this.gridContainer.addEventListener('scroll', () => {
            this.updateSelectionIndicator();
        });
    }

    private renderLayout() {
        this.container.innerHTML = `
            <div id="spreadsheet-container">
                <div class="ribbon-container">
                    <div class="ribbon-tabs">
                        <div class="ribbon-tab-item file-tab" data-pane="pane-file">File</div>
                        <div class="ribbon-tab-item active" data-pane="pane-home">Home</div>
                        <div class="ribbon-tab-item" data-pane="pane-insert">Insert</div>
                        <div class="ribbon-tab-item" data-pane="pane-formulas">Formulas</div>
                        <div class="ribbon-tab-item" data-pane="pane-data">Data</div>
                        <div class="ribbon-tab-item" data-pane="pane-review">Review</div>
                        <div class="ribbon-tab-item" data-pane="pane-view">View</div>
                        <div class="ribbon-tab-item" data-pane="pane-page-layout">Page Layout</div>
                        <div class="ribbon-tab-item" data-pane="pane-help">Help</div>
                    </div>
                    <div class="ribbon-content">
                        <!-- FILE PANE -->
                        <div class="ribbon-pane" id="pane-file">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-new" title="New Workbook"><i data-lucide="file-plus"></i></button>
                                    <button class="tool-btn" id="btn-import-excel" title="Import Excel"><i data-lucide="upload"></i></button>
                                    <button class="tool-btn" id="btn-export-excel" title="Export Excel"><i data-lucide="download"></i></button>
                                    <button class="tool-btn" id="btn-export-csv" title="Export CSV"><i data-lucide="file-text"></i></button>
                                    <button class="tool-btn" id="btn-export-json" title="Export JSON"><i data-lucide="code"></i></button>
                                    <button class="tool-btn" id="btn-print-pdf" title="Print to PDF"><i data-lucide="printer"></i></button>
                                    <input type="file" id="input-import-excel" style="display:none" accept=".xlsx,.xls">
                                </div>
                                <div class="ribbon-label">Workbook</div>
                            </div>
                        </div>

                        <!-- HOME PANE -->
                        <div class="ribbon-pane active" id="pane-home">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-undo" title="Undo (Ctrl+Z)"><i data-lucide="undo-2"></i></button>
                                    <button class="tool-btn" id="btn-redo" title="Redo (Ctrl+Y)"><i data-lucide="redo-2"></i></button>
                                </div>
                                <div class="ribbon-label">History</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-copy" title="Copy (Ctrl+C)"><i data-lucide="copy"></i></button>
                                    <button class="tool-btn" id="btn-paste" title="Paste (Ctrl+V)"><i data-lucide="clipboard-paste"></i></button>
                                    <button class="tool-btn" id="btn-paste-special" title="Paste Special"><i data-lucide="clipboard-list"></i></button>
                                    <button class="tool-btn" id="btn-format-painter" title="Format Painter"><i data-lucide="paintbrush"></i></button>
                                </div>
                                <div class="ribbon-label">Clipboard</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-bold" title="Bold (Ctrl+B)"><i data-lucide="bold"></i></button>
                                    <button class="tool-btn" id="btn-italic" title="Italic (Ctrl+I)"><i data-lucide="italic"></i></button>
                                    <button class="tool-btn" id="btn-underline" title="Underline (Ctrl+U)"><i data-lucide="underline"></i></button>
                                    <div class="dropdown" id="dropdown-borders">
                                        <button class="tool-btn" id="btn-borders" title="Borders">
                                            <i data-lucide="square"></i>
                                        </button>
                                        <div class="dropdown-content border-options">
                                            <div data-border="bottom">Bottom Border</div>
                                            <div data-border="top">Top Border</div>
                                            <div data-border="left">Left Border</div>
                                            <div data-border="right">Right Border</div>
                                            <div class="divider"></div>
                                            <div data-border="none">No Border</div>
                                            <div data-border="all">All Borders</div>
                                            <div data-border="outer">Outside Borders</div>
                                        </div>
                                    </div>
                                    <div class="color-picker-wrapper">
                                        <button class="tool-btn" id="btn-fill-color" title="Fill Color">
                                            <i data-lucide="paint-bucket"></i>
                                            <span style="border-bottom: 3px solid #000; position: absolute; bottom: 4px; left: 6px; right: 6px;"></span>
                                        </button>
                                        <input type="color" id="input-fill-color" style="display:none" value="#ffffff">
                                    </div>
                                    <div class="color-picker-wrapper">
                                        <button class="tool-btn" id="btn-text-color" title="Text Color">
                                            <i data-lucide="type"></i>
                                            <span style="border-bottom: 3px solid #000; position: absolute; bottom: 4px; left: 6px; right: 6px;"></span>
                                        </button>
                                        <input type="color" id="input-text-color" style="display:none" value="#000000">
                                    </div>
                                </div>
                                <div class="ribbon-label">Font</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-align-left" title="Align Left"><i data-lucide="align-left"></i></button>
                                    <button class="tool-btn" id="btn-align-center" title="Align Center"><i data-lucide="align-center"></i></button>
                                    <button class="tool-btn" id="btn-align-right" title="Align Right"><i data-lucide="align-right"></i></button>
                                    <div class="divider-v"></div>
                                    <button class="tool-btn" id="btn-align-top" title="Align Top"><i data-lucide="align-start-vertical"></i></button>
                                    <button class="tool-btn" id="btn-align-middle" title="Align Middle"><i data-lucide="align-center-vertical"></i></button>
                                    <button class="tool-btn" id="btn-align-bottom" title="Align Bottom"><i data-lucide="align-end-vertical"></i></button>
                                    <div class="divider-v"></div>
                                    <button class="tool-btn" id="btn-text-wrap" title="Text Wrap"><i data-lucide="wrap-text"></i></button>
                                    <button class="tool-btn" id="btn-merge-cells" title="Merge Cells"><i data-lucide="combine"></i></button>
                                </div>
                                <div class="ribbon-label">Alignment</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <select class="tool-select" id="select-format">
                                        <option value="text">General</option>
                                        <option value="number">Number</option>
                                        <option value="currency">Currency</option>
                                        <option value="percentage">Percentage</option>
                                        <option value="date">Date</option>
                                    </select>
                                </div>
                                <div class="ribbon-label">Number</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-insert-row" title="Insert Row"><i data-lucide="rows"></i></button>
                                    <button class="tool-btn" id="btn-insert-col" title="Insert Column"><i data-lucide="columns"></i></button>
                                    <div class="divider-v"></div>
                                    <button class="tool-btn" id="btn-insert-50-rows" title="Add 50 Rows"><i data-lucide="list-plus"></i></button>
                                    <button class="tool-btn" id="btn-insert-50-cols" title="Add 50 Columns"><i data-lucide="layout-grid"></i></button>
                                    <div class="divider-v"></div>
                                    <div class="dropdown" id="dropdown-format">
                                        <button class="tool-btn" id="btn-format" title="Format">
                                            <i data-lucide="maximize"></i>
                                        </button>
                                        <div class="dropdown-content format-options">
                                            <div id="option-row-height">Row Height...</div>
                                            <div id="option-col-width">Column Width...</div>
                                        </div>
                                    </div>
                                </div>
                                <div class="ribbon-label">Cells</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-find-replace" title="Find & Replace"><i data-lucide="search"></i></button>
                                </div>
                                <div class="ribbon-label">Editing</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-new-sheet" title="New Sheet"><i data-lucide="plus-square"></i></button>
                                    <button class="tool-btn" id="btn-rename-sheet" title="Rename Sheet"><i data-lucide="edit-3"></i></button>
                                </div>
                                <div class="ribbon-label">Sheet</div>
                            </div>
                        </div>

                        <!-- INSERT PANE -->
                        <div class="ribbon-pane" id="pane-insert">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-link-cell" title="Link Cell"><i data-lucide="link"></i></button>
                                    <button class="tool-btn" id="btn-create-table" title="Create Table"><i data-lucide="table"></i></button>
                                    <button class="tool-btn" id="btn-insert-sparkline" title="Insert Sparkline"><i data-lucide="line-chart"></i></button>
                                    <button class="tool-btn" id="btn-insert-chart" title="Insert Chart"><i data-lucide="bar-chart"></i></button>
                                    <button class="tool-btn" id="btn-pivot-table" title="Pivot Table"><i data-lucide="layout"></i></button>
                                </div>
                                <div class="ribbon-label">Insert</div>
                            </div>
                        </div>

                        <!-- FORMULAS PANE -->
                        <div class="ribbon-pane" id="pane-formulas">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-autosum" title="AutoSum"><i data-lucide="sigma"></i></button>
                                </div>
                                <div class="ribbon-label">Function Library</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-define-name" title="Define Range Name"><i data-lucide="list-plus"></i></button>
                                </div>
                                <div class="ribbon-label">Defined Names</div>
                            </div>
                        </div>

                        <!-- DATA PANE -->
                        <div class="ribbon-pane" id="pane-data">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-sort-asc" title="Sort A-Z"><i data-lucide="sort-asc"></i></button>
                                    <button class="tool-btn" id="btn-sort-desc" title="Sort Z-A"><i data-lucide="sort-desc"></i></button>
                                    <button class="tool-btn" id="btn-filter" title="Filter"><i data-lucide="filter"></i></button>
                                    <button class="tool-btn" id="btn-data-validation" title="Validation"><i data-lucide="check-circle"></i></button>
                                </div>
                                <div class="ribbon-label">Sort & Filter</div>
                            </div>
                        </div>

                        <!-- REVIEW PANE -->
                        <div class="ribbon-pane" id="pane-review">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-new-comment" title="New Comment"><i data-lucide="edit-3"></i></button>
                                    <button class="tool-btn" id="btn-delete-comment" title="Delete Comment"><i data-lucide="x"></i></button>
                                </div>
                                <div class="ribbon-label">Comments</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" disabled title="Protect Sheet"><i data-lucide="lock"></i></button>
                                </div>
                                <div class="ribbon-label">Protection</div>
                            </div>
                        </div>

                        <!-- VIEW PANE -->
                        <div class="ribbon-pane" id="pane-view">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn active" id="btn-toggle-gridlines" title="Toggle Gridlines"><i data-lucide="grid"></i></button>
                                </div>
                                <div class="ribbon-label">Show</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-freeze-panes" title="Freeze Panes"><i data-lucide="locate-fixed"></i></button>
                                    <button class="tool-btn" id="btn-unfreeze-panes" title="Unfreeze Panes"><i data-lucide="unlock"></i></button>
                                </div>
                                <div class="ribbon-label">Window</div>
                            </div>
                        </div>

                        <!-- HELP PANE -->
                        <div class="ribbon-pane" id="pane-help">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-shortcuts" title="Keyboard Shortcuts"><i data-lucide="keyboard"></i></button>
                                    <button class="tool-btn" id="btn-formula-help" title="Formula Help"><i data-lucide="help-circle"></i></button>
                                </div>
                                <div class="ribbon-label">Help</div>
                            </div>
                        </div>

                        <!-- PAGE LAYOUT PANE -->
                        <div class="ribbon-pane" id="pane-page-layout">
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <select class="tool-select" id="select-theme">
                                        <option value="light">Light</option>
                                        <option value="dark">Dark</option>
                                        <option value="sepia">Sepia Theme</option>
                                        <option value="ocean">Ocean Mode</option>
                                    </select>
                                </div>
                                <div class="ribbon-label">Themes</div>
                            </div>
                            <div class="ribbon-group">
                                <div class="toolbar">
                                    <button class="tool-btn" id="btn-print" title="Print to PDF"><i data-lucide="printer"></i></button>
                                </div>
                                <div class="ribbon-label">Page Setup</div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="formula-bar">
                    <div class="cell-address"></div>
                    <input type="text" class="formula-input" placeholder="Formula bar">
                </div>
                <div class="grid-container">
                    <div class="selection-indicator" id="selection-indicator">
                        <div class="selection-handle"></div>
                    </div>
                    <div class="spreadsheet-grid">
                    </div>
                </div>
                <div class="tab-bar">
                    <div class="tabs-list"></div>
                    <div class="add-sheet-btn" title="New Sheet"><i data-lucide="plus"></i></div>
                </div>
                <div id="formula-suggestions" class="formula-suggestions" style="display: none;"></div>
                <div id="formula-preview" class="formula-preview" style="display: none;"></div>
            </div>
        `;
    }

    private initGrid() {
        const fragment = document.createDocumentFragment();
        const sheetName = this.store.getActiveSheetName();
        const metadata = this.store.getMetadata(sheetName);
        
        // Corner
        const corner = document.createElement('div');
        corner.className = 'grid-header corner-header';
        fragment.appendChild(corner);

        // Column Headers
        for (let i = 0; i < DEFAULT_COLS; i++) {
            const colName = COL_NAMES[i];
            const header = document.createElement('div');
            header.className = 'grid-header col-header';
            header.textContent = colName;
            header.dataset.col = colName;
            header.style.gridColumn = (i + 2).toString();
            header.style.gridRow = '1';
            header.style.width = 'var(--col-width-' + colName + ', 100px)';
            
            if (i < metadata.frozenCols) {
                header.style.left = `calc(var(--header-size) + ${this.getColOffset(i)})`;
                header.style.zIndex = '30';
                header.classList.add('frozen-header');
            }

            const resizer = document.createElement('div');
            resizer.className = 'resizer col-resizer';
            resizer.dataset.col = colName;
            header.appendChild(resizer);
            
            fragment.appendChild(header);
        }

        // Rows
        for (let r = 1; r <= DEFAULT_ROWS; r++) {
            const rowHeader = document.createElement('div');
            rowHeader.className = 'grid-header row-header';
            rowHeader.textContent = r.toString();
            rowHeader.dataset.row = r.toString();
            rowHeader.style.gridColumn = '1';
            rowHeader.style.gridRow = (r + 1).toString();
            rowHeader.style.height = 'var(--row-height-' + r + ', 25px)';
            
            if (r <= metadata.frozenRows) {
                rowHeader.style.top = `calc(var(--header-size) + ${this.getRowOffset(r - 1)})`;
                rowHeader.style.zIndex = '30';
                rowHeader.classList.add('frozen-header');
            }

            const resizer = document.createElement('div');
            resizer.className = 'resizer row-resizer';
            resizer.dataset.row = r.toString();
            rowHeader.appendChild(resizer);
            
            fragment.appendChild(rowHeader);

            for (let c = 0; c < DEFAULT_COLS; c++) {
                const cellId = `${COL_NAMES[c]}${r}`;
                const cell = document.createElement('div');
                cell.className = 'grid-cell';
                cell.dataset.id = cellId;
                cell.style.gridColumn = (c + 2).toString();
                cell.style.gridRow = (r + 1).toString();
                cell.style.width = 'var(--col-width-' + COL_NAMES[c] + ', 100px)';
                cell.style.height = 'var(--row-height-' + r + ', 25px)';
                cell.contentEditable = 'false';

                const isFrozenRow = r <= metadata.frozenRows;
                const isFrozenCol = c < metadata.frozenCols;

                if (isFrozenRow || isFrozenCol) {
                    cell.style.position = 'sticky';
                    if (isFrozenRow) {
                        cell.style.top = `calc(var(--header-size) + ${this.getRowOffset(r - 1)})`;
                    }
                    if (isFrozenCol) {
                        cell.style.left = `calc(var(--header-size) + ${this.getColOffset(c)})`;
                    }
                    
                    if (isFrozenRow && isFrozenCol) {
                        cell.style.zIndex = '15';
                    } else {
                        cell.style.zIndex = '5';
                    }
                    cell.classList.add('frozen-cell');
                    if (r === metadata.frozenRows) cell.classList.add('frozen-row-edge');
                    if (c === metadata.frozenCols - 1) cell.classList.add('frozen-col-edge');
                }

                fragment.appendChild(cell);
            }
        }

        this.grid.innerHTML = '';
        this.grid.appendChild(fragment);
        this.updateGridDisplay();
        this.renderCharts();
    }

    private getColOffset(index: number): string {
        if (index === 0) return '0px';
        let parts = [];
        for (let i = 0; i < index; i++) {
            parts.push(`var(--col-width-${COL_NAMES[i]}, 100px)`);
        }
        return `calc(${parts.join(' + ')})`;
    }

    private getRowOffset(index: number): string {
        if (index === 0) return '0px';
        let parts = [];
        for (let i = 0; i < index; i++) {
            parts.push(`var(--row-height-${i + 1}, 25px)`);
        }
        return `calc(${parts.join(' + ')})`;
    }

    private initTabs() {
        const tabsList = this.tabBar.querySelector('.tabs-list')!;
        tabsList.innerHTML = '';
        this.store.getSheets().forEach(sheetName => {
            const tab = document.createElement('div');
            tab.className = `tab ${sheetName === this.store.getActiveSheetName() ? 'active' : ''}`;
            tab.textContent = sheetName;
            tab.title = 'Double-click to rename';
            tab.onclick = () => {
                this.commitActiveEdit();
                this.store.setActiveSheet(sheetName);
                this.initTabs();
                this.initGrid();
            };
            tab.ondblclick = (e) => {
                e.stopPropagation();
                const newName = prompt('Rename sheet to:', sheetName);
                if (newName && newName !== sheetName) {
                    if (this.store.renameSheet(sheetName, newName)) {
                        this.initTabs();
                        this.initGrid();
                    }
                }
            };
            tabsList.appendChild(tab);
        });
    }

    private initEventListeners() {
        // Ribbon Tabs
        this.container.querySelectorAll('.ribbon-tab-item').forEach(tab => {
            tab.addEventListener('click', () => {
                const paneId = (tab as HTMLElement).dataset.pane;
                if (!paneId) return;

                // Update tabs
                this.container.querySelectorAll('.ribbon-tab-item').forEach(t => t.classList.remove('active'));
                tab.classList.add('active');

                // Update panes
                this.container.querySelectorAll('.ribbon-pane').forEach(p => p.classList.remove('active'));
                document.getElementById(paneId)?.classList.add('active');
            });
        });

        this.grid.addEventListener('mouseover', (e) => {
            const target = e.target as HTMLElement;
            if (target.classList.contains('grid-cell') && target.classList.contains('has-comment')) {
                const cell = this.store.getCell(this.store.getActiveSheetName(), target.dataset.id!);
                if (cell.comment) {
                    this.showHoverComment(cell.comment, e);
                }
            }
        });

        this.grid.addEventListener('mouseout', (e) => {
            const target = e.target as HTMLElement;
            if (target.classList.contains('grid-cell')) {
                this.hideHoverComment();
            }
        });

        this.grid.addEventListener('click', (e) => {
            const target = e.target as HTMLElement;
            const cellEl = target.closest('.grid-cell') as HTMLElement;
            if (!cellEl) return;

            const id = cellEl.dataset.id!;
            const cell = this.store.getCell(this.store.getActiveSheetName(), id);

            if (cell.validation) {
                if (cell.validation.type === 'checkbox') {
                    this.store.saveHistory();
                    const newVal = cell.rawValue === 'TRUE' ? 'FALSE' : 'TRUE';
                    cell.rawValue = newVal;
                    cell.computedValue = newVal === 'TRUE';
                    this.updateGridDisplay();
                    this.recalculate(id);
                    return;
                }

                if (cell.validation.type === 'list' && (target.classList.contains('dropdown-indicator') || target === cellEl)) {
                    this.showDropdownMenu(cell, cellEl);
                }
            }
        });

        // Cell Selection & Dragging
        this.grid.addEventListener('mousedown', (e) => {
            const target = e.target as HTMLElement;
            if (this.isEditing && target.classList.contains('editing')) {
                return;
            }

            if (target.classList.contains('grid-cell')) {
                e.preventDefault();
                const id = target.dataset.id!;
                const isMulti = e.ctrlKey || e.metaKey;
                const isRange = e.shiftKey;

                this.isSelecting = true;
                this.selectCell(id, isMulti, isRange);
            } else if (target.classList.contains('col-header')) {
                e.preventDefault();
                const col = target.dataset.col!;
                this.isSelecting = true;
                this.selectColumn(col, e.ctrlKey || e.metaKey, e.shiftKey);
            } else if (target.classList.contains('row-header')) {
                e.preventDefault();
                const row = target.dataset.row!;
                this.isSelecting = true;
                this.selectRow(row, e.ctrlKey || e.metaKey, e.shiftKey);
            } else if (target.classList.contains('corner-header')) {
                e.preventDefault();
                this.selectionAnchor = 'A1';
                this.selectCell(`${COL_NAMES[DEFAULT_COLS - 1]}${DEFAULT_ROWS}`, false, true);
            }
        });

        window.addEventListener('mousemove', (e) => {
            if (!this.isSelecting) return;
            
            const target = document.elementFromPoint(e.clientX, e.clientY) as HTMLElement;
            if (target) {
                if (target.classList.contains('grid-cell')) {
                    const id = target.dataset.id!;
                    if (id !== this.activeCell) {
                        this.selectCell(id, false, true);
                    }
                } else if (target.classList.contains('col-header')) {
                    const col = target.dataset.col!;
                    this.selectColumn(col, false, true);
                } else if (target.classList.contains('row-header')) {
                    const row = target.dataset.row!;
                    this.selectRow(row, false, true);
                }
            }
        });

        window.addEventListener('mouseup', () => {
            if (this.isSelecting && this.isFormatPainterActive && this.formatPainterSource) {
                this.applyFormatPainter();
            }
            this.isSelecting = false;
        });

        // Cell Editing
        this.grid.addEventListener('dblclick', (e) => {
            const target = e.target as HTMLElement;
            if (target.classList.contains('grid-cell')) {
                this.startEditing(target, e as MouseEvent);
            }
        });

        // Formula Bar Input
        this.formulaBar.addEventListener('input', () => {
            this.updateSuggestions(this.formulaBar, this.formulaBar.value);
        });

        this.formulaBar.addEventListener('blur', () => {
            setTimeout(() => {
                if (this.suggestionBox) this.suggestionBox.style.display = 'none';
                if (this.previewBox) this.previewBox.style.display = 'none';
                this.suggestionIndex = -1;
            }, 200);
        });

        this.formulaBar.addEventListener('keydown', (e) => {
            if (this.suggestionBox?.style.display === 'block') {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    const matches = this.getMatches(this.formulaBar.value);
                    this.suggestionIndex = (this.suggestionIndex + 1) % matches.length;
                    this.updateSuggestions(this.formulaBar, this.formulaBar.value);
                    return;
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    const matches = this.getMatches(this.formulaBar.value);
                    this.suggestionIndex = (this.suggestionIndex - 1 + matches.length) % matches.length;
                    this.updateSuggestions(this.formulaBar, this.formulaBar.value);
                    return;
                } else if (e.key === 'Tab' || e.key === 'Enter') {
                    if (this.suggestionIndex >= 0) {
                        e.preventDefault();
                        const matches = this.getMatches(this.formulaBar.value);
                        this.applySuggestion(this.formulaBar, matches[this.suggestionIndex]);
                        return;
                    }
                }
            }

            if (e.key === 'Enter' && this.activeCell) {
                this.commitValue(this.activeCell, this.formulaBar.value);
                this.formulaBar.blur();
                this.grid.focus();
            }
        });

        // Keyboard Navigation
        window.addEventListener('keydown', (e) => {
            const isCtrl = e.ctrlKey || e.metaKey;
            
            // Undo/Redo Shortcuts
            if (isCtrl) {
                if (e.key === 'z') {
                    e.preventDefault();
                    this.handleUndo();
                    return;
                }
                if (e.key === 'y') {
                    e.preventDefault();
                    this.handleRedo();
                    return;
                }
                if (e.key === 'b') {
                    e.preventDefault();
                    this.toggleFormat('bold');
                    return;
                }
                if (e.key === 'i') {
                    e.preventDefault();
                    this.toggleFormat('italic');
                    return;
                }
                if (e.key === 'u') {
                    e.preventDefault();
                    this.toggleFormat('underline');
                    return;
                }
                if (e.key === 'c') {
                    // Let native copy happen if text is selected in an input
                    if (document.activeElement?.tagName !== 'INPUT') {
                        e.preventDefault();
                        this.copyToClipboard();
                    }
                    return;
                }
                if (e.key === 'x') {
                    if (document.activeElement?.tagName !== 'INPUT') {
                        e.preventDefault();
                        this.cutToClipboard();
                    }
                    return;
                }
                if (e.key === 'v') {
                    if (document.activeElement?.tagName !== 'INPUT') {
                        e.preventDefault();
                        this.pasteFromClipboard();
                    }
                    return;
                }
                if (e.key === 'a') {
                    if (document.activeElement?.tagName !== 'INPUT') {
                        e.preventDefault();
                        this.selectedCells.clear();
                        this.selectionAnchor = 'A1';
                        this.selectCell(`${COL_NAMES[DEFAULT_COLS-1]}${DEFAULT_ROWS}`, false, true);
                    }
                    return;
                }
                if (e.key === 'f') {
                    e.preventDefault();
                    this.showFindReplaceDialog();
                    return;
                }
                if (e.key === 'Home') {
                    e.preventDefault();
                    this.selectCell('A1');
                    return;
                }
                if (e.key === 'End') {
                    e.preventDefault();
                    this.selectCell(`${COL_NAMES[DEFAULT_COLS-1]}${DEFAULT_ROWS}`);
                    return;
                }
            }

            if (!this.activeCell || document.activeElement?.tagName === 'INPUT') return;

            const col = this.activeCell.match(/[A-Z]+/)?.[0]!;
            const row = parseInt(this.activeCell.match(/[0-9]+/)?.[0]!);
            let nextCell: string | null = null;

            switch (e.key) {
                case 'ArrowUp': if (row > 1) nextCell = `${col}${row - 1}`; break;
                case 'ArrowDown': if (row < DEFAULT_ROWS) nextCell = `${col}${row + 1}`; break;
                case 'ArrowLeft': 
                    const colIdx = COL_NAMES.indexOf(col);
                    if (colIdx > 0) nextCell = `${COL_NAMES[colIdx - 1]}${row}`;
                    break;
                case 'ArrowRight':
                    const colIdxR = COL_NAMES.indexOf(col);
                    if (colIdxR < DEFAULT_COLS - 1) nextCell = `${COL_NAMES[colIdxR + 1]}${row}`;
                    break;
                case 'Tab':
                    e.preventDefault();
                    const cIdx = COL_NAMES.indexOf(col);
                    if (e.shiftKey) {
                        if (cIdx > 0) nextCell = `${COL_NAMES[cIdx - 1]}${row}`;
                    } else {
                        if (cIdx < DEFAULT_COLS - 1) nextCell = `${COL_NAMES[cIdx + 1]}${row}`;
                    }
                    break;
                case 'Enter':
                    e.preventDefault();
                    if (e.shiftKey) {
                        if (row > 1) nextCell = `${col}${row - 1}`;
                    } else {
                        if (row < DEFAULT_ROWS) nextCell = `${col}${row + 1}`;
                    }
                    break;
                case 'F2':
                    e.preventDefault();
                    const cellElF2 = this.grid.querySelector(`[data-id="${this.activeCell}"]`) as HTMLElement;
                    this.startEditing(cellElF2);
                    break;
                case 'Delete':
                case 'Backspace':
                    e.preventDefault();
                    if (this.selectedCells.size > 0) {
                        this.store.saveHistory();
                        this.selectedCells.forEach(id => this.commitValue(id, '', this.store.getActiveSheetName(), true));
                        this.updateGridDisplay();
                    } else if (this.activeCell) {
                        this.commitValue(this.activeCell, '');
                    }
                    break;
            }

            if (nextCell) {
                e.preventDefault();
                this.selectCell(nextCell, false, e.shiftKey);
            }
        });

        // Toolbar
        document.getElementById('btn-new')?.addEventListener('click', () => {
            if (confirm('Create a new workbook? All unsaved changes will be lost.')) {
                window.location.reload();
            }
        });
        document.getElementById('btn-import-excel')?.addEventListener('click', () => {
            document.getElementById('input-import-excel')?.click();
        });
        document.getElementById('input-import-excel')?.addEventListener('change', (e) => {
            const file = (e.target as HTMLInputElement).files?.[0];
            if (file) this.importFromExcel(file);
        });
        document.getElementById('btn-export-excel')?.addEventListener('click', () => this.exportToExcelXLSX());
        document.getElementById('btn-export-csv')?.addEventListener('click', () => this.exportToCSV());
        document.getElementById('btn-export-json')?.addEventListener('click', () => this.exportToJSON());

        document.getElementById('btn-bold')?.addEventListener('click', () => this.toggleFormat('bold'));
        document.getElementById('btn-italic')?.addEventListener('click', () => this.toggleFormat('italic'));
        document.getElementById('btn-underline')?.addEventListener('click', () => this.toggleFormat('underline'));
        document.getElementById('btn-text-wrap')?.addEventListener('click', () => this.toggleTextWrap());
        document.getElementById('btn-merge-cells')?.addEventListener('click', () => this.mergeCells());
        
        document.getElementById('btn-fill-color')?.addEventListener('click', () => {
            document.getElementById('input-fill-color')?.click();
        });
        document.getElementById('input-fill-color')?.addEventListener('input', (e) => {
            const color = (e.target as HTMLInputElement).value;
            this.toggleFormat('backgroundColor', color);
        });

        document.getElementById('btn-text-color')?.addEventListener('click', () => {
            document.getElementById('input-text-color')?.click();
        });
        document.getElementById('input-text-color')?.addEventListener('input', (e) => {
            const color = (e.target as HTMLInputElement).value;
            this.toggleFormat('color', color);
        });

        document.getElementById('btn-align-left')?.addEventListener('click', () => this.toggleFormat('textAlign', 'left'));
        document.getElementById('btn-align-center')?.addEventListener('click', () => this.toggleFormat('textAlign', 'center'));
        document.getElementById('btn-align-right')?.addEventListener('click', () => this.toggleFormat('textAlign', 'right'));
        document.getElementById('btn-align-top')?.addEventListener('click', () => this.toggleFormat('verticalAlign', 'top'));
        document.getElementById('btn-align-middle')?.addEventListener('click', () => this.toggleFormat('verticalAlign', 'middle'));
        document.getElementById('btn-align-bottom')?.addEventListener('click', () => this.toggleFormat('verticalAlign', 'bottom'));

        document.getElementById('select-format')?.addEventListener('change', (e) => {
            const format = (e.target as HTMLSelectElement).value as any;
            this.toggleFormat('type', format);
        });

        document.getElementById('btn-sort-asc')?.addEventListener('click', () => this.sortRange('asc'));
        document.getElementById('btn-sort-desc')?.addEventListener('click', () => this.sortRange('desc'));
        document.getElementById('btn-filter')?.addEventListener('click', () => this.toggleFilter());
        document.getElementById('btn-data-validation')?.addEventListener('click', () => this.showDataValidationDialog());

        document.getElementById('btn-copy')?.addEventListener('click', () => this.copyToClipboard());
        document.getElementById('btn-paste')?.addEventListener('click', () => this.pasteFromClipboard());
        document.getElementById('btn-paste-special')?.addEventListener('click', () => this.showPasteSpecialDialog());
        document.getElementById('btn-format-painter')?.addEventListener('click', () => this.activateFormatPainter());
        document.getElementById('btn-autosum')?.addEventListener('click', () => this.autoSum());
        document.getElementById('btn-toggle-gridlines')?.addEventListener('click', (e) => {
            const btn = e.target as HTMLElement;
            btn.classList.toggle('active');
            this.grid.classList.toggle('no-gridlines');
        });

        document.getElementById('btn-freeze-panes')?.addEventListener('click', () => {
            this.commitActiveEdit();
            if (!this.activeCell) return;
            const col = this.activeCell.match(/[A-Z]+/)?.[0]!;
            const row = parseInt(this.activeCell.match(/[0-9]+/)?.[0]!);
            const cIdx = COL_NAMES.indexOf(col);
            
            this.store.setMetadata(this.store.getActiveSheetName(), {
                frozenRows: row - 1,
                frozenCols: cIdx
            });
            this.initGrid();
        });

        document.getElementById('btn-unfreeze-panes')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.store.setMetadata(this.store.getActiveSheetName(), {
                frozenRows: 0,
                frozenCols: 0
            });
            this.initGrid();
        });

        document.getElementById('btn-insert-row')?.addEventListener('click', () => {
            this.commitActiveEdit();
            if (!this.activeCell) return;
            const row = parseInt(this.activeCell.match(/[0-9]+/)?.[0]!);
            this.store.insertRow(this.store.getActiveSheetName(), row);
            this.initGrid();
        });

        document.getElementById('btn-insert-col')?.addEventListener('click', () => {
            this.commitActiveEdit();
            if (!this.activeCell) return;
            const col = this.activeCell.match(/[A-Z]+/)?.[0]!;
            const colIdx = COL_NAMES.indexOf(col);
            this.store.insertCol(this.store.getActiveSheetName(), colIdx);
            this.initGrid();
        });

        document.getElementById('btn-insert-50-rows')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.store.insertRow(this.store.getActiveSheetName(), DEFAULT_ROWS, 50);
            this.initGrid();
        });

        document.getElementById('btn-insert-50-cols')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.store.insertCol(this.store.getActiveSheetName(), DEFAULT_COLS, 50);
            this.initGrid();
        });

        document.getElementById('btn-find-replace')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.showFindReplaceDialog();
        });
        document.getElementById('btn-new-sheet')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.store.saveHistory();
            const newName = this.store.addSheet();
            this.store.setActiveSheet(newName);
            this.initTabs();
            this.initGrid();
        });
        document.getElementById('btn-rename-sheet')?.addEventListener('click', () => {
            this.commitActiveEdit();
            const currentName = this.store.getActiveSheetName();
            const newName = prompt('Rename sheet to:', currentName);
            if (newName && newName !== currentName) {
                if (this.store.renameSheet(currentName, newName)) {
                    this.initTabs();
                    this.initGrid();
                }
            }
        });
        document.getElementById('btn-link-cell')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.showLinkCellDialog();
        });
        document.getElementById('btn-create-table')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.showCreateTableDialog();
        });
        document.getElementById('btn-insert-chart')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.insertChart();
        });
        document.getElementById('btn-insert-sparkline')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.insertSparkline();
        });
        document.getElementById('btn-conditional')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.showConditionalFormattingDialog();
        });
        document.getElementById('btn-undo')?.addEventListener('click', () => this.handleUndo());
        document.getElementById('btn-redo')?.addEventListener('click', () => this.handleRedo());
        document.getElementById('btn-shortcuts')?.addEventListener('click', () => this.showShortcutsDialog());
        document.getElementById('btn-formula-help')?.addEventListener('click', () => this.showFormulaHelpDialog());

        // Borders
        const borderBtn = document.getElementById('btn-borders');
        const borderDropdown = document.getElementById('dropdown-borders');
        borderBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            borderDropdown?.classList.toggle('active');
        });

        borderDropdown?.querySelectorAll('.dropdown-content div[data-border]').forEach(item => {
            item.addEventListener('click', (e) => {
                const type = (item as HTMLElement).dataset.border as any;
                this.toggleBorder(type);
                borderDropdown.classList.remove('active');
            });
        });

        window.addEventListener('click', () => {
            borderDropdown?.classList.remove('active');
            document.getElementById('dropdown-format')?.classList.remove('active');
        });

        // Format
        const formatBtn = document.getElementById('btn-format');
        const formatDropdown = document.getElementById('dropdown-format');
        formatBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            formatDropdown?.classList.toggle('active');
        });

        document.getElementById('option-row-height')?.addEventListener('click', () => {
            this.showRowHeightDialog();
            formatDropdown?.classList.remove('active');
        });

        document.getElementById('option-col-width')?.addEventListener('click', () => {
            this.showColWidthDialog();
            formatDropdown?.classList.remove('active');
        });
        
        // Add Sheet
        this.tabBar.querySelector('.add-sheet-btn')?.addEventListener('click', () => {
            this.commitActiveEdit();
            this.store.saveHistory();
            const newName = this.store.addSheet();
            this.store.setActiveSheet(newName);
            this.initTabs();
            this.initGrid();
        });

        // Themes
        document.getElementById('select-theme')?.addEventListener('change', (e) => {
            const theme = (e.target as HTMLSelectElement).value;
            document.body.setAttribute('data-theme', theme);
            const sheetName = this.store.getActiveSheetName();
            const metadata = this.store.getMetadata(sheetName);
            this.store.setMetadata(sheetName, { ...metadata, theme: theme as any });
        });

        document.getElementById('btn-print')?.addEventListener('click', () => {
            window.print();
        });

        document.getElementById('btn-new-comment')?.addEventListener('click', () => this.showCommentDialog());
        document.getElementById('btn-delete-comment')?.addEventListener('click', () => this.deleteComment());
        document.getElementById('btn-pivot-table')?.addEventListener('click', () => this.showPivotTableDialog());
        document.getElementById('btn-define-name')?.addEventListener('click', () => this.showDefineNameDialog());
        document.getElementById('btn-print-pdf')?.addEventListener('click', () => window.print());

        // Resizing
        let isResizing = false;
        let startX: number, startY: number, startSize: number, target: HTMLElement | null = null;

        this.grid.addEventListener('mousedown', (e) => {
            const resizer = (e.target as HTMLElement).closest('.resizer') as HTMLElement;
            if (!resizer) return;

            isResizing = true;
            target = resizer;
            startX = e.pageX;
            startY = e.pageY;

            if (resizer.classList.contains('col-resizer')) {
                const col = resizer.dataset.col!;
                const header = resizer.parentElement!;
                startSize = header.offsetWidth;
            } else {
                const row = resizer.dataset.row!;
                const header = resizer.parentElement!;
                startSize = header.offsetHeight;
            }

            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        });

        const onMouseMove = (e: MouseEvent) => {
            if (!isResizing || !target) return;

            if (target.classList.contains('col-resizer')) {
                const diff = e.pageX - startX;
                const newWidth = Math.max(40, startSize + diff);
                const col = target.dataset.col!;
                this.grid.style.setProperty(`--col-width-${col}`, `${newWidth}px`);
            } else {
                const diff = e.pageY - startY;
                const newHeight = Math.max(20, startSize + diff);
                const row = target.dataset.row!;
                this.grid.style.setProperty(`--row-height-${row}`, `${newHeight}px`);
            }
            this.updateSelectionIndicator();
        };

        const onMouseUp = () => {
            isResizing = false;
            document.removeEventListener('mousemove', onMouseMove);
            document.removeEventListener('mouseup', onMouseUp);
        };
    }

    private updateSelectionIndicator() {
        if (!this.activeCell || !this.selectionIndicator || this.selectedCells.size === 0) {
            if (this.selectionIndicator) this.selectionIndicator.style.display = 'none';
            return;
        }
        
        const activeCellEl = this.grid.querySelector(`[data-id="${this.activeCell}"]`) as HTMLElement;
        if (!activeCellEl) return;

        const containerRect = this.gridContainer.getBoundingClientRect();

        if (this.lastSelectionWasRange || this.isSelecting) {
            let minLeft = Infinity, minTop = Infinity, maxRight = -Infinity, maxBottom = -Infinity;
            let hasVisibleCells = false;

            this.selectedCells.forEach(id => {
                const cellEl = this.grid.querySelector(`[data-id="${id}"]`) as HTMLElement;
                if (cellEl) {
                    hasVisibleCells = true;
                    const rect = cellEl.getBoundingClientRect();
                    minLeft = Math.min(minLeft, rect.left);
                    minTop = Math.min(minTop, rect.top);
                    maxRight = Math.max(maxRight, rect.right);
                    maxBottom = Math.max(maxBottom, rect.bottom);
                }
            });

            if (hasVisibleCells) {
                this.selectionIndicator.style.display = 'block';
                this.selectionIndicator.style.left = (minLeft - containerRect.left + this.gridContainer.scrollLeft) + 'px';
                this.selectionIndicator.style.top = (minTop - containerRect.top + this.gridContainer.scrollTop) + 'px';
                this.selectionIndicator.style.width = (maxRight - minLeft) + 'px';
                this.selectionIndicator.style.height = (maxBottom - minTop) + 'px';
            }
        } else {
            const rect = activeCellEl.getBoundingClientRect();
            this.selectionIndicator.style.display = 'block';
            this.selectionIndicator.style.left = (rect.left - containerRect.left + this.gridContainer.scrollLeft) + 'px';
            this.selectionIndicator.style.top = (rect.top - containerRect.top + this.gridContainer.scrollTop) + 'px';
            this.selectionIndicator.style.width = rect.width + 'px';
            this.selectionIndicator.style.height = rect.height + 'px';
        }
    }

    private commitActiveEdit() {
        if (this.isEditing) {
            const editingCell = this.grid.querySelector('.grid-cell.editing') as HTMLElement;
            if (editingCell) {
                editingCell.blur();
            }
        }
    }

    private selectColumn(col: string, multi: boolean = false, range: boolean = false) {
        if (this.isEditing) this.commitActiveEdit();

        const startId = `${col}1`;
        const endId = `${col}${DEFAULT_ROWS}`;

        this.lastSelectionWasRange = true; // Full column is always treated as range for indicator
        
        if (!multi && !range) {
            this.selectedCells.clear();
            this.selectionAnchor = startId;
        }

        if (range && this.selectionAnchor) {
            this.selectedCells.clear();
            const anchorCol = this.selectionAnchor.match(/[A-Z]+/)?.[0]!;
            const rangeCells = this.getRange(`${anchorCol}1`, `${col}${DEFAULT_ROWS}`);
            rangeCells.forEach(id => this.selectedCells.add(id));
        } else {
            const colCells = this.getRange(startId, endId);
            colCells.forEach(id => {
                if (multi) {
                    if (this.selectedCells.has(id)) this.selectedCells.delete(id);
                    else this.selectedCells.add(id);
                } else {
                    this.selectedCells.add(id);
                }
            });
        }

        this.activeCell = startId;
        this.updateToolbarState();
        this.updateGridDisplay();
        this.updateSelectionIndicator();
    }

    private selectRow(row: string, multi: boolean = false, range: boolean = false) {
        if (this.isEditing) this.commitActiveEdit();

        const startId = `A${row}`;
        const endId = `${COL_NAMES[DEFAULT_COLS - 1]}${row}`;

        this.lastSelectionWasRange = true; // Full row is always treated as range for indicator
        
        if (!multi && !range) {
            this.selectedCells.clear();
            this.selectionAnchor = startId;
        }

        if (range && this.selectionAnchor) {
            this.selectedCells.clear();
            const anchorRow = this.selectionAnchor.match(/[0-9]+/)?.[0]!;
            const rangeCells = this.getRange(`A${anchorRow}`, `${COL_NAMES[DEFAULT_COLS - 1]}${row}`);
            rangeCells.forEach(id => this.selectedCells.add(id));
        } else {
            const rowCells = this.getRange(startId, endId);
            rowCells.forEach(id => {
                if (multi) {
                    if (this.selectedCells.has(id)) this.selectedCells.delete(id);
                    else this.selectedCells.add(id);
                } else {
                    this.selectedCells.add(id);
                }
            });
        }

        this.activeCell = startId;
        this.updateToolbarState();
        this.updateGridDisplay();
        this.updateSelectionIndicator();
    }

    private selectCell(id: string, multi: boolean = false, range: boolean = false) {
        if (this.isEditing && this.activeCell !== id) {
            this.commitActiveEdit();
        }

        this.lastSelectionWasRange = range;
        if (!multi && !range) {
            this.selectedCells.clear();
            this.selectionAnchor = id;
        }

        if (range && this.selectionAnchor) {
            this.selectedCells.clear();
            const rangeCells = this.getRange(this.selectionAnchor, id);
            rangeCells.forEach(cell => this.selectedCells.add(cell));
        } else {
            if (multi) {
                if (this.selectedCells.has(id)) {
                    this.selectedCells.delete(id);
                } else {
                    this.selectedCells.add(id);
                }
            } else {
                this.selectedCells.add(id);
            }
        }

        this.activeCell = id;
        const cellEl = this.grid.querySelector(`[data-id="${id}"]`) as HTMLElement;
        if (cellEl && !range) {
            cellEl.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }

        // Add pulse animation
        if (this.selectionIndicator && !range) {
            this.selectionIndicator.classList.remove('pulse');
            void this.selectionIndicator.offsetWidth; // Trigger reflow
            this.selectionIndicator.classList.add('pulse');
        }

        const cellData = this.store.getCell(this.store.getActiveSheetName(), id);
        this.cellAddress.textContent = id;
        this.formulaBar.value = cellData.formula || cellData.rawValue;
        
        this.updateToolbarState();
        this.updateGridDisplay();
    }

    private updateToolbarState() {
        if (!this.activeCell) return;
        const cellData = this.store.getCell(this.store.getActiveSheetName(), this.activeCell);
        
        document.getElementById('btn-bold')?.classList.toggle('active', cellData.format.bold);
        document.getElementById('btn-italic')?.classList.toggle('active', cellData.format.italic);
        document.getElementById('btn-underline')?.classList.toggle('active', cellData.format.underline);
        
        document.getElementById('btn-align-left')?.classList.toggle('active', cellData.format.textAlign === 'left');
        document.getElementById('btn-align-center')?.classList.toggle('active', cellData.format.textAlign === 'center');
        document.getElementById('btn-align-right')?.classList.toggle('active', cellData.format.textAlign === 'right');

        document.getElementById('btn-align-top')?.classList.toggle('active', cellData.format.verticalAlign === 'top');
        document.getElementById('btn-align-middle')?.classList.toggle('active', cellData.format.verticalAlign === 'middle');
        document.getElementById('btn-align-bottom')?.classList.toggle('active', cellData.format.verticalAlign === 'bottom');

        const formatSelect = document.getElementById('select-format') as HTMLSelectElement;
        if (formatSelect) formatSelect.value = cellData.format.type;

        // Update color indicators
        const fillBtn = document.getElementById('btn-fill-color');
        if (fillBtn) {
            const span = fillBtn.querySelector('span');
            if (span) span.style.borderBottomColor = cellData.format.backgroundColor;
        }
        const textBtn = document.getElementById('btn-text-color');
        if (textBtn) {
            const span = textBtn.querySelector('span');
            if (span) span.style.borderBottomColor = cellData.format.color;
        }
    }

    private getRange(startId: string, endId: string): string[] {
        const startCol = startId.match(/[A-Z]+/)?.[0]!;
        const startRow = parseInt(startId.match(/[0-9]+/)?.[0]!);
        const endCol = endId.match(/[A-Z]+/)?.[0]!;
        const endRow = parseInt(endId.match(/[0-9]+/)?.[0]!);

        const startColIdx = COL_NAMES.indexOf(startCol);
        const endColIdx = COL_NAMES.indexOf(endCol);

        const minCol = Math.min(startColIdx, endColIdx);
        const maxCol = Math.max(startColIdx, endColIdx);
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);

        const range: string[] = [];
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                range.push(`${COL_NAMES[c]}${r}`);
            }
        }
        return range;
    }

    private startEditing(cellEl: HTMLElement, event?: MouseEvent) {
        if (this.isEditing && cellEl.classList.contains('editing')) return;

        const id = cellEl.dataset.id!;
        const cellData = this.store.getCell(this.store.getActiveSheetName(), id);
        
        this.isEditing = true;
        cellEl.contentEditable = 'true';
        cellEl.classList.add('editing');
        cellEl.textContent = cellData.formula || cellData.rawValue;
        cellEl.focus();

        // Position caret
        const sel = window.getSelection();
        if (sel) {
            let range: Range | null = null;
            if (event) {
                if ((document as any).caretRangeFromPoint) {
                    range = (document as any).caretRangeFromPoint(event.clientX, event.clientY);
                } else if ((document as any).caretPositionFromPoint) {
                    const pos = (document as any).caretPositionFromPoint(event.clientX, event.clientY);
                    if (pos) {
                        range = document.createRange();
                        range.setStart(pos.offsetNode, pos.offset);
                        range.collapse(true);
                    }
                }
            }

            if (!range) {
                range = document.createRange();
                range.selectNodeContents(cellEl);
                range.collapse(false);
            }

            sel.removeAllRanges();
            sel.addRange(range);
        }

        const onBlur = () => {
            this.commitValue(id, cellEl.textContent || '');
            this.isEditing = false;
            cellEl.contentEditable = 'false';
            cellEl.classList.remove('editing');
            if (this.suggestionBox) this.suggestionBox.style.display = 'none';
            if (this.previewBox) this.previewBox.style.display = 'none';
            this.suggestionIndex = -1;
            cellEl.removeEventListener('blur', onBlur);
            cellEl.removeEventListener('keydown', onKeydown);
            cellEl.removeEventListener('input', onInput);
        };

        const onInput = () => {
            this.updateSuggestions(cellEl, cellEl.textContent || '');
        };

        const onKeydown = (e: KeyboardEvent) => {
            if (this.suggestionBox?.style.display === 'block') {
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    e.stopPropagation();
                    const matches = this.getMatches(cellEl.textContent || '');
                    this.suggestionIndex = (this.suggestionIndex + 1) % matches.length;
                    this.updateSuggestions(cellEl, cellEl.textContent || '');
                    return;
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    e.stopPropagation();
                    const matches = this.getMatches(cellEl.textContent || '');
                    this.suggestionIndex = (this.suggestionIndex - 1 + matches.length) % matches.length;
                    this.updateSuggestions(cellEl, cellEl.textContent || '');
                    return;
                } else if (e.key === 'Tab' || e.key === 'Enter') {
                    if (this.suggestionIndex >= 0) {
                        e.preventDefault();
                        e.stopPropagation();
                        const matches = this.getMatches(cellEl.textContent || '');
                        this.applySuggestion(cellEl, matches[this.suggestionIndex]);
                        return;
                    }
                }
            }

            if (e.key === 'Enter') {
                e.preventDefault();
                e.stopPropagation();
                cellEl.blur();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                e.stopPropagation();
                this.updateGridDisplay(); // Revert
                cellEl.blur();
            } else if (e.key === 'Delete' || e.key === 'Backspace') {
                // Let native behavior happen inside contentEditable
                e.stopPropagation();
            }
        };

        cellEl.addEventListener('blur', onBlur);
        cellEl.addEventListener('keydown', onKeydown);
        cellEl.addEventListener('input', onInput);
    }

    private getMatches(text: string): string[] {
        const lastPart = text.split(/[\(\),+\-*/^]/).pop()?.toUpperCase() || '';
        if (lastPart.length === 0) return [];
        return FUNCTIONS.filter(f => f.startsWith(lastPart));
    }

    private updateSuggestions(el: HTMLElement | HTMLInputElement, text: string) {
        if (!this.suggestionBox || !this.previewBox) return;

        if (!text.startsWith('=')) {
            this.suggestionBox.style.display = 'none';
            this.previewBox.style.display = 'none';
            return;
        }

        // Real-time Preview
        try {
            const parser = new FormulaParser(this.store, this.store.getActiveSheetName());
            const result = parser.parse(text);
            if (result !== '#ERROR!' && result !== '#NAME?') {
                this.previewBox.style.display = 'block';
                this.previewBox.textContent = `Preview: ${result}`;
                const rect = el.getBoundingClientRect();
                this.previewBox.style.left = `${rect.left}px`;
                this.previewBox.style.top = `${rect.bottom + 5}px`;
            } else {
                this.previewBox.style.display = 'none';
            }
        } catch (e) {
            this.previewBox.style.display = 'none';
        }

        // Suggestions
        const matches = this.getMatches(text);
        if (matches.length > 0) {
            this.suggestionBox.style.display = 'block';
            this.suggestionBox.innerHTML = matches.map((m, i) => {
                const details = FUNCTION_DETAILS[m];
                return `
                    <div class="suggestion-item ${i === this.suggestionIndex ? 'active' : ''}">
                        <div class="suggestion-name">${m}</div>
                        <div class="suggestion-sig">${details.sig}</div>
                        <div class="suggestion-desc">${details.desc}</div>
                    </div>
                `;
            }).join('');
            
            const rect = el.getBoundingClientRect();
            this.suggestionBox.style.left = `${rect.left}px`;
            this.suggestionBox.style.top = `${rect.bottom + (this.previewBox.style.display === 'block' ? 30 : 5)}px`;

            this.suggestionBox.querySelectorAll('.suggestion-item').forEach((item, i) => {
                item.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.applySuggestion(el, matches[i]);
                });
            });
        } else {
            this.suggestionBox.style.display = 'none';
            this.suggestionIndex = -1;
        }
    }

    private applySuggestion(el: HTMLElement | HTMLInputElement, suggestion: string) {
        const isInput = el instanceof HTMLInputElement;
        const text = isInput ? el.value : el.textContent || '';
        const lastSymbolIndex = Math.max(
            text.lastIndexOf('('),
            text.lastIndexOf(','),
            text.lastIndexOf('+'),
            text.lastIndexOf('-'),
            text.lastIndexOf('*'),
            text.lastIndexOf('/'),
            text.lastIndexOf('^'),
            text.lastIndexOf('=')
        );
        
        const prefix = text.substring(0, lastSymbolIndex + 1);
        const newText = prefix + suggestion + '(';
        
        if (isInput) {
            el.value = newText;
            el.focus();
        } else {
            el.textContent = newText;
            // Move cursor to end
            const range = document.createRange();
            const sel = window.getSelection();
            if (el.lastChild) {
                range.setStart(el.lastChild, el.lastChild.textContent?.length || 0);
                range.collapse(true);
                sel?.removeAllRanges();
                sel?.addRange(range);
            }
        }

        this.suggestionBox!.style.display = 'none';
        this.suggestionIndex = -1;
        this.updateSuggestions(el, newText);
    }

    private commitValue(id: string, input: string, sheetName: string = this.store.getActiveSheetName(), skipHistory: boolean = false) {
        const cell = this.store.getCell(sheetName, id);
        
        // Check if value actually changed
        if (cell.formula === input || (cell.formula === '' && cell.rawValue === input)) return;

        if (!skipHistory) this.store.saveHistory();

        // Remove old precedents
        cell.precedents.forEach(p => {
            let pSheet = sheetName;
            let pId = p;
            if (p.includes('!')) {
                const parts = p.split('!');
                pSheet = parts[0];
                pId = parts[1];
                if (pSheet.startsWith("'") && pSheet.endsWith("'")) {
                    pSheet = pSheet.slice(1, -1);
                }
            }
            const pCell = this.store.getCell(pSheet, pId);
            pCell.dependents.delete(`${sheetName}!${id}`);
        });
        cell.precedents.clear();

        if (input.startsWith('=')) {
            cell.formula = input;
            cell.rawValue = '';
            
            const parser = new FormulaParser(this.store, sheetName);
            const refs = parser.extractReferences(input);
            refs.forEach(ref => {
                let pSheet = sheetName;
                let pId = ref;
                if (ref.includes('!')) {
                    const parts = ref.split('!');
                    pSheet = parts[0];
                    pId = parts[1];
                    if (pSheet.startsWith("'") && pSheet.endsWith("'")) {
                        pSheet = pSheet.slice(1, -1);
                    }
                }
                const fullRef = `${pSheet}!${pId}`;
                cell.precedents.add(fullRef);
                const pCell = this.store.getCell(pSheet, pId);
                pCell.dependents.add(`${sheetName}!${id}`);
            });
        } else {
            cell.formula = '';
            cell.rawValue = input;
        }

        this.recalculate(id, sheetName);
        if (sheetName === this.store.getActiveSheetName()) {
            this.updateGridDisplay();
            if (this.activeCell === id) {
                this.formulaBar.value = input;
            }
        }
    }

    private recalculate(startId: string, startSheet: string = this.store.getActiveSheetName()) {
        const queue: { sheet: string, id: string }[] = [{ sheet: startSheet, id: startId }];
        const visited = new Set<string>();

        while (queue.length > 0) {
            const { sheet: sName, id } = queue.shift()!;
            const key = `${sName}!${id}`;
            if (visited.has(key)) {
                this.store.updateCell(sName, id, { computedValue: '#CYCLE!' });
                continue;
            }
            visited.add(key);

            const parser = new FormulaParser(this.store, sName);
            const cell = this.store.getCell(sName, id);
            if (cell.formula) {
                cell.computedValue = parser.parse(cell.formula);
            } else {
                cell.computedValue = cell.rawValue;
            }

            cell.dependents.forEach(dep => {
                let depSheet = sName;
                let depId = dep;
                if (dep.includes('!')) {
                    const parts = dep.split('!');
                    depSheet = parts[0];
                    depId = parts[1];
                    if (depSheet.startsWith("'") && depSheet.endsWith("'")) {
                        depSheet = depSheet.slice(1, -1);
                    }
                }
                queue.push({ sheet: depSheet, id: depId });
            });
        }
    }

    private renderSparkline(data: number[]): string {
        if (data.length < 2) return '';
        
        const width = 100;
        const height = 20;
        const min = Math.min(...data);
        const max = Math.max(...data);
        const range = max - min || 1;
        
        const points = data.map((val, i) => {
            const x = (i / (data.length - 1)) * width;
            const y = height - ((val - min) / range) * height;
            return `${x},${y}`;
        }).join(' ');
        
        return `
            <svg width="100%" height="100%" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" style="display: block;">
                <polyline points="${points}" fill="none" stroke="#1a73e8" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" />
            </svg>
        `;
    }

    private updateHeaderHighlights() {
        const activeCols = new Set<string>();
        const activeRows = new Set<string>();
        
        this.selectedCells.forEach(id => {
            const col = id.match(/[A-Z]+/)?.[0];
            const row = id.match(/[0-9]+/)?.[0];
            if (col) activeCols.add(col);
            if (row) activeRows.add(row);
        });
        
        this.grid.querySelectorAll('.col-header').forEach(h => {
            const col = (h as HTMLElement).dataset.col;
            if (col) h.classList.toggle('active-header', activeCols.has(col));
        });
        
        this.grid.querySelectorAll('.row-header').forEach(h => {
            const row = (h as HTMLElement).dataset.row;
            if (row) h.classList.toggle('active-header', activeRows.has(row));
        });

        // Corner header highlight if everything is selected
        const isAllSelected = this.selectedCells.size === DEFAULT_ROWS * DEFAULT_COLS;
        this.grid.querySelector('.corner-header')?.classList.toggle('active-header', isAllSelected);
    }

    private updateGridDisplay() {
        const sheetName = this.store.getActiveSheetName();
        const cells = this.grid.querySelectorAll('.grid-cell');
        cells.forEach(cellEl => {
            const id = (cellEl as HTMLElement).dataset.id!;
            const data = this.store.getCell(sheetName, id);
            
            let displayValue = data.computedValue?.toString() || '';
            
            // Apply number formatting
            if (data.format.type === 'currency' && !isNaN(Number(data.computedValue))) {
                displayValue = '$' + Number(data.computedValue).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            } else if (data.format.type === 'number' && !isNaN(Number(data.computedValue))) {
                displayValue = Number(data.computedValue).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            } else if (data.format.type === 'percentage' && !isNaN(Number(data.computedValue))) {
                displayValue = (Number(data.computedValue) * 100).toFixed(2) + '%';
            } else if (data.format.type === 'date' && !isNaN(Date.parse(data.computedValue))) {
                displayValue = new Date(data.computedValue).toLocaleDateString();
            }

            if (data.computedValue && typeof data.computedValue === 'object' && data.computedValue.type === 'sparkline') {
                (cellEl as HTMLElement).classList.add('sparkline-cell');
                (cellEl as HTMLElement).innerHTML = this.renderSparkline(data.computedValue.data);
            } else {
                (cellEl as HTMLElement).classList.remove('sparkline-cell');
                (cellEl as HTMLElement).textContent = displayValue;
            }
            
            // Base formatting
            let finalFormat = { ...data.format };
            (cellEl as HTMLElement).style.whiteSpace = finalFormat.wrapText ? 'normal' : 'nowrap';
            (cellEl as HTMLElement).style.overflowWrap = finalFormat.wrapText ? 'break-word' : 'normal';
            (cellEl as HTMLElement).style.wordBreak = finalFormat.wrapText ? 'break-all' : 'normal';
            (cellEl as HTMLElement).style.alignItems = finalFormat.wrapText ? 'flex-start' : 'center';
            (cellEl as HTMLElement).style.padding = finalFormat.wrapText ? '4px' : '0 4px';

            // Table Formatting
            const metadata = this.store.getMetadata(sheetName);
            const table = metadata.tables.find(t => this.isCellInTable(id, t));
            if (table) {
                const { r, c } = this.parseCellId(id);
                const { startR, startC, endR, endC } = this.parseRange(table.range);
                
                (cellEl as HTMLElement).classList.add('table-cell', `table-${table.style}`);
                
                if (table.hasHeader && r === startR) {
                    (cellEl as HTMLElement).classList.add('table-header');
                    if (!(cellEl as any)._tableListener) {
                        (cellEl as any)._tableListener = (e: MouseEvent) => {
                            if (e.offsetX > (cellEl as HTMLElement).offsetWidth - 24) {
                                this.showFilterMenu(table, c - startC, e);
                            }
                        };
                        (cellEl as HTMLElement).addEventListener('click', (cellEl as any)._tableListener);
                    }
                } else {
                    const relativeRow = r - (table.hasHeader ? startR + 1 : startR);
                    (cellEl as HTMLElement).classList.add(relativeRow % 2 === 0 ? 'table-row-even' : 'table-row-odd');
                }

                // Check if row should be filtered out
                let isFiltered = false;
                Object.entries(table.filters).forEach(([colIdx, allowedValues]) => {
                    if (allowedValues.length > 0) {
                        const cellData = this.store.getCell(sheetName, `${COL_NAMES[startC + Number(colIdx)]}${r}`);
                        if (!allowedValues.includes(String(cellData.computedValue))) {
                            isFiltered = true;
                        }
                    }
                });

                if (isFiltered && r > (table.hasHeader ? startR : startR - 1)) {
                    (cellEl as HTMLElement).style.display = 'none';
                    this.grid.style.setProperty(`--row-height-${r}`, '0px');
                    const rowHeader = this.grid.querySelector(`.row-header[style*="grid-row: ${r + 1}"]`) as HTMLElement;
                    if (rowHeader) rowHeader.style.display = 'none';
                } else {
                    (cellEl as HTMLElement).style.display = '';
                    // Restore original height if it was set, otherwise default to auto/25px
                    // For simplicity, we'll try to get it from sheet metadata or use a default
                    this.grid.style.removeProperty(`--row-height-${r}`);
                    const rowHeader = this.grid.querySelector(`.row-header[style*="grid-row: ${r + 1}"]`) as HTMLElement;
                    if (rowHeader) rowHeader.style.display = '';
                }
            } else {
                (cellEl as HTMLElement).classList.remove('table-cell', 'table-blue', 'table-green', 'table-orange', 'table-header', 'table-row-even', 'table-row-odd');
                (cellEl as HTMLElement).style.display = '';
                const { r } = this.parseCellId(id);
                const rowHeader = this.grid.querySelector(`.row-header[style*="grid-row: ${r + 1}"]`) as HTMLElement;
                if (rowHeader) rowHeader.style.display = '';
            }

            // Evaluate Conditional Formatting
            data.conditionalRules.forEach(rule => {
                const val = data.computedValue;
                let match = false;
                const ruleVal = isNaN(Number(rule.value)) ? rule.value : Number(rule.value);
                const cellVal = isNaN(Number(val)) ? val : Number(val);

                switch (rule.operator) {
                    case 'gt': match = Number(cellVal) > Number(ruleVal); break;
                    case 'lt': match = Number(cellVal) < Number(ruleVal); break;
                    case 'eq': match = cellVal == ruleVal; break;
                    case 'contains': match = String(cellVal).includes(String(ruleVal)); break;
                }

                if (match) {
                    finalFormat = { ...finalFormat, ...rule.format };
                }
            });

            // Apply final formatting
            (cellEl as HTMLElement).style.fontWeight = finalFormat.bold ? 'bold' : 'normal';
            (cellEl as HTMLElement).style.fontStyle = finalFormat.italic ? 'italic' : 'normal';
            (cellEl as HTMLElement).style.textDecoration = finalFormat.underline ? 'underline' : 'none';
            
            // Apply borders
            (cellEl as HTMLElement).style.borderTop = finalFormat.borderTop || '';
            (cellEl as HTMLElement).style.borderBottom = finalFormat.borderBottom || '';
            (cellEl as HTMLElement).style.borderLeft = finalFormat.borderLeft || '';
            (cellEl as HTMLElement).style.borderRight = finalFormat.borderRight || '';
            
            const isSelected = this.selectedCells.has(id);
            const isActive = this.activeCell === id;
            
            let bgColor = finalFormat.backgroundColor;
            if (isSelected) {
                if (isActive) {
                    bgColor = finalFormat.backgroundColor;
                } else {
                    // Layer the semi-transparent selection color over the cell background
                    bgColor = `linear-gradient(var(--active-cell-bg), var(--active-cell-bg)), ${finalFormat.backgroundColor}`;
                }
            }
            
            (cellEl as HTMLElement).style.background = bgColor;
            (cellEl as HTMLElement).classList.toggle('selected', isSelected);
            (cellEl as HTMLElement).classList.toggle('active', isActive);
            (cellEl as HTMLElement).style.color = finalFormat.color;
            (cellEl as HTMLElement).style.textAlign = finalFormat.textAlign;

            // Comments
            (cellEl as HTMLElement).classList.toggle('has-comment', !!data.comment);

            // Data Validation
            let isValid = true;
            if (data.validation) {
                const val = data.computedValue;
                if (data.validation.type === 'number') {
                    const n = Number(val);
                    if (isNaN(n) || (data.validation.min !== undefined && n < data.validation.min) || (data.validation.max !== undefined && n > data.validation.max)) {
                        isValid = false;
                    }
                } else if (data.validation.type === 'list') {
                    if (data.validation.criteria && !data.validation.criteria.includes(String(val))) {
                        isValid = false;
                    }
                }
                
                // Add dropdown indicator if list and is active
                if (data.validation.type === 'list' && !(cellEl as any)._dropdownIndicator) {
                    const indicator = document.createElement('div');
                    indicator.className = 'dropdown-indicator';
                    indicator.innerHTML = '▼';
                    (cellEl as HTMLElement).appendChild(indicator);
                    (cellEl as any)._dropdownIndicator = indicator;
                } else if (data.validation.type !== 'list' && (cellEl as any)._dropdownIndicator) {
                    (cellEl as any)._dropdownIndicator.remove();
                    delete (cellEl as any)._dropdownIndicator;
                }
                
                // Handle Checkbox
                if (data.validation.type === 'checkbox') {
                    (cellEl as HTMLElement).innerHTML = `<input type="checkbox" ${val === true ? 'checked' : ''} style="pointer-events: none;">`;
                }
            }
            (cellEl as HTMLElement).classList.toggle('invalid-value', !isValid);
            
            // Align items in flex container
            switch (finalFormat.verticalAlign) {
                case 'top': (cellEl as HTMLElement).style.alignItems = 'flex-start'; break;
                case 'middle': (cellEl as HTMLElement).style.alignItems = 'center'; break;
                case 'bottom': (cellEl as HTMLElement).style.alignItems = 'flex-end'; break;
                default: (cellEl as HTMLElement).style.alignItems = 'center';
            }

            switch (finalFormat.textAlign) {
                case 'left': (cellEl as HTMLElement).style.justifyContent = 'flex-start'; break;
                case 'center': (cellEl as HTMLElement).style.justifyContent = 'center'; break;
                case 'right': (cellEl as HTMLElement).style.justifyContent = 'flex-end'; break;
                default: (cellEl as HTMLElement).style.justifyContent = 'flex-start';
            }
        });
        this.updateSelectionIndicator();
        this.updateHeaderHighlights();
    }

    private hoverCommentBox: HTMLElement | null = null;

    private showHoverComment(comment: string, e: MouseEvent) {
        if (!this.hoverCommentBox) {
            this.hoverCommentBox = document.createElement('div');
            this.hoverCommentBox.className = 'cell-comment-box';
            document.body.appendChild(this.hoverCommentBox);
        }
        this.hoverCommentBox.textContent = comment;
        this.hoverCommentBox.style.display = 'block';
        this.hoverCommentBox.style.left = `${e.pageX + 10}px`;
        this.hoverCommentBox.style.top = `${e.pageY + 10}px`;
    }

    private hideHoverComment() {
        if (this.hoverCommentBox) {
            this.hoverCommentBox.style.display = 'none';
        }
    }

    private showDropdownMenu(cell: CellData, cellEl: HTMLElement) {
        const id = cellEl.dataset.id!;
        const rect = cellEl.getBoundingClientRect();
        const menu = document.createElement('div');
        menu.className = 'dropdown-content active';
        menu.style.position = 'fixed';
        menu.style.top = `${rect.bottom}px`;
        menu.style.left = `${rect.left}px`;
        menu.style.minWidth = `${rect.width}px`;
        menu.style.zIndex = '3000';
        menu.style.display = 'block';
        menu.style.maxHeight = '200px';
        menu.style.overflowY = 'auto';
        menu.style.background = 'white';
        menu.style.border = '1px solid #ccc';
        menu.style.boxShadow = '0 2px 8px rgba(0,0,0,0.15)';

        (cell.validation?.criteria || []).forEach(val => {
            const item = document.createElement('div');
            item.className = 'dropdown-item';
            item.textContent = val;
            item.style.padding = '4px 8px';
            item.style.cursor = 'pointer';
            item.style.fontSize = '12px';
            item.onmouseenter = () => item.style.background = '#f0f0f0';
            item.onmouseleave = () => item.style.background = 'transparent';
            item.onclick = (e) => {
                e.stopPropagation();
                this.store.saveHistory();
                cell.rawValue = val;
                cell.computedValue = val;
                this.updateGridDisplay();
                this.recalculate(id);
                if (menu.parentNode) menu.remove();
            };
            menu.appendChild(item);
        });

        document.body.appendChild(menu);

        const close = (e: MouseEvent) => {
            if (!menu.contains(e.target as Node)) {
                if (menu.parentNode) menu.remove();
                window.removeEventListener('click', close);
            }
        };
        setTimeout(() => window.addEventListener('click', close), 10);
    }

    private getSelectionRange() {
        const bounds = this.getSelectionBounds();
        if (!bounds) return '';
        const { startR, startC, endR, endC } = bounds;
        if (startR === endR && startC === endC) return `${COL_NAMES[startC]}${startR}`;
        return `${COL_NAMES[startC]}${startR}:${COL_NAMES[endC]}${endR}`;
    }

    private showRowHeightDialog() {
        if (this.selectedCells.size === 0) {
            alert('Please select at least one cell.');
            return;
        }

        let currentHeight = 25;
        if (this.activeCell) {
            const { r } = this.parseCellId(this.activeCell);
            currentHeight = parseInt(getComputedStyle(this.grid).getPropertyValue(`--row-height-${r + 1}`)) || 25;
        }

        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100vw';
        overlay.style.height = '100vh';
        overlay.style.background = 'rgba(0,0,0,0.5)';
        overlay.style.zIndex = '2000';

        const dialog = document.createElement('div');
        dialog.className = 'modal-content';
        dialog.style.position = 'fixed';
        dialog.style.top = '50%';
        dialog.style.left = '50%';
        dialog.style.transform = 'translate(-50%, -50%)';
        dialog.style.background = 'white';
        dialog.style.padding = '20px';
        dialog.style.borderRadius = '8px';
        dialog.style.width = '250px';
        dialog.style.zIndex = '2001';

        dialog.innerHTML = `
            <h3 style="margin-top:0">Row Height</h3>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Row Height:</label>
                <input type="number" id="input-height-val" value="${currentHeight}" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            </div>
            <div style="display:flex; justify-content:flex-end; gap:10px;">
                <button class="btn btn-secondary" id="btn-cancel-height">Cancel</button>
                <button class="btn btn-primary" id="btn-apply-height">OK</button>
            </div>
        `;

        document.body.appendChild(overlay);
        document.body.appendChild(dialog);

        const close = () => {
            if (overlay.parentNode) overlay.remove();
            if (dialog.parentNode) dialog.remove();
        };

        dialog.querySelector('#btn-cancel-height')?.addEventListener('click', close);
        dialog.querySelector('#btn-apply-height')?.addEventListener('click', () => {
            const val = parseInt((dialog.querySelector('#input-height-val') as HTMLInputElement).value);
            if (!isNaN(val) && val > 0) {
                const rows = new Set<number>();
                this.selectedCells.forEach(id => {
                    rows.add(this.parseCellId(id).r + 1);
                });
                rows.forEach(r => {
                    this.grid.style.setProperty(`--row-height-${r}`, `${val}px`);
                });
                this.updateGridDisplay();
            }
            close();
        });
    }

    private showColWidthDialog() {
        if (this.selectedCells.size === 0) {
            alert('Please select at least one cell.');
            return;
        }

        let currentWidth = 100;
        if (this.activeCell) {
            const col = this.activeCell.match(/[A-Z]+/)?.[0];
            if (col) {
                currentWidth = parseInt(getComputedStyle(this.grid).getPropertyValue(`--col-width-${col}`)) || 100;
            }
        }

        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100vw';
        overlay.style.height = '100vh';
        overlay.style.background = 'rgba(0,0,0,0.5)';
        overlay.style.zIndex = '2000';

        const dialog = document.createElement('div');
        dialog.className = 'modal-content';
        dialog.style.position = 'fixed';
        dialog.style.top = '50%';
        dialog.style.left = '50%';
        dialog.style.transform = 'translate(-50%, -50%)';
        dialog.style.background = 'white';
        dialog.style.padding = '20px';
        dialog.style.borderRadius = '8px';
        dialog.style.width = '250px';
        dialog.style.zIndex = '2001';

        dialog.innerHTML = `
            <h3 style="margin-top:0">Column Width</h3>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Column Width:</label>
                <input type="number" id="input-width-val" value="${currentWidth}" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            </div>
            <div style="display:flex; justify-content:flex-end; gap:10px;">
                <button class="btn btn-secondary" id="btn-cancel-width">Cancel</button>
                <button class="btn btn-primary" id="btn-apply-width">OK</button>
            </div>
        `;

        document.body.appendChild(overlay);
        document.body.appendChild(dialog);

        const close = () => {
            if (overlay.parentNode) overlay.remove();
            if (dialog.parentNode) dialog.remove();
        };

        dialog.querySelector('#btn-cancel-width')?.addEventListener('click', close);
        dialog.querySelector('#btn-apply-width')?.addEventListener('click', () => {
            const val = parseInt((dialog.querySelector('#input-width-val') as HTMLInputElement).value);
            if (!isNaN(val) && val > 0) {
                const cols = new Set<string>();
                this.selectedCells.forEach(id => {
                    const col = id.match(/[A-Z]+/)?.[0];
                    if (col) cols.add(col);
                });
                cols.forEach(col => {
                    this.grid.style.setProperty(`--col-width-${col}`, `${val}px`);
                });
                this.updateGridDisplay();
            }
            close();
        });
    }

    private showPivotTableDialog() {
        if (this.selectedCells.size === 0) {
            alert('Please select a range of cells first.');
            return;
        }

        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100vw';
        overlay.style.height = '100vh';
        overlay.style.background = 'rgba(0,0,0,0.5)';
        overlay.style.zIndex = '2000';

        const dialog = document.createElement('div');
        dialog.className = 'modal-content';
        dialog.style.position = 'fixed';
        dialog.style.top = '50%';
        dialog.style.left = '50%';
        dialog.style.transform = 'translate(-50%, -50%)';
        dialog.style.background = 'white';
        dialog.style.padding = '20px';
        dialog.style.borderRadius = '8px';
        dialog.style.width = '350px';
        dialog.style.zIndex = '2001';

        dialog.innerHTML = `
            <h3 style="margin-top:0">Create Pivot Table</h3>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Row Field (Column):</label>
                <input type="text" id="pivot-row-col" placeholder="e.g. A" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            </div>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Column Field (Column):</label>
                <input type="text" id="pivot-col-col" placeholder="e.g. B" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            </div>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Value Field (Column):</label>
                <input type="text" id="pivot-val-col" placeholder="e.g. C" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            </div>
            <div style="display:flex; justify-content:flex-end; gap:10px;">
                <button class="btn btn-secondary" id="btn-cancel-pivot">Cancel</button>
                <button class="btn btn-primary" id="btn-apply-pivot">Create</button>
            </div>
        `;

        document.body.appendChild(overlay);
        document.body.appendChild(dialog);

        const close = () => {
            if (overlay.parentNode) overlay.remove();
            if (dialog.parentNode) dialog.remove();
        };

        dialog.querySelector('#btn-cancel-pivot')?.addEventListener('click', close);
        dialog.querySelector('#btn-apply-pivot')?.addEventListener('click', () => {
            const rowCol = (dialog.querySelector('#pivot-row-col') as HTMLInputElement).value.toUpperCase();
            const colCol = (dialog.querySelector('#pivot-col-col') as HTMLInputElement).value.toUpperCase();
            const valCol = (dialog.querySelector('#pivot-val-col') as HTMLInputElement).value.toUpperCase();

            this.createPivotTable(rowCol, colCol, valCol);
            close();
        });
    }

    private createPivotTable(rowCol: string, colCol: string, valCol: string) {
        const bounds = this.getSelectionBounds();
        if (!bounds) return;

        const data: any[] = [];
        const sheetName = this.store.getActiveSheetName();

        for (let r = bounds.startR; r <= bounds.endR; r++) {
            data.push({
                row: this.store.getCell(sheetName, `${rowCol}${r}`).computedValue,
                col: this.store.getCell(sheetName, `${colCol}${r}`).computedValue,
                val: Number(this.store.getCell(sheetName, `${valCol}${r}`).computedValue) || 0
            });
        }

        const pivotData: Record<string, Record<string, number>> = {};
        const allCols = new Set<string>();

        data.forEach(d => {
            if (!pivotData[d.row]) pivotData[d.row] = {};
            pivotData[d.row][d.col] = (pivotData[d.row][d.col] || 0) + d.val;
            allCols.add(d.col);
        });

        const sortedCols = Array.from(allCols).sort();
        const pivotRows = Object.keys(pivotData).sort();

        // Create new sheet for pivot
        const newSheet = this.store.addSheet('PivotTable');
        this.store.setActiveSheet(newSheet);

        // Header
        this.store.updateCell(newSheet, 'A1', { rawValue: `${rowCol} \\ ${colCol}`, computedValue: `${rowCol} \\ ${colCol}` });
        sortedCols.forEach((col, i) => {
            const id = `${COL_NAMES[i + 1]}1`;
            this.store.updateCell(newSheet, id, { rawValue: String(col), computedValue: col });
        });

        // Body
        pivotRows.forEach((row, rIdx) => {
            const rowId = `A${rIdx + 2}`;
            this.store.updateCell(newSheet, rowId, { rawValue: String(row), computedValue: row });
            sortedCols.forEach((col, cIdx) => {
                const id = `${COL_NAMES[cIdx + 1]}${rIdx + 2}`;
                const val = pivotData[row][col] || 0;
                this.store.updateCell(newSheet, id, { rawValue: String(val), computedValue: val });
            });
        });

        this.initTabs();
        this.initGrid();
        this.updateGridDisplay();
    }

    private showCommentDialog() {
        if (!this.activeCell) {
            alert('Please select a cell first.');
            return;
        }

        const sheetName = this.store.getActiveSheetName();
        const cell = this.store.getCell(sheetName, this.activeCell);
        const currentComment = cell.comment || '';

        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100vw';
        overlay.style.height = '100vh';
        overlay.style.background = 'rgba(0,0,0,0.5)';
        overlay.style.zIndex = '2000';

        const dialog = document.createElement('div');
        dialog.className = 'modal-content';
        dialog.style.position = 'fixed';
        dialog.style.top = '50%';
        dialog.style.left = '50%';
        dialog.style.transform = 'translate(-50%, -50%)';
        dialog.style.background = 'white';
        dialog.style.padding = '20px';
        dialog.style.borderRadius = '8px';
        dialog.style.width = '300px';
        dialog.style.zIndex = '2001';

        dialog.innerHTML = `
            <h3 style="margin-top:0">Cell Comment</h3>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Comment:</label>
                <textarea id="input-comment-val" style="width:100%; height:80px; padding:6px; border:1px solid #ccc; border-radius:4px; font-family:inherit; resize:none;">${currentComment}</textarea>
            </div>
            <div style="display:flex; justify-content:flex-end; gap:10px;">
                <button class="btn btn-secondary" id="btn-cancel-comment">Cancel</button>
                <button class="btn btn-primary" id="btn-apply-comment">OK</button>
            </div>
        `;

        document.body.appendChild(overlay);
        document.body.appendChild(dialog);

        const close = () => {
            if (overlay.parentNode) overlay.remove();
            if (dialog.parentNode) dialog.remove();
        };

        dialog.querySelector('#btn-cancel-comment')?.addEventListener('click', close);
        dialog.querySelector('#btn-apply-comment')?.addEventListener('click', () => {
            const val = (dialog.querySelector('#input-comment-val') as HTMLTextAreaElement).value.trim();
            this.store.saveHistory();
            cell.comment = val || undefined;
            this.updateGridDisplay();
            close();
        });
    }

    private deleteComment() {
        if (!this.activeCell) return;
        const cell = this.store.getCell(this.store.getActiveSheetName(), this.activeCell);
        if (cell.comment) {
            this.store.saveHistory();
            delete cell.comment;
            this.updateGridDisplay();
        }
    }

    private showDefineNameDialog() {
        if (this.selectedCells.size === 0) {
            alert('Please select a range first.');
            return;
        }

        const range = this.getSelectionRange();
        const sheetName = this.store.getActiveSheetName();

        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100vw';
        overlay.style.height = '100vh';
        overlay.style.background = 'rgba(0,0,0,0.5)';
        overlay.style.zIndex = '2000';

        const dialog = document.createElement('div');
        dialog.className = 'modal-content';
        dialog.style.position = 'fixed';
        dialog.style.top = '50%';
        dialog.style.left = '50%';
        dialog.style.transform = 'translate(-50%, -50%)';
        dialog.style.background = 'white';
        dialog.style.padding = '20px';
        dialog.style.borderRadius = '8px';
        dialog.style.width = '300px';
        dialog.style.zIndex = '2001';

        dialog.innerHTML = `
            <h3 style="margin-top:0">Define Name</h3>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Name:</label>
                <input type="text" id="input-range-name" placeholder="e.g. SalesData" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
            </div>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Refers to:</label>
                <input type="text" value="${sheetName}!${range}" readonly style="width:100%; padding:6px; border:1px solid #eee; border-radius:4px; background:#f9f9f9;">
            </div>
            <div style="display:flex; justify-content:flex-end; gap:10px;">
                <button class="btn btn-secondary" id="btn-cancel-name">Cancel</button>
                <button class="btn btn-primary" id="btn-apply-name">OK</button>
            </div>
        `;

        document.body.appendChild(overlay);
        document.body.appendChild(dialog);

        const close = () => {
            if (overlay.parentNode) overlay.remove();
            if (dialog.parentNode) dialog.remove();
        };

        dialog.querySelector('#btn-cancel-name')?.addEventListener('click', close);
        dialog.querySelector('#btn-apply-name')?.addEventListener('click', () => {
            const name = (dialog.querySelector('#input-range-name') as HTMLInputElement).value.trim();
            if (!name || !/^[a-zA-Z_][a-zA-Z0-9_]*$/.test(name)) {
                alert('Invalid name. Must start with letter/underscore and contain no spaces.');
                return;
            }

            const metadata = this.store.getMetadata(sheetName);
            if (!metadata.namedRanges) metadata.namedRanges = [];
            
            metadata.namedRanges.push({
                id: Math.random().toString(36).substr(2, 9),
                name,
                range: `${sheetName}!${range}`
            });
            
            close();
        });
    }

    private showDataValidationDialog() {
        if (!this.activeCell) {
            alert('Please select a cell first.');
            return;
        }

        const cell = this.store.getCell(this.store.getActiveSheetName(), this.activeCell);

        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100vw';
        overlay.style.height = '100vh';
        overlay.style.background = 'rgba(0,0,0,0.5)';
        overlay.style.zIndex = '2000';

        const dialog = document.createElement('div');
        dialog.className = 'modal-content';
        dialog.style.position = 'fixed';
        dialog.style.top = '50%';
        dialog.style.left = '50%';
        dialog.style.transform = 'translate(-50%, -50%)';
        dialog.style.background = 'white';
        dialog.style.padding = '20px';
        dialog.style.borderRadius = '8px';
        dialog.style.width = '300px';
        dialog.style.zIndex = '2001';

        dialog.innerHTML = `
            <h3 style="margin-top:0">Data Validation</h3>
            <div style="margin-bottom: 15px;">
                <label style="display:block; margin-bottom:5px; font-size:13px;">Criteria:</label>
                <select id="select-validation-type" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
                    <option value="none">None</option>
                    <option value="list">List (separated by commas)</option>
                    <option value="number">Number (Between)</option>
                    <option value="checkbox">Checkbox</option>
                </select>
            </div>
            <div id="validation-details"></div>
            <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:20px;">
                <button class="btn btn-secondary" id="btn-cancel-val">Cancel</button>
                <button class="btn btn-primary" id="btn-apply-val">OK</button>
            </div>
        `;

        document.body.appendChild(overlay);
        document.body.appendChild(dialog);

        const detailsBox = dialog.querySelector('#validation-details') as HTMLElement;
        const typeSelect = dialog.querySelector('#select-validation-type') as HTMLSelectElement;

        typeSelect.onchange = () => {
            if (typeSelect.value === 'list') {
                detailsBox.innerHTML = `
                    <label style="display:block; margin-bottom:5px; font-size:13px;">Source:</label>
                    <input type="text" id="input-val-criteria" placeholder="e.g. Yes,No,Maybe" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
                `;
            } else if (typeSelect.value === 'number') {
                detailsBox.innerHTML = `
                    <div style="display:flex; gap:10px;">
                        <div>
                            <label style="display:block; margin-bottom:5px; font-size:13px;">Min:</label>
                            <input type="number" id="input-val-min" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
                        </div>
                        <div>
                            <label style="display:block; margin-bottom:5px; font-size:13px;">Max:</label>
                            <input type="number" id="input-val-max" style="width:100%; padding:6px; border:1px solid #ccc; border-radius:4px;">
                        </div>
                    </div>
                `;
            } else {
                detailsBox.innerHTML = '';
            }
        };

        const close = () => {
            if (overlay.parentNode) overlay.remove();
            if (dialog.parentNode) dialog.remove();
        };

        dialog.querySelector('#btn-cancel-val')?.addEventListener('click', close);
        dialog.querySelector('#btn-apply-val')?.addEventListener('click', () => {
            const type = typeSelect.value as any;
            if (type === 'none') {
                cell.validation = undefined;
            } else {
                const validation: DataValidation = {
                    type,
                    allowEmpty: true
                };
                if (type === 'list') {
                    validation.criteria = (dialog.querySelector('#input-val-criteria') as HTMLInputElement).value.split(',').map(s => s.trim());
                } else if (type === 'number') {
                    validation.min = Number((dialog.querySelector('#input-val-min') as HTMLInputElement).value);
                    validation.max = Number((dialog.querySelector('#input-val-max') as HTMLInputElement).value);
                }
                cell.validation = validation;
                
                // If checkbox, auto-initialize value
                if (type === 'checkbox') {
                    cell.rawValue = cell.rawValue === 'TRUE' ? 'TRUE' : 'FALSE';
                    cell.computedValue = cell.rawValue === 'TRUE';
                }
            }
            this.updateGridDisplay();
            close();
        });
    }

    private showCreateTableDialog() {
        if (this.selectedCells.size === 0) {
            alert('Please select a range of cells first.');
            return;
        }

        let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
        this.selectedCells.forEach(id => {
            const { r, c } = this.parseCellId(id);
            minR = Math.min(minR, r);
            maxR = Math.max(maxR, r);
            minC = Math.min(minC, c);
            maxC = Math.max(maxC, c);
        });

        const range = `${COL_NAMES[minC]}${minR}:${COL_NAMES[maxC]}${maxR}`;

        const dialog = document.createElement('div');
        dialog.className = 'modal-overlay';
        dialog.innerHTML = `
            <div class="modal-content" style="width: 350px;">
                <div class="modal-header">
                    <h3>Create Table</h3>
                    <button class="close-btn"><i data-lucide="x"></i></button>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label>Where is the data for your table?</label>
                        <input type="text" id="table-range" value="${range}" readonly>
                    </div>
                    <div class="form-group" style="display: flex; align-items: center; gap: 8px;">
                        <input type="checkbox" id="table-has-header" checked>
                        <label for="table-has-header">My table has headers</label>
                    </div>
                    <div class="form-group">
                        <label>Table Style</label>
                        <select id="table-style" class="tool-select" style="width: 100%;">
                            <option value="blue">Blue</option>
                            <option value="green">Green</option>
                            <option value="orange">Orange</option>
                        </select>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn-secondary" id="btn-cancel-table">Cancel</button>
                    <button class="btn-primary" id="btn-confirm-table">OK</button>
                </div>
            </div>
        `;

        document.body.appendChild(dialog);
        this.refreshIcons();

        const close = () => { if (dialog.parentNode) dialog.remove(); };
        dialog.querySelector('.close-btn')?.addEventListener('click', close);
        dialog.querySelector('#btn-cancel-table')?.addEventListener('click', close);
        
        dialog.querySelector('#btn-confirm-table')?.addEventListener('click', () => {
            const hasHeader = (dialog.querySelector('#table-has-header') as HTMLInputElement).checked;
            const style = (dialog.querySelector('#table-style') as HTMLSelectElement).value as any;
            
            this.createTable(range, hasHeader, style);
            close();
        });
    }

    private createTable(range: string, hasHeader: boolean, style: 'blue' | 'green' | 'orange') {
        const sheetName = this.store.getActiveSheetName();
        const metadata = this.store.getMetadata(sheetName);
        
        const newTable: Table = {
            id: 'table_' + Date.now(),
            range,
            hasHeader,
            style,
            filters: {}
        };

        metadata.tables.push(newTable);
        this.store.saveHistory();
        this.updateGridDisplay();
    }

    private isCellInTable(cellId: string, table: Table): boolean {
        const { r, c } = this.parseCellId(cellId);
        const { startR, startC, endR, endC } = this.parseRange(table.range);
        return r >= startR && r <= endR && c >= startC && c <= endC;
    }

    private parseCellId(id: string) {
        const col = id.match(/[A-Z]+/)?.[0]!;
        const row = parseInt(id.match(/[0-9]+/)?.[0]!);
        return { r: row, c: COL_NAMES.indexOf(col) };
    }

    private parseRange(range: string) {
        const [start, end] = range.split(':');
        const s = this.parseCellId(start);
        const e = this.parseCellId(end);
        return {
            startR: Math.min(s.r, e.r),
            startC: Math.min(s.c, e.c),
            endR: Math.max(s.r, e.r),
            endC: Math.max(s.c, e.c)
        };
    }

    private showFilterMenu(table: Table, colIndex: number, event: MouseEvent) {
        const existingMenu = document.querySelector('.filter-menu');
        if (existingMenu && existingMenu.parentNode) existingMenu.remove();

        const menu = document.createElement('div');
        menu.className = 'filter-menu';
        menu.style.top = event.clientY + 'px';
        menu.style.left = event.clientX + 'px';

        const sheetName = this.store.getActiveSheetName();
        const { startR, startC, endR, endC } = this.parseRange(table.range);
        const dataRowsStart = table.hasHeader ? startR + 1 : startR;
        
        // Get unique values in column
        const values = new Set<string>();
        for (let r = dataRowsStart; r <= endR; r++) {
            const cell = this.store.getCell(sheetName, `${COL_NAMES[startC + colIndex]}${r}`);
            values.add(String(cell.computedValue));
        }

        const sortedValues = Array.from(values).sort();

        menu.innerHTML = `
            <div class="filter-item" id="sort-asc">Sort A to Z</div>
            <div class="filter-item" id="sort-desc">Sort Z to A</div>
            <hr style="margin: 4px 0; border: none; border-top: 1px solid #eee;">
            <input type="text" class="filter-search" placeholder="Search...">
            <div class="filter-list" style="max-height: 200px; overflow-y: auto;">
                <div class="filter-item"><input type="checkbox" id="filter-all" checked> (Select All)</div>
                ${sortedValues.map(v => `
                    <div class="filter-item">
                        <input type="checkbox" class="filter-val" value="${v}" checked>
                        <span>${v || '(Blanks)'}</span>
                    </div>
                `).join('')}
            </div>
            <div style="margin-top: 8px; display: flex; justify-content: flex-end; gap: 8px;">
                <button class="btn-secondary" style="padding: 2px 8px; font-size: 11px;" id="filter-cancel">Cancel</button>
                <button class="btn-primary" style="padding: 2px 8px; font-size: 11px;" id="filter-ok">OK</button>
            </div>
        `;

        document.body.appendChild(menu);

        menu.querySelector('#sort-asc')?.addEventListener('click', () => {
            this.sortTable(table, colIndex, 'asc');
            if (menu.parentNode) menu.remove();
        });

        menu.querySelector('#sort-desc')?.addEventListener('click', () => {
            this.sortTable(table, colIndex, 'desc');
            if (menu.parentNode) menu.remove();
        });

        const okBtn = menu.querySelector('#filter-ok');
        const cancelBtn = menu.querySelector('#filter-cancel');
        const filterAll = menu.querySelector('#filter-all') as HTMLInputElement;
        const searchInput = menu.querySelector('.filter-search') as HTMLInputElement;
        const filterList = menu.querySelector('.filter-list');

        filterAll?.addEventListener('change', () => {
            menu.querySelectorAll('.filter-val').forEach(el => {
                (el as HTMLInputElement).checked = filterAll.checked;
            });
        });

        searchInput?.addEventListener('input', () => {
            const term = searchInput.value.toLowerCase();
            menu.querySelectorAll('.filter-item').forEach(item => {
                if (item.contains(filterAll)) return;
                const text = item.textContent?.toLowerCase() || '';
                (item as HTMLElement).style.display = text.includes(term) ? 'flex' : 'none';
            });
        });

        okBtn?.addEventListener('click', () => {
            const checked = Array.from(menu.querySelectorAll('.filter-val:checked')).map(el => (el as HTMLInputElement).value);
            this.filterTable(table, colIndex, checked.length === values.size ? [] : checked);
            if (menu.parentNode) menu.remove();
        });

        cancelBtn?.addEventListener('click', () => {
            if (menu.parentNode) menu.remove();
        });

        // Close on click outside
        const outsideClick = (e: MouseEvent) => {
            if (!menu.contains(e.target as Node)) {
                if (menu.parentNode) menu.remove();
                window.removeEventListener('mousedown', outsideClick);
            }
        };
        window.addEventListener('mousedown', outsideClick);
    }

    private sortTable(table: Table, colIndex: number, direction: 'asc' | 'desc') {
        this.store.saveHistory();
        const sheetName = this.store.getActiveSheetName();
        const { startR, startC, endR, endC } = this.parseRange(table.range);
        const dataRowsStart = table.hasHeader ? startR + 1 : startR;

        const rows: any[][] = [];
        for (let r = dataRowsStart; r <= endR; r++) {
            const rowData = [];
            for (let c = startC; c <= endC; c++) {
                const cell = this.store.getCell(sheetName, `${COL_NAMES[c]}${r}`);
                rowData.push({
                    formula: cell.formula,
                    rawValue: cell.rawValue,
                    computedValue: cell.computedValue,
                    format: { ...cell.format }
                });
            }
            rows.push(rowData);
        }

        rows.sort((a, b) => {
            const valA = a[colIndex].computedValue;
            const valB = b[colIndex].computedValue;
            if (valA < valB) return direction === 'asc' ? -1 : 1;
            if (valA > valB) return direction === 'asc' ? 1 : -1;
            return 0;
        });

        // Write back
        rows.forEach((rowData, i) => {
            const r = dataRowsStart + i;
            rowData.forEach((cellData: CellData, j: number) => {
                const c = startC + j;
                const targetId = `${COL_NAMES[c]}${r}`;
                // We need to be careful with formulas and references. 
                // For simplicity, we move the raw values and formulas.
                this.store.updateCell(sheetName, targetId, {
                    formula: cellData.formula,
                    rawValue: cellData.rawValue,
                    computedValue: cellData.computedValue,
                    format: { ...cellData.format }
                });
            });
        });

        this.updateGridDisplay();
    }

    private filterTable(table: Table, colIndex: number, allowedValues: string[]) {
        table.filters[colIndex] = allowedValues;
        this.updateGridDisplay();
    }

    private showConditionalFormattingDialog() {
        if (!this.activeCell) {
            alert('Please select a cell first.');
            return;
        }

        const sheetName = this.store.getActiveSheetName();
        const cell = this.store.getCell(sheetName, this.activeCell);

        const dialog = document.createElement('div');
        dialog.style.position = 'fixed';
        dialog.style.top = '50%';
        dialog.style.left = '50%';
        dialog.style.transform = 'translate(-50%, -50%)';
        dialog.style.background = 'white';
        dialog.style.padding = '20px';
        dialog.style.boxShadow = '0 0 20px rgba(0,0,0,0.3)';
        dialog.style.zIndex = '1000';
        dialog.style.borderRadius = '8px';
        dialog.style.width = '350px';

        dialog.innerHTML = `
            <h3 style="margin-top:0">Conditional Formatting for ${this.activeCell}</h3>
            <div style="margin-bottom:15px">
                <label style="display:block;font-size:12px;margin-bottom:5px">Condition</label>
                <select id="cond-op" style="width:100%;padding:5px">
                    <option value="gt">Greater Than</option>
                    <option value="lt">Less Than</option>
                    <option value="eq">Equal To</option>
                    <option value="contains">Text Contains</option>
                </select>
            </div>
            <div style="margin-bottom:15px">
                <label style="display:block;font-size:12px;margin-bottom:5px">Value</label>
                <input type="text" id="cond-val" style="width:100%;padding:5px" placeholder="Value to compare">
            </div>
            <div style="margin-bottom:15px">
                <label style="display:block;font-size:12px;margin-bottom:5px">Format</label>
                <div style="display:flex;gap:10px">
                    <input type="color" id="cond-bg" value="#ffcccc" title="Background Color">
                    <input type="color" id="cond-text" value="#990000" title="Text Color">
                </div>
            </div>
            <div style="display:flex;justify-content:flex-end;gap:10px">
                <button id="cond-cancel" style="padding:5px 15px">Cancel</button>
                <button id="cond-clear" style="padding:5px 15px;background:#eee;border:1px solid #ccc">Clear Rules</button>
                <button id="cond-save" style="padding:5px 15px;background:#1a73e8;color:white;border:none;border-radius:4px">Add Rule</button>
            </div>
        `;

        document.body.appendChild(dialog);

        const close = () => { if (dialog.parentNode) dialog.remove(); };

        dialog.querySelector('#cond-cancel')?.addEventListener('click', close);
        
        dialog.querySelector('#cond-clear')?.addEventListener('click', () => {
            this.store.saveHistory();
            cell.conditionalRules = [];
            this.updateGridDisplay();
            close();
        });

        dialog.querySelector('#cond-save')?.addEventListener('click', () => {
            const op = (dialog.querySelector('#cond-op') as HTMLSelectElement).value as any;
            const val = (dialog.querySelector('#cond-val') as HTMLInputElement).value;
            const bg = (dialog.querySelector('#cond-bg') as HTMLInputElement).value;
            const text = (dialog.querySelector('#cond-text') as HTMLInputElement).value;

            this.store.saveHistory();
            cell.conditionalRules.push({
                operator: op,
                value: val,
                format: {
                    backgroundColor: bg,
                    color: text
                }
            });

            this.updateGridDisplay();
            close();
        });
    }

    private activateFormatPainter() {
        this.toggleFormatPainter();
    }

    private toggleFormatPainter() {
        if (!this.activeCell) return;
        
        const sheetName = this.store.getActiveSheetName();
        const cell = this.store.getCell(sheetName, this.activeCell);
        this.formatPainterSource = { ...cell.format };
        this.isFormatPainterActive = true;
        
        const btn = document.getElementById('btn-format-painter');
        if (btn) btn.classList.add('active');
        this.grid.style.cursor = 'crosshair';
    }

    private applyFormatPainter() {
        if (!this.formatPainterSource) return;
        
        const sheetName = this.store.getActiveSheetName();
        this.store.saveHistory();
        
        this.selectedCells.forEach(id => {
            const cell = this.store.getCell(sheetName, id);
            cell.format = { ...this.formatPainterSource! };
        });
        
        this.isFormatPainterActive = false;
        this.formatPainterSource = null;
        
        const btn = document.getElementById('btn-format-painter');
        if (btn) btn.classList.remove('active');
        this.grid.style.cursor = 'default';
        
        this.updateGridDisplay();
    }

    private setTextColor(color: string) {
        this.commitActiveEdit();
        if (this.selectedCells.size === 0) return;
        const sheetName = this.store.getActiveSheetName();
        this.store.saveHistory();
        this.selectedCells.forEach(id => {
            const cell = this.store.getCell(sheetName, id);
            cell.format.color = color;
        });

        const btn = document.getElementById('btn-text-color');
        if (btn) {
            const span = btn.querySelector('span');
            if (span) span.style.borderBottomColor = color;
        }

        this.updateGridDisplay();
    }

    private setFillColor(color: string) {
        this.commitActiveEdit();
        if (this.selectedCells.size === 0) return;
        const sheetName = this.store.getActiveSheetName();
        this.store.saveHistory();
        this.selectedCells.forEach(id => {
            const cell = this.store.getCell(sheetName, id);
            cell.format.backgroundColor = color;
        });

        const btn = document.getElementById('btn-fill-color');
        if (btn) {
            const span = btn.querySelector('span');
            if (span) span.style.borderBottomColor = color;
        }

        this.updateGridDisplay();
    }

    private toggleFormat(type: keyof CellFormat, value?: any) {
        this.commitActiveEdit();
        if (this.selectedCells.size === 0) return;
        const sheetName = this.store.getActiveSheetName();
        this.store.saveHistory();
        
        if (value !== undefined) {
            this.selectedCells.forEach(id => {
                const cell = this.store.getCell(sheetName, id);
                (cell.format as any)[type] = value;
            });
        } else {
            const baseState = this.activeCell ? !(this.store.getCell(sheetName, this.activeCell).format as any)[type] : true;
            this.selectedCells.forEach(id => {
                const cell = this.store.getCell(sheetName, id);
                (cell.format as any)[type] = baseState;
            });
        }
        
        this.updateGridDisplay();
        this.updateToolbarState();
    }

    private setAlignment(align: 'left' | 'center' | 'right') {
        this.commitActiveEdit();
        if (this.selectedCells.size === 0) return;
        const sheetName = this.store.getActiveSheetName();
        this.store.saveHistory();
        
        this.selectedCells.forEach(id => {
            const cell = this.store.getCell(sheetName, id);
            cell.format.textAlign = align;
        });
        
        this.updateGridDisplay();
        this.updateToolbarState();
    }

    private setCellType(type: 'text' | 'number' | 'currency') {
        this.commitActiveEdit();
        if (this.selectedCells.size === 0) return;
        const sheetName = this.store.getActiveSheetName();
        this.store.saveHistory();
        
        this.selectedCells.forEach(id => {
            const cell = this.store.getCell(sheetName, id);
            cell.format.type = type;
        });
        
        this.updateGridDisplay();
        this.updateToolbarState();
    }

    private clipboard: { id: string, data: { formula: string, rawValue: string, computedValue: any, format: CellFormat } }[] = [];

    private copyToClipboard() {
        this.commitActiveEdit();
        if (this.selectedCells.size === 0) return;
        const sheetName = this.store.getActiveSheetName();
        this.clipboard = [];
        
        this.selectedCells.forEach(id => {
            const cell = this.store.getCell(sheetName, id);
            this.clipboard.push({
                id,
                data: {
                    formula: cell.formula,
                    rawValue: cell.rawValue,
                    computedValue: cell.computedValue,
                    format: { ...cell.format }
                }
            });
        });
    }

    private cutToClipboard() {
        this.commitActiveEdit();
        if (this.selectedCells.size === 0) return;
        this.copyToClipboard();
        this.store.saveHistory();
        const sheetName = this.store.getActiveSheetName();
        this.selectedCells.forEach(id => {
            this.commitValue(id, '', sheetName, true);
        });
        this.updateGridDisplay();
    }

    private pasteFromClipboard(mode: 'all' | 'values' | 'formats' | 'formulas' = 'all') {
        this.commitActiveEdit();
        if (!this.activeCell || this.clipboard.length === 0) return;
        this.store.saveHistory();
        const sheetName = this.store.getActiveSheetName();

        const ids = this.clipboard.map(c => c.id);
        const minCol = ids.map(id => colToIdx(id.match(/[A-Z]+/)?.[0]!)).reduce((a, b) => Math.min(a, b));
        const minRow = ids.map(id => parseInt(id.match(/[0-9]+/)?.[0]!)).reduce((a, b) => Math.min(a, b));

        const targetCol = colToIdx(this.activeCell.match(/[A-Z]+/)?.[0]!);
        const targetRow = parseInt(this.activeCell.match(/[0-9]+/)?.[0]!);

        this.clipboard.forEach(item => {
            const itemCol = colToIdx(item.id.match(/[A-Z]+/)?.[0]!);
            const itemRow = parseInt(item.id.match(/[0-9]+/)?.[0]!);

            const offsetCol = itemCol - minCol;
            const offsetRow = itemRow - minRow;

            const destColIdx = targetCol + offsetCol;
            const destRow = targetRow + offsetRow;

            if (destColIdx < DEFAULT_COLS && destRow <= DEFAULT_ROWS) {
                const destId = `${COL_NAMES[destColIdx]}${destRow}`;
                const cell = this.store.getCell(sheetName, destId);
                
                if (mode === 'all' || mode === 'formulas') {
                    cell.formula = item.data.formula;
                    cell.rawValue = item.data.rawValue;
                } else if (mode === 'values') {
                    cell.formula = '';
                    cell.rawValue = String(item.data.computedValue); 
                }

                if (mode === 'all' || mode === 'formats') {
                    cell.format = { ...item.data.format };
                }
                
                this.recalculate(destId);
            }
        });

        this.updateGridDisplay();
        this.updateToolbarState();
    }

    private showPasteSpecialDialog() {
        if (this.clipboard.length === 0) {
            alert('Clipboard is empty.');
            return;
        }

        const dialog = document.createElement('div');
        dialog.className = 'modal-overlay';
        dialog.style.zIndex = '2000';
        dialog.innerHTML = `
            <div class="modal-content" style="max-width: 300px;">
                <div class="modal-header">
                    <h3>Paste Special</h3>
                    <span class="close-btn">&times;</span>
                </div>
                <div class="modal-body">
                    <div style="display: flex; flex-direction: column; gap: 10px;">
                        <label><input type="radio" name="paste-mode" value="all" checked> All</label>
                        <label><input type="radio" name="paste-mode" value="values"> Values</label>
                        <label><input type="radio" name="paste-mode" value="formats"> Formats</label>
                        <label><input type="radio" name="paste-mode" value="formulas"> Formulas</label>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="btn-cancel-paste">Cancel</button>
                    <button class="btn btn-primary" id="btn-confirm-paste">Paste</button>
                </div>
            </div>
        `;
        document.body.appendChild(dialog);

        const close = () => { if (dialog.parentNode) dialog.remove(); };
        dialog.querySelector('.close-btn')?.addEventListener('click', close);
        dialog.querySelector('#btn-cancel-paste')?.addEventListener('click', close);

        dialog.querySelector('#btn-confirm-paste')?.addEventListener('click', () => {
            const mode = (dialog.querySelector('input[name="paste-mode"]:checked') as HTMLInputElement).value as any;
            this.pasteFromClipboard(mode);
            close();
        });
    }

    private autoSum() {
        if (!this.activeCell) return;
        this.store.saveHistory();
        
        const col = this.activeCell.match(/[A-Z]+/)?.[0]!;
        const row = parseInt(this.activeCell.match(/[0-9]+/)?.[0]!);
        const colIdx = COL_NAMES.indexOf(col);
        const sheetName = this.store.getActiveSheetName();
        
        // Look up for numbers
        let startRow = row - 1;
        while (startRow >= 1) {
            const val = this.store.getCell(sheetName, `${col}${startRow}`).computedValue;
            if (isNaN(Number(val)) || val === '') break;
            startRow--;
        }
        startRow++;

        if (startRow < row) {
            const range = `${col}${startRow}:${col}${row - 1}`;
            this.commitValue(this.activeCell, `=SUM(${range})`);
        } else {
            // Look left
            let startColIdx = colIdx - 1;
            while (startColIdx >= 0) {
                const val = this.store.getCell(sheetName, `${COL_NAMES[startColIdx]}${row}`).computedValue;
                if (isNaN(Number(val)) || val === '') break;
                startColIdx--;
            }
            startColIdx++;
            
            if (startColIdx < colIdx) {
                const range = `${COL_NAMES[startColIdx]}${row}:${COL_NAMES[colIdx - 1]}${row}`;
                this.commitValue(this.activeCell, `=SUM(${range})`);
            }
        }
    }

    private handleUndo() {
        this.commitActiveEdit();
        if (this.store.undo()) {
            this.initTabs();
            this.initGrid();
            if (this.activeCell) this.selectCell(this.activeCell);
        }
    }

    private handleRedo() {
        this.commitActiveEdit();
        if (this.store.redo()) {
            this.initTabs();
            this.initGrid();
            if (this.activeCell) this.selectCell(this.activeCell);
        }
    }

    private refreshIcons() {
        createIcons({
            icons: {
                FilePlus, Upload, Download, FileText, Code, Undo2, Redo2, Copy, ClipboardPaste, ClipboardList, Paintbrush, Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight, AlignStartVertical, AlignCenterVertical, AlignEndVertical, WrapText, Combine, Rows, Columns, Search, PlusSquare, Edit3, Link, Table, LineChart, BarChart, Sigma, SortAsc, SortDesc, Filter, CheckCircle, Lock, Unlock, Grid, LocateFixed, Keyboard, HelpCircle, PaintBucket, Type, X, Plus, Printer, Square, ListPlus, LayoutGrid, Maximize
            }
        });
    }

    private showShortcutsDialog() {
        const dialog = document.createElement('div');
        dialog.className = 'modal-overlay';
        
        dialog.innerHTML = `
            <div class="modal-content" style="min-width: 500px;">
                <div class="modal-header">
                    <span>Keyboard Shortcuts</span>
                    <span class="modal-close"><i data-lucide="x"></i></span>
                </div>
                <div class="modal-body" style="max-height: 400px; overflow-y: auto;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                        <thead>
                            <tr style="border-bottom: 2px solid #eee; text-align: left;">
                                <th style="padding: 8px;">Shortcut</th>
                                <th style="padding: 8px;">Action</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>B</kbd></td>
                                <td style="padding: 8px;">Bold</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>I</kbd></td>
                                <td style="padding: 8px;">Italic</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>U</kbd></td>
                                <td style="padding: 8px;">Underline</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>C</kbd></td>
                                <td style="padding: 8px;">Copy</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>X</kbd></td>
                                <td style="padding: 8px;">Cut</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>V</kbd></td>
                                <td style="padding: 8px;">Paste</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>Z</kbd></td>
                                <td style="padding: 8px;">Undo</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>Y</kbd></td>
                                <td style="padding: 8px;">Redo</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>A</kbd></td>
                                <td style="padding: 8px;">Select All</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>F</kbd></td>
                                <td style="padding: 8px;">Find & Replace</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>Home</kbd></td>
                                <td style="padding: 8px;">Go to A1</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Ctrl</kbd> + <kbd>End</kbd></td>
                                <td style="padding: 8px;">Go to last cell</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>F2</kbd> or <kbd>Enter</kbd></td>
                                <td style="padding: 8px;">Edit cell</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Tab</kbd> / <kbd>Shift</kbd> + <kbd>Tab</kbd></td>
                                <td style="padding: 8px;">Move Right / Left</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Enter</kbd> / <kbd>Shift</kbd> + <kbd>Enter</kbd></td>
                                <td style="padding: 8px;">Move Down / Up</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Arrows</kbd></td>
                                <td style="padding: 8px;">Navigate</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Shift</kbd> + <kbd>Arrows</kbd></td>
                                <td style="padding: 8px;">Expand selection</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Delete</kbd> / <kbd>Backspace</kbd></td>
                                <td style="padding: 8px;">Clear cell(s)</td>
                            </tr>
                            <tr style="border-bottom: 1px solid #eee;">
                                <td style="padding: 8px;"><kbd>Esc</kbd></td>
                                <td style="padding: 8px;">Cancel editing</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div class="modal-footer">
                    <button class="btn-primary" id="btn-close-shortcuts">Close</button>
                </div>
            </div>
        `;

        document.body.appendChild(dialog);
        this.refreshIcons();

        const close = () => { if (dialog.parentNode) dialog.remove(); };
        dialog.querySelector('.modal-close')?.addEventListener('click', close);
        dialog.querySelector('#btn-close-shortcuts')?.addEventListener('click', close);
        dialog.addEventListener('click', (e) => {
            if (e.target === dialog) close();
        });
    }

    private showFormulaHelpDialog() {
        const dialog = document.createElement('div');
        dialog.className = 'modal-overlay';
        
        const formulas = [
            { name: 'SUM', description: 'Adds all the numbers in a range of cells.', example: '=SUM(A1:A10)', category: 'Math' },
            { name: 'AVERAGE', description: 'Returns the average (arithmetic mean) of the arguments.', example: '=AVERAGE(B1:B20)', category: 'Statistical' },
            { name: 'MIN', description: 'Returns the smallest value in a set of values.', example: '=MIN(C1:C50)', category: 'Statistical' },
            { name: 'MAX', description: 'Returns the largest value in a set of values.', example: '=MAX(D1:D50)', category: 'Statistical' },
            { name: 'COUNT', description: 'Counts the number of cells that contain numbers.', example: '=COUNT(E1:E100)', category: 'Statistical' },
            { name: 'VLOOKUP', description: 'Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify.', example: '=VLOOKUP("Apple", A1:C10, 2, FALSE)', category: 'Lookup' },
            { name: 'IF', description: 'Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.', example: '=IF(A1>10, "High", "Low")', category: 'Logical' },
            { name: 'CONCATENATE', description: 'Joins two or more text strings into one string.', example: '=CONCATENATE(A1, " ", B1)', category: 'Text' },
            { name: 'UPPER', description: 'Converts text to uppercase.', example: '=UPPER(A1)', category: 'Text' },
            { name: 'LOWER', description: 'Converts text to lowercase.', example: '=LOWER(A1)', category: 'Text' },
            { name: 'TRIM', description: 'Removes all spaces from text except for single spaces between words.', example: '=TRIM(A1)', category: 'Text' },
            { name: 'NOW', description: 'Returns the current date and time.', example: '=NOW()', category: 'Date & Time' },
            { name: 'TODAY', description: 'Returns the current date.', example: '=TODAY()', category: 'Date & Time' },
            { name: 'SPARKLINE', description: 'Creates a small chart within a single cell.', example: '=SPARKLINE(A1:A10)', category: 'Visual' }
        ];

        dialog.innerHTML = `
            <div class="modal-content" style="min-width: 600px;">
                <div class="modal-header">
                    <span>Formula Help</span>
                    <span class="modal-close"><i data-lucide="x"></i></span>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <input type="text" id="formula-search" placeholder="Search formulas..." style="width: 100%; padding: 10px; border-radius: 4px; border: 1px solid #ddd;">
                    </div>
                    <div id="formula-list" style="max-height: 400px; overflow-y: auto; margin-top: 10px; display: flex; flex-direction: column; gap: 10px;">
                        ${this.generateFormulaList(formulas)}
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn-primary" id="btn-close-formula-help">Close</button>
                </div>
            </div>
        `;

        document.body.appendChild(dialog);
        this.refreshIcons();

        const searchInput = dialog.querySelector('#formula-search') as HTMLInputElement;
        const formulaList = dialog.querySelector('#formula-list') as HTMLElement;

        searchInput.addEventListener('input', () => {
            const query = searchInput.value.toLowerCase();
            const filtered = formulas.filter(f => 
                f.name.toLowerCase().includes(query) || 
                f.description.toLowerCase().includes(query) ||
                f.category.toLowerCase().includes(query)
            );
            formulaList.innerHTML = this.generateFormulaList(filtered);
        });

        const close = () => { if (dialog.parentNode) dialog.remove(); };
        dialog.querySelector('.modal-close')?.addEventListener('click', close);
        dialog.querySelector('#btn-close-formula-help')?.addEventListener('click', close);
        dialog.addEventListener('click', (e) => {
            if (e.target === dialog) close();
        });

        searchInput.focus();
    }

    private generateFormulaList(formulas: any[]): string {
        if (formulas.length === 0) return '<div style="text-align: center; padding: 20px; color: #666;">No formulas found.</div>';
        
        return formulas.map(f => `
            <div class="formula-help-item" style="padding: 12px; border: 1px solid #eee; border-radius: 6px; background: #fcfcfc;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                    <strong style="color: #1a73e8; font-size: 14px;">${f.name}</strong>
                    <span style="font-size: 11px; color: #777; background: #eee; padding: 2px 6px; border-radius: 10px;">${f.category}</span>
                </div>
                <div style="font-size: 13px; color: #333; margin-bottom: 8px;">${f.description}</div>
                <div style="font-family: var(--font-mono); font-size: 12px; background: #f1f3f4; padding: 6px; border-radius: 4px; color: #c5221f;">
                    ${f.example}
                </div>
            </div>
        `).join('');
    }

    private showFindReplaceDialog() {
        const dialog = document.createElement('div');
        dialog.className = 'modal-overlay';
        
        dialog.innerHTML = `
            <div class="modal-content" style="min-width: 450px;">
                <div class="modal-header">
                    <span>Find and Replace</span>
                    <span class="modal-close"><i data-lucide="x"></i></span>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label>Find what:</label>
                        <input type="text" id="find-text" placeholder="Text to search for...">
                    </div>
                    <div class="form-group">
                        <label>Replace with:</label>
                        <input type="text" id="replace-text" placeholder="Replacement text...">
                    </div>
                    <div class="form-row" style="margin-top: 10px;">
                        <div class="form-group" style="flex: 1;">
                            <label>Within:</label>
                            <select id="search-scope">
                                <option value="sheet">Sheet</option>
                                <option value="workbook">Workbook</option>
                            </select>
                        </div>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 8px; margin-top: 10px;">
                        <label style="display: flex; align-items: center; gap: 8px; font-size: 13px; cursor: pointer;">
                            <input type="checkbox" id="match-case"> Match case
                        </label>
                        <label style="display: flex; align-items: center; gap: 8px; font-size: 13px; cursor: pointer;">
                            <input type="checkbox" id="match-entire"> Match entire cell contents
                        </label>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn-secondary" id="btn-find-next">Find Next</button>
                    <button class="btn-secondary" id="btn-replace">Replace</button>
                    <button class="btn-primary" id="btn-replace-all">Replace All</button>
                    <button class="btn-secondary" id="btn-close-find">Close</button>
                </div>
            </div>
        `;

        document.body.appendChild(dialog);
        this.refreshIcons();

        const findInput = dialog.querySelector('#find-text') as HTMLInputElement;
        const replaceInput = dialog.querySelector('#replace-text') as HTMLInputElement;
        const scopeSelect = dialog.querySelector('#search-scope') as HTMLSelectElement;
        const matchCaseCheck = dialog.querySelector('#match-case') as HTMLInputElement;
        const matchEntireCheck = dialog.querySelector('#match-entire') as HTMLInputElement;

        const close = () => { if (dialog.parentNode) dialog.remove(); };

        dialog.querySelector('.modal-close')?.addEventListener('click', close);
        dialog.querySelector('#btn-close-find')?.addEventListener('click', close);

        let lastFoundCell: { sheet: string, id: string } | null = null;

        const findNext = (select = true) => {
            const findText = findInput.value;
            if (!findText) return null;

            const scope = scopeSelect.value as 'sheet' | 'workbook';
            const matchCase = matchCaseCheck.checked;
            const matchEntire = matchEntireCheck.checked;

            const sheetsToSearch = scope === 'sheet' 
                ? [this.store.getActiveSheetName()] 
                : this.store.getSheets();

            const allCells: { sheet: string, id: string, value: string }[] = [];
            sheetsToSearch.forEach(sheetName => {
                const sheet = this.store.getWorkbookData().get(sheetName) as Map<string, CellData>;
                sheet.forEach((cell, id) => {
                    const val = cell.formula || String(cell.rawValue || '');
                    if (val) {
                        allCells.push({ sheet: sheetName, id, value: val });
                    }
                });
            });

            allCells.sort((a, b) => {
                if (a.sheet !== b.sheet) {
                    const sheets = this.store.getSheets();
                    return sheets.indexOf(a.sheet) - sheets.indexOf(b.sheet);
                }
                const rA = parseInt(a.id.match(/[0-9]+/)?.[0]!);
                const rB = parseInt(b.id.match(/[0-9]+/)?.[0]!);
                if (rA !== rB) return rA - rB;
                return colToIdx(a.id.match(/[A-Z]+/)?.[0]!) - colToIdx(b.id.match(/[A-Z]+/)?.[0]!);
            });

            if (allCells.length === 0) {
                alert('No data found to search.');
                return null;
            }

            let startIndex = 0;
            if (lastFoundCell) {
                startIndex = allCells.findIndex(c => c.sheet === lastFoundCell!.sheet && c.id === lastFoundCell!.id) + 1;
            }

            for (let i = 0; i < allCells.length; i++) {
                const idx = (startIndex + i) % allCells.length;
                const cell = allCells[idx];
                let cellVal = cell.value;
                let searchVal = findText;

                if (!matchCase) {
                    cellVal = cellVal.toLowerCase();
                    searchVal = searchVal.toLowerCase();
                }

                const isMatch = matchEntire ? cellVal === searchVal : cellVal.includes(searchVal);

                if (isMatch) {
                    lastFoundCell = { sheet: cell.sheet, id: cell.id };
                    if (select) {
                        if (this.store.getActiveSheetName() !== cell.sheet) {
                            this.store.setActiveSheet(cell.sheet);
                            this.initTabs();
                            this.initGrid();
                        }
                        this.selectCell(cell.id);
                    }
                    return cell;
                }
            }
            alert('No matches found.');
            return null;
        };

        dialog.querySelector('#btn-find-next')?.addEventListener('click', () => findNext());

        dialog.querySelector('#btn-replace')?.addEventListener('click', () => {
            const current = findNext(false);
            if (current) {
                const findText = findInput.value;
                const replaceText = replaceInput.value;
                const matchCase = matchCaseCheck.checked;
                const matchEntire = matchEntireCheck.checked;

                let cellVal = current.value;
                let searchVal = findText;
                if (!matchCase) {
                    cellVal = cellVal.toLowerCase();
                    searchVal = searchVal.toLowerCase();
                }
                
                const isMatch = matchEntire ? cellVal === searchVal : cellVal.includes(searchVal);
                if (isMatch) {
                    const cell = this.store.getCell(current.sheet, current.id);
                    let newValue: string;
                    const sourceVal = cell.formula || String(cell.rawValue || '');
                    if (matchEntire) {
                        newValue = replaceText;
                    } else {
                        const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi');
                        newValue = sourceVal.replace(regex, replaceText);
                    }
                    
                    this.commitValue(current.id, newValue, current.sheet);
                    if (current.sheet === this.store.getActiveSheetName()) {
                        this.selectCell(current.id);
                    }
                }
            }
        });

        dialog.querySelector('#btn-replace-all')?.addEventListener('click', () => {
            const findText = findInput.value;
            if (!findText) return;
            const replaceText = replaceInput.value;
            const scope = scopeSelect.value as 'sheet' | 'workbook';
            const matchCase = matchCaseCheck.checked;
            const matchEntire = matchEntireCheck.checked;

            const sheetsToSearch = scope === 'sheet' 
                ? [this.store.getActiveSheetName()] 
                : this.store.getSheets();

            let count = 0;
            this.store.saveHistory();

            sheetsToSearch.forEach(sheetName => {
                const sheet = this.store.getWorkbookData().get(sheetName) as Map<string, CellData>;
                sheet.forEach((cell, id) => {
                    let cellVal = cell.formula || String(cell.rawValue || '');
                    let searchVal = findText;

                    if (!matchCase) {
                        cellVal = cellVal.toLowerCase();
                        searchVal = searchVal.toLowerCase();
                    }

                    const isMatch = matchEntire ? cellVal === searchVal : cellVal.includes(searchVal);

                    if (isMatch) {
                        let newValue: string;
                        const sourceVal = cell.formula || String(cell.rawValue || '');
                        if (matchEntire) {
                            newValue = replaceText;
                        } else {
                            const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi');
                            newValue = sourceVal.replace(regex, replaceText);
                        }
                        this.commitValue(id, newValue, sheetName, true);
                        count++;
                    }
                });
            });

            this.updateGridDisplay();
            alert(`Replaced ${count} occurrences.`);
            close();
        });
    }

    private showLinkCellDialog() {
        if (!this.activeCell) {
            alert('Please select a cell first.');
            return;
        }

        const dialog = document.createElement('div');
        dialog.className = 'modal-overlay';
        
        const sheetNames = this.store.getSheetNames();
        const sheetOptions = sheetNames.map(name => `<option value="${name}">${name}</option>`).join('');

        dialog.innerHTML = `
            <div class="modal-content" style="min-width: 350px;">
                <div class="modal-header">
                    <span>Link Cell</span>
                    <span class="modal-close"><i data-lucide="x"></i></span>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label>Target Sheet:</label>
                        <select id="link-sheet">
                            ${sheetOptions}
                        </select>
                    </div>
                    <div class="form-group">
                        <label>Target Cell (e.g., A1):</label>
                        <input type="text" id="link-cell-id" placeholder="Enter cell address...">
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn-primary" id="btn-create-link">Create Link</button>
                    <button class="btn-secondary" id="btn-close-link">Cancel</button>
                </div>
            </div>
        `;

        document.body.appendChild(dialog);
        this.refreshIcons();

        const sheetSelect = dialog.querySelector('#link-sheet') as HTMLSelectElement;
        const cellInput = dialog.querySelector('#link-cell-id') as HTMLInputElement;

        const close = () => { if (dialog.parentNode) dialog.remove(); };

        dialog.querySelector('.modal-close')?.addEventListener('click', close);
        dialog.querySelector('#btn-close-link')?.addEventListener('click', close);

        dialog.querySelector('#btn-create-link')?.addEventListener('click', () => {
            const targetSheet = sheetSelect.value;
            const targetCell = cellInput.value.toUpperCase().trim();

            if (!targetCell.match(/^[A-Z]+[0-9]+$/)) {
                alert('Invalid cell address. Please use format like A1, B2, etc.');
                return;
            }

            const targetSheetName = targetSheet.includes(' ') ? `'${targetSheet}'` : targetSheet;
            const formula = `=${targetSheetName}!${targetCell}`;
            this.commitValue(this.activeCell!, formula);
            this.updateGridDisplay();
            this.selectCell(this.activeCell!);
            close();
        });
    }

    private insertSparkline() {
        if (!this.activeCell) return;
        this.commitValue(this.activeCell, '=SPARKLINE(A1:A5)');
        const cellEl = this.grid.querySelector(`.grid-cell[data-id="${this.activeCell}"]`) as HTMLElement;
        if (cellEl) this.startEditing(cellEl);
    }

    private toggleTextWrap() {
        if (this.selectedCells.size === 0) return;
        this.store.saveHistory();
        this.selectedCells.forEach(id => {
            const cell = this.store.getCell(this.store.getActiveSheetName(), id);
            cell.format.wrapText = !cell.format.wrapText;
        });
        this.updateGridDisplay();
    }

    private mergeCells() {
        // Basic merge: just hide other cells in range and expand the first one
        // This is a simplified version for now
        if (this.selectedCells.size < 2) return;
        alert('Cell merging is currently in development. It will be available in a future update.');
    }

    private sortRange(direction: 'asc' | 'desc') {
        if (this.selectedCells.size === 0) return;
        this.store.saveHistory();
        
        const sheetName = this.store.getActiveSheetName();
        const cells = Array.from(this.selectedCells);
        
        // Group by column
        const colMap = new Map<string, string[]>();
        cells.forEach(id => {
            const col = id.match(/[A-Z]+/)?.[0]!;
            if (!colMap.has(col)) colMap.set(col, []);
            colMap.get(col)!.push(id);
        });

        colMap.forEach((ids, col) => {
            const sortedIds = [...ids].sort((a, b) => {
                const rowA = parseInt(a.match(/[0-9]+/)?.[0]!);
                const rowB = parseInt(b.match(/[0-9]+/)?.[0]!);
                return rowA - rowB;
            });

            const values = sortedIds.map(id => this.store.getCell(sheetName, id).computedValue);
            values.sort((a, b) => {
                if (direction === 'asc') return String(a).localeCompare(String(b), undefined, { numeric: true });
                return String(b).localeCompare(String(a), undefined, { numeric: true });
            });

            sortedIds.forEach((id, i) => {
                this.commitValue(id, String(values[i]), sheetName, true);
            });
        });
        this.updateGridDisplay();
    }

    private toggleFilter() {
        if (!this.activeCell) return;
        
        const sheetName = this.store.getActiveSheetName();
        const metadata = this.store.getMetadata(sheetName);
        
        // Find if any table already covers the active cell
        const existingTableIdx = metadata.tables.findIndex(t => this.isCellInTable(this.activeCell!, t));
        
        if (existingTableIdx !== -1) {
            const table = metadata.tables[existingTableIdx];
            // Reset row heights for the table range
            const { startR, endR } = this.parseRange(table.range);
            for (let r = startR; r <= endR; r++) {
                this.grid.style.removeProperty(`--row-height-${r}`);
            }
            // Remove filter (table)
            metadata.tables.splice(existingTableIdx, 1);
        } else {
            // Apply filter (default blue table)
            let range = this.getSelectionRange();
            if (!range.includes(':')) {
                range = this.findContiguousRange(this.activeCell);
            }
            
            const newTable: Table = {
                id: 'table-' + Math.random().toString(36).substr(2, 9),
                range: range,
                hasHeader: true,
                style: 'blue',
                filters: {}
            };
            metadata.tables.push(newTable);
        }
        
        this.updateGridDisplay();
    }

    private findContiguousRange(cellId: string): string {
        const { r, c } = this.parseCellId(cellId);
        const sheetName = this.store.getActiveSheetName();
        
        let minR = r, maxR = r, minC = c, maxC = c;
        
        // Find top-left
        while (minR > 1 && this.store.getCell(sheetName, `${COL_NAMES[minC]}${minR - 1}`).rawValue !== '') minR--;
        while (minC > 0 && this.store.getCell(sheetName, `${COL_NAMES[minC - 1]}${minR}`).rawValue !== '') minC--;
        
        // Find bottom-right
        while (maxR < DEFAULT_ROWS && this.store.getCell(sheetName, `${COL_NAMES[maxC]}${maxR + 1}`).rawValue !== '') maxR++;
        while (maxC < DEFAULT_COLS - 1 && this.store.getCell(sheetName, `${COL_NAMES[maxC + 1]}${maxR}`).rawValue !== '') maxC++;
        
        return `${COL_NAMES[minC]}${minR}:${COL_NAMES[maxC]}${maxR}`;
    }

    private async importFromExcel(file: File) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            
            this.store.saveHistory();
            workbook.SheetNames.forEach((name, idx) => {
                if (idx === 0) {
                    this.store.renameSheet(this.store.getActiveSheetName(), name);
                } else {
                    this.store.addSheet(name);
                }
                
                const worksheet = workbook.Sheets[name];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
                
                jsonData.forEach((row, r) => {
                    row.forEach((val, c) => {
                        const id = `${COL_NAMES[c]}${r + 1}`;
                        this.commitValue(id, String(val), name, true);
                    });
                });
            });
            
            this.initTabs();
            this.initGrid();
            this.updateGridDisplay();
        };
        reader.readAsArrayBuffer(file);
    }

    private exportToExcelXLSX() {
        const wb = XLSX.utils.book_new();
        this.store.getSheets().forEach(name => {
            const sheet = this.store.getWorkbookData().get(name)!;
            const data: any[][] = [];
            for (let r = 1; r <= DEFAULT_ROWS; r++) {
                const row: any[] = [];
                for (let c = 0; c < DEFAULT_COLS; c++) {
                    const cell = sheet.get(`${COL_NAMES[c]}${r}`);
                    row.push(cell ? cell.computedValue : '');
                }
                data.push(row);
            }
            const ws = XLSX.utils.aoa_to_sheet(data);
            XLSX.utils.book_append_sheet(wb, ws, name);
        });
        XLSX.writeFile(wb, 'spreadsheet.xlsx');
    }

    private exportToCSV() {
        const name = this.store.getActiveSheetName();
        const sheet = this.store.getWorkbookData().get(name)!;
        let csv = '';
        for (let r = 1; r <= DEFAULT_ROWS; r++) {
            const row: any[] = [];
            for (let c = 0; c < DEFAULT_COLS; c++) {
                const cell = sheet.get(`${COL_NAMES[c]}${r}`);
                let val = cell ? String(cell.computedValue) : '';
                if (val.includes(',') || val.includes('"') || val.includes('\n')) {
                    val = `"${val.replace(/"/g, '""')}"`;
                }
                row.push(val);
            }
            csv += row.join(',') + '\n';
        }
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${name}.csv`;
        a.click();
    }

    private exportToJSON() {
        const data = this.store.serialize();
        const blob = new Blob([data], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'spreadsheet.json';
        a.click();
    }

    private insertChart() {
        if (this.selectedCells.size === 0) {
            alert('Please select a range of data for the chart.');
            return;
        }

        const range = this.getSelectionRangeString();
        const type = prompt('Enter chart type (bar, line, pie, scatter):', 'bar') as any;
        if (!['bar', 'line', 'pie', 'scatter'].includes(type)) return;

        const title = prompt('Enter chart title:', 'New Chart') || 'New Chart';
        
        const sheetName = this.store.getActiveSheetName();
        const metadata = this.store.getMetadata(sheetName);
        
        const chart: ChartData = {
            id: 'chart-' + Date.now(),
            type,
            range,
            title,
            x: 100,
            y: 100,
            width: 400,
            height: 300
        };

        metadata.charts.push(chart);
        this.store.setMetadata(sheetName, metadata);
        this.renderCharts();
    }

    private getSelectionRangeString(): string {
        const cells = Array.from(this.selectedCells);
        if (cells.length === 0) return '';
        
        let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
        cells.forEach(id => {
            const { r, c } = this.parseCellId(id);
            minR = Math.min(minR, r);
            maxR = Math.max(maxR, r);
            minC = Math.min(minC, c);
            maxC = Math.max(maxC, c);
        });

        return `${COL_NAMES[minC]}${minR}:${COL_NAMES[maxC]}${maxR}`;
    }

    private renderCharts() {
        // Remove existing charts from DOM
        this.container.querySelectorAll('.chart-container').forEach(c => c.remove());
        
        const sheetName = this.store.getActiveSheetName();
        const metadata = this.store.getMetadata(sheetName);
        
        metadata.charts.forEach(chart => {
            const chartDiv = document.createElement('div');
            chartDiv.className = 'chart-container';
            chartDiv.style.position = 'absolute';
            chartDiv.style.left = chart.x + 'px';
            chartDiv.style.top = chart.y + 'px';
            chartDiv.style.width = chart.width + 'px';
            chartDiv.style.height = chart.height + 'px';
            chartDiv.style.backgroundColor = 'white';
            chartDiv.style.border = '1px solid #ccc';
            chartDiv.style.boxShadow = '0 2px 10px rgba(0,0,0,0.1)';
            chartDiv.style.zIndex = '100';
            chartDiv.style.padding = '10px';
            chartDiv.style.cursor = 'move';

            const header = document.createElement('div');
            header.style.display = 'flex';
            header.style.justifyContent = 'space-between';
            header.style.marginBottom = '10px';
            header.innerHTML = `<strong>${chart.title}</strong><span class="close-chart" style="cursor:pointer"><i data-lucide="x"></i></span>`;
            chartDiv.appendChild(header);

            const svgContainer = document.createElement('div');
            svgContainer.id = chart.id;
            svgContainer.style.width = '100%';
            svgContainer.style.height = 'calc(100% - 30px)';
            chartDiv.appendChild(svgContainer);

            this.container.appendChild(chartDiv);
        this.refreshIcons();

            // Dragging logic
            let isDragging = false;
            let offset = { x: 0, y: 0 };
            header.onmousedown = (e) => {
                isDragging = true;
                offset = { x: e.clientX - chartDiv.offsetLeft, y: e.clientY - chartDiv.offsetTop };
            };
            window.onmousemove = (e) => {
                if (!isDragging) return;
                chart.x = e.clientX - offset.x;
                chart.y = e.clientY - offset.y;
                chartDiv.style.left = chart.x + 'px';
                chartDiv.style.top = chart.y + 'px';
            };
            window.onmouseup = () => {
                if (isDragging) {
                    isDragging = false;
                    this.store.setMetadata(sheetName, metadata);
                }
            };

            header.querySelector('.close-chart')?.addEventListener('click', () => {
                metadata.charts = metadata.charts.filter(c => c.id !== chart.id);
                this.store.setMetadata(sheetName, metadata);
                chartDiv.remove();
            });

            this.drawD3Chart(chart);
        });
    }

    private drawD3Chart(chart: ChartData) {
        const container = document.getElementById(chart.id);
        if (!container) return;
        
        const width = chart.width - 20;
        const height = chart.height - 50;
        
        const { startR, startC, endR, endC } = this.parseRange(chart.range);
        const sheetName = this.store.getActiveSheetName();
        
        const data: any[] = [];
        for (let r = startR; r <= endR; r++) {
            const row: any = { label: `Row ${r}` };
            for (let c = startC; c <= endC; c++) {
                const val = this.store.getCell(sheetName, `${COL_NAMES[c]}${r}`).computedValue;
                row[`col${c}`] = isNaN(Number(val)) ? 0 : Number(val);
                if (c === startC && isNaN(Number(val))) row.label = String(val);
            }
            data.push(row);
        }

        const svg = d3.select(`#${chart.id}`)
            .append('svg')
            .attr('width', '100%')
            .attr('height', '100%')
            .attr('viewBox', `0 0 ${width} ${height}`);

        const margin = { top: 20, right: 20, bottom: 30, left: 40 };
        const innerWidth = width - margin.left - margin.right;
        const innerHeight = height - margin.top - margin.bottom;

        const g = svg.append('g').attr('transform', `translate(${margin.left},${margin.top})`);

        const x = d3.scaleBand().rangeRound([0, innerWidth]).padding(0.1);
        const y = d3.scaleLinear().rangeRound([innerHeight, 0]);

        x.domain(data.map(d => d.label));
        y.domain([0, d3.max(data, d => d.col1 || d.col0) as number]);

        if (chart.type === 'bar') {
            g.append('g')
                .attr('class', 'axis axis--x')
                .attr('transform', `translate(0,${innerHeight})`)
                .call(d3.axisBottom(x));

            g.append('g')
                .attr('class', 'axis axis--y')
                .call(d3.axisLeft(y).ticks(10));

            g.selectAll('.bar')
                .data(data)
                .enter().append('rect')
                .attr('class', 'bar')
                .attr('x', d => x(d.label)!)
                .attr('y', d => y(d.col1 || d.col0))
                .attr('width', x.bandwidth())
                .attr('height', d => innerHeight - y(d.col1 || d.col0))
                .attr('fill', '#4285f4');
        } else if (chart.type === 'line') {
            const line = d3.line<any>()
                .x(d => x(d.label)! + x.bandwidth() / 2)
                .y(d => y(d.col1 || d.col0));

            g.append('g')
                .attr('class', 'axis axis--x')
                .attr('transform', `translate(0,${innerHeight})`)
                .call(d3.axisBottom(x));

            g.append('g')
                .attr('class', 'axis axis--y')
                .call(d3.axisLeft(y).ticks(10));

            g.append('path')
                .datum(data)
                .attr('fill', 'none')
                .attr('stroke', '#4285f4')
                .attr('stroke-width', 2)
                .attr('d', line);
        } else if (chart.type === 'pie') {
            const radius = Math.min(innerWidth, innerHeight) / 2;
            const pie = d3.pie<any>().value(d => d.col1 || d.col0);
            const arc = d3.arc<any>().innerRadius(0).outerRadius(radius);
            const color = d3.scaleOrdinal(d3.schemeCategory10);

            const pieG = g.append('g')
                .attr('transform', `translate(${innerWidth / 2},${innerHeight / 2})`);

            const arcs = pieG.selectAll('.arc')
                .data(pie(data))
                .enter().append('g')
                .attr('class', 'arc');

            arcs.append('path')
                .attr('d', arc)
                .attr('fill', (d, i) => color(i.toString()));
        }
    }
    private toggleBorder(type: 'all' | 'none' | 'outer' | 'top' | 'bottom' | 'left' | 'right') {
        if (this.selectedCells.size === 0) return;
        this.store.saveHistory();
        const sheetName = this.store.getActiveSheetName();

        const borderStyle = '1px solid black';

        if (type === 'all' || type === 'none') {
            this.selectedCells.forEach(id => {
                const style = type === 'all' ? borderStyle : '';
                this.store.updateCell(sheetName, id, {
                    format: {
                        ...this.store.getCell(sheetName, id).format,
                        borderTop: style,
                        borderBottom: style,
                        borderLeft: style,
                        borderRight: style
                    }
                });
            });
        } else if (type === 'outer') {
            const bounds = this.getSelectionBounds();
            if (bounds) {
                const { startR, startC, endR, endC } = bounds;
                this.selectedCells.forEach(id => {
                    const { r, c } = this.parseCellId(id);
                    const cell = this.store.getCell(sheetName, id);
                    const format = { ...cell.format };
                    if (r === startR) format.borderTop = borderStyle;
                    if (r === endR) format.borderBottom = borderStyle;
                    if (c === startC) format.borderLeft = borderStyle;
                    if (c === endC) format.borderRight = borderStyle;
                    this.store.updateCell(sheetName, id, { format });
                });
            }
        } else {
            this.selectedCells.forEach(id => {
                const cell = this.store.getCell(sheetName, id);
                const format = { ...cell.format };
                const key = `border${type.charAt(0).toUpperCase() + type.slice(1)}` as any;
                (format as any)[key] = (format as any)[key] ? '' : borderStyle;
                this.store.updateCell(sheetName, id, { format });
            });
        }
        this.updateGridDisplay();
    }

    private getSelectionBounds() {
        if (this.selectedCells.size === 0) return null;
        let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
        this.selectedCells.forEach(id => {
            const { r, c } = this.parseCellId(id);
            minR = Math.min(minR, r);
            maxR = Math.max(maxR, r);
            minC = Math.min(minC, c);
            maxC = Math.max(maxC, c);
        });
        return { startR: minR, startC: minC, endR: maxR, endC: maxC };
    }

    private initAutosave() {
        setInterval(() => {
            const data = this.store.serialize();
            localStorage.setItem('spreadsheet_autosave', data);
            console.log('Autosave completed');
        }, 30000); // Autosave every 30 seconds
    }

    private checkAutosave() {
        const savedData = localStorage.getItem('spreadsheet_autosave');
        if (savedData) {
            this.showRestorePrompt(savedData);
        }
    }

    private showRestorePrompt(data: string) {
        const dialog = document.createElement('div');
        dialog.className = 'modal-overlay';
        dialog.style.zIndex = '2000';
        dialog.innerHTML = `
            <div class="modal-content" style="max-width: 400px;">
                <div class="modal-header">
                    <h3>Restore Session?</h3>
                </div>
                <div class="modal-body">
                    <p>A previously saved session was found. Would you like to restore it?</p>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="btn-ignore-restore">Ignore</button>
                    <button class="btn btn-primary" id="btn-confirm-restore">Restore</button>
                </div>
            </div>
        `;
        document.body.appendChild(dialog);
        this.refreshIcons();

        dialog.querySelector('#btn-ignore-restore')?.addEventListener('click', () => {
            localStorage.removeItem('spreadsheet_autosave');
            if (dialog.parentNode) dialog.remove();
        });

        dialog.querySelector('#btn-confirm-restore')?.addEventListener('click', () => {
            if (this.store.deserialize(data)) {
                this.initGrid();
                this.initTabs();
                this.updateGridDisplay();
            }
            if (dialog.parentNode) dialog.remove();
        });
    }
}

export function initSpreadsheet() {
    const container = document.createElement('div');
    container.id = 'app-root';
    document.getElementById('root')!.appendChild(container);
    new SpreadsheetEngine('app-root');
}
