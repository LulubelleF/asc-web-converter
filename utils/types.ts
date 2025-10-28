export interface ParsedRow {
    studentNo?: string;
    sin?: string;
    gov?: string;
    cert?: string;
    name?: string;
    eosDate?: string;
    amount?: string | number;
    [k: string]: any;
}


export interface ParsedTotals {
    count?: number;
    amount?: number;
}


export interface ParsedPage {
    sheetName?: string;
    headerLines?: string [];
    columns?: string [];
    rows: ParsedRow[];
    totals?: ParsedTotals;
}

