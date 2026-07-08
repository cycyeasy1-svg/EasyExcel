import { MoonOutlined, SunOutlined } from "@ant-design/icons";
import type { ICellData, IWorkbookData, IWorksheetData } from "@univerjs/core";
import { App, Spin } from "antd";
import { useCallback, useEffect, useMemo, useState } from "react";
import { applyDarkMode, handler, loadDarkMode } from "../../util/vscode.ts";
import { loadWorkbookBuffer } from "../../util/loadWorkbookContent.ts";
import './ExcelDiff.less';

type DiffKind = 'added' | 'deleted' | 'modified';
type DiffFilter = 'all' | DiffKind;

interface DiffSidePayload {
    label: string;
    buffer?: number[];
    missing?: boolean;
}

interface ExcelDiffPayload {
    ext: string;
    fileName: string;
    sourceLabel: string;
    left: DiffSidePayload;
    right: DiffSidePayload;
}

interface ParsedSide {
    label: string;
    missing: boolean;
    workbook?: IWorkbookData;
}

interface CellDiff {
    key: string;
    address: string;
    kind: DiffKind;
    before: string;
    after: string;
}

interface SheetDiff {
    key: string;
    name: string;
    status: 'normal' | 'added' | 'deleted';
    cells: CellDiff[];
    counts: Record<DiffKind, number>;
}

interface WorkbookDiff {
    sheets: SheetDiff[];
    counts: Record<DiffKind, number>;
}

const emptyCounts = (): Record<DiffKind, number> => ({ added: 0, deleted: 0, modified: 0 });

function colName(ci: number): string {
    let n = ci;
    let name = '';
    do {
        name = String.fromCharCode(65 + (n % 26)) + name;
        n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return name;
}

function cellAddress(ri: number, ci: number): string {
    return `${colName(ci)}${ri + 1}`;
}

function sheetName(workbook: IWorkbookData | undefined, sheetId: string, index: number): string {
    return workbook?.sheets?.[sheetId]?.name || `Sheet ${index + 1}`;
}

function richText(cell: ICellData): string | undefined {
    const dataStream = (cell.p as { body?: { dataStream?: string } } | undefined)?.body?.dataStream;
    return typeof dataStream === 'string' ? dataStream.replace(/\r?\n$/, '') : undefined;
}

function cellText(cell: ICellData | undefined): string {
    if (!cell) return '';
    const value = cell.v == null ? '' : String(cell.v);
    const formula = cell.f == null ? '' : String(cell.f);
    const text = richText(cell) ?? value;
    if (formula && text) return `${formula}\n${text}`;
    return formula || text;
}

function collectCells(sheet: Partial<IWorksheetData> | undefined): Map<string, { ri: number; ci: number; text: string }> {
    const cells = new Map<string, { ri: number; ci: number; text: string }>();
    const rows = sheet?.cellData as Record<string, Record<string, ICellData>> | undefined;
    if (!rows) return cells;
    for (const [rowKey, row] of Object.entries(rows)) {
        const ri = Number(rowKey);
        if (!Number.isFinite(ri) || !row) continue;
        for (const [colKey, cell] of Object.entries(row)) {
            const ci = Number(colKey);
            if (!Number.isFinite(ci)) continue;
            const text = cellText(cell);
            if (text === '') continue;
            cells.set(`${ri}:${ci}`, { ri, ci, text });
        }
    }
    return cells;
}

function addCellDiff(sheet: SheetDiff, diff: CellDiff) {
    sheet.cells.push(diff);
    sheet.counts[diff.kind] += 1;
}

function makeSheetDiff(name: string, key: string, status: SheetDiff['status']): SheetDiff {
    return { key, name, status, cells: [], counts: emptyCounts() };
}

function compareSheet(
    left: Partial<IWorksheetData> | undefined,
    right: Partial<IWorksheetData> | undefined,
    name: string,
    key: string,
): SheetDiff {
    const status: SheetDiff['status'] = left ? (right ? 'normal' : 'deleted') : 'added';
    const sheet = makeSheetDiff(name, key, status);
    const before = collectCells(left);
    const after = collectCells(right);
    const keys = new Set([...before.keys(), ...after.keys()]);

    for (const cellKey of [...keys].sort((a, b) => {
        const [ari, aci] = a.split(':').map(Number);
        const [bri, bci] = b.split(':').map(Number);
        return ari - bri || aci - bci;
    })) {
        const l = before.get(cellKey);
        const r = after.get(cellKey);
        if (l?.text === r?.text) continue;
        const ri = l?.ri ?? r?.ri ?? 0;
        const ci = l?.ci ?? r?.ci ?? 0;
        const kind: DiffKind = l ? (r ? 'modified' : 'deleted') : 'added';
        addCellDiff(sheet, {
            key: cellKey,
            address: cellAddress(ri, ci),
            kind,
            before: l?.text ?? '',
            after: r?.text ?? '',
        });
    }

    return sheet;
}

function sheetEntries(workbook: IWorkbookData | undefined): { id: string; name: string; sheet: Partial<IWorksheetData> }[] {
    if (!workbook) return [];
    return workbook.sheetOrder.map((id, index) => ({
        id,
        name: sheetName(workbook, id, index),
        sheet: workbook.sheets[id],
    }));
}

function compareWorkbooks(left: IWorkbookData | undefined, right: IWorkbookData | undefined): WorkbookDiff {
    const leftEntries = sheetEntries(left);
    const rightEntries = sheetEntries(right);
    const leftByName = new Map(leftEntries.map(entry => [entry.name, entry.sheet]));
    const rightByName = new Map(rightEntries.map(entry => [entry.name, entry.sheet]));
    const orderedNames = [
        ...leftEntries.map(entry => entry.name),
        ...rightEntries.map(entry => entry.name).filter(name => !leftByName.has(name)),
    ];
    const sheets = orderedNames.map(name => compareSheet(leftByName.get(name), rightByName.get(name), name, name));
    const counts = emptyCounts();
    for (const sheet of sheets) {
        counts.added += sheet.counts.added;
        counts.deleted += sheet.counts.deleted;
        counts.modified += sheet.counts.modified;
    }
    return { sheets, counts };
}

async function parseSide(side: DiffSidePayload, ext: string, fileName: string): Promise<ParsedSide> {
    if (side.missing || !side.buffer) {
        return { label: side.label, missing: true };
    }
    const buffer = await loadWorkbookBuffer({ buffer: side.buffer });
    const { loadForUniver } = await import("./univer/loader.ts");
    const { workbookData } = await loadForUniver(buffer, ext, fileName);
    return { label: side.label, missing: false, workbook: workbookData };
}

function visibleCells(sheet: SheetDiff | undefined, filter: DiffFilter): CellDiff[] {
    if (!sheet) return [];
    if (filter === 'all') return sheet.cells;
    return sheet.cells.filter(cell => cell.kind === filter);
}

function ExcelDiffViewer() {
    const [dark, setDark] = useState(loadDarkMode);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [payload, setPayload] = useState<ExcelDiffPayload | null>(null);
    const [diff, setDiff] = useState<WorkbookDiff | null>(null);
    const [activeSheet, setActiveSheet] = useState(0);
    const [filter, setFilter] = useState<DiffFilter>('all');

    useEffect(() => {
        document.body.classList.toggle('easyexcel-dark', dark);
    }, [dark]);

    const toggleDark = useCallback(() => {
        setDark(prev => {
            const next = !prev;
            applyDarkMode(next);
            return next;
        });
    }, []);

    useEffect(() => {
        handler.on("diffOpen", (nextPayload: ExcelDiffPayload) => {
            setPayload(nextPayload);
            setLoading(true);
            setError(null);
            Promise.all([
                parseSide(nextPayload.left, nextPayload.ext, nextPayload.fileName),
                parseSide(nextPayload.right, nextPayload.ext, nextPayload.fileName),
            ]).then(([left, right]) => {
                setDiff(compareWorkbooks(left.workbook, right.workbook));
                setActiveSheet(0);
                setFilter('all');
            }).catch(err => {
                setError((err as Error).message || String(err));
            }).finally(() => setLoading(false));
        }).emit("init");
    }, []);

    const active = diff?.sheets[activeSheet];
    const rows = useMemo(() => visibleCells(active, filter), [active, filter]);
    const total = diff ? diff.counts.added + diff.counts.deleted + diff.counts.modified : 0;

    return (
        <div className="excel-diff-viewer">
            <Spin spinning={loading} fullscreen={true} />
            <header className="excel-diff-header">
                <div className="excel-diff-title-block">
                    <span className="excel-diff-kicker">Excel Diff</span>
                    <strong title={payload?.fileName}>{payload?.fileName ?? 'Workbook'}</strong>
                    <span>{payload?.sourceLabel ?? ''}</span>
                </div>
                <button
                    type="button"
                    className="excel-diff-icon-button"
                    title={dark ? 'Switch to light mode' : 'Switch to dark mode'}
                    onClick={toggleDark}
                >
                    {dark ? <SunOutlined /> : <MoonOutlined />}
                </button>
            </header>

            {error && !loading && (
                <div className="excel-diff-error">
                    <h2>Failed to compare workbook</h2>
                    <p>{error}</p>
                </div>
            )}

            {!error && !loading && diff && (
                <>
                    <section className="excel-diff-summary" aria-label="Diff summary">
                        <div><strong>{total}</strong><span>Total</span></div>
                        <div className="is-added"><strong>{diff.counts.added}</strong><span>Added</span></div>
                        <div className="is-deleted"><strong>{diff.counts.deleted}</strong><span>Deleted</span></div>
                        <div className="is-modified"><strong>{diff.counts.modified}</strong><span>Modified</span></div>
                    </section>

                    <main className="excel-diff-body">
                        <aside className="excel-diff-sheets" aria-label="Sheets">
                            {diff.sheets.map((sheet, index) => {
                                const count = sheet.counts.added + sheet.counts.deleted + sheet.counts.modified;
                                return (
                                    <button
                                        type="button"
                                        key={sheet.key}
                                        className={index === activeSheet ? 'active' : ''}
                                        onClick={() => {
                                            setActiveSheet(index);
                                            setFilter('all');
                                        }}
                                        title={sheet.name}
                                    >
                                        <span>{sheet.name}</span>
                                        <em>{count}</em>
                                    </button>
                                );
                            })}
                        </aside>

                        <section className="excel-diff-content">
                            <div className="excel-diff-toolbar">
                                <div className="excel-diff-sheet-heading">
                                    <h1>{active?.name ?? 'Sheet'}</h1>
                                    {active && active.status !== 'normal' && (
                                        <span className={`sheet-status ${active.status}`}>{active.status}</span>
                                    )}
                                </div>
                                <div className="excel-diff-filters" aria-label="Filter changes">
                                    {(['all', 'added', 'deleted', 'modified'] as DiffFilter[]).map(item => (
                                        <button
                                            key={item}
                                            type="button"
                                            className={filter === item ? 'active' : ''}
                                            onClick={() => setFilter(item)}
                                        >
                                            {item}
                                        </button>
                                    ))}
                                </div>
                            </div>

                            {rows.length === 0 ? (
                                <div className="excel-diff-empty">
                                    <strong>No cell changes</strong>
                                    <span>{filter === 'all' ? 'This sheet has no value changes.' : 'No changes match this filter.'}</span>
                                </div>
                            ) : (
                                <div className="excel-diff-table-wrap">
                                    <table className="excel-diff-table">
                                        <thead>
                                            <tr>
                                                <th>Cell</th>
                                                <th>Change</th>
                                                <th>{payload?.left.label ?? 'Before'}</th>
                                                <th>{payload?.right.label ?? 'After'}</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {rows.map(cell => (
                                                <tr key={cell.key} className={`is-${cell.kind}`}>
                                                    <td className="cell-address">{cell.address}</td>
                                                    <td><span className="change-pill">{cell.kind}</span></td>
                                                    <td><pre>{cell.before || '(empty)'}</pre></td>
                                                    <td><pre>{cell.after || '(empty)'}</pre></td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            )}
                        </section>
                    </main>
                </>
            )}
        </div>
    );
}

export default function ExcelDiff() {
    return (
        <App className="excel-app" message={{ top: 16 }}>
            <ExcelDiffViewer />
        </App>
    );
}
