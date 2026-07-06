import { MinusOutlined, MoonOutlined, PlusOutlined, SunOutlined } from "@ant-design/icons";
import { App, Button, Modal, Radio, Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler, loadDarkMode, applyDarkMode } from "../../util/vscode.ts";
import { loadOfficeBuffer } from "../../util/loadOfficeContent.ts";
import './Excel.less';
import { MIN_VIEW_COLS, MIN_VIEW_ROWS } from "./excel_meta.ts";
import { detectCsvEncoding } from "./csvEncoding.ts";
import { loadSheets } from "./excel_reader.ts";
import { export_xlsx, exportSaveAs, buildFormattingSnapshot, hasFormattingChanged } from "./excel_writer.ts";
import Spreadsheet from './x-spreadsheet/index';
import FindReplacePanel from './FindReplacePanel';
import { parseSpreadsheetLink } from './excel_hyperlink';
import { initExcelLocale, t } from './excel_i18n';
import { getConfigs } from '../../util/vscodeConfig.ts';
import type { UniverAdapter } from './univer/adapter';

initExcelLocale();

type SpreadsheetEngine = 'legacy' | 'univer';

/** 引擎选择：URL 参数（开发调试用）优先，其次宿主 {{configs}}，默认 legacy */
function resolveEngine(): SpreadsheetEngine {
    try {
        const fromUrl = new URLSearchParams(window.location.search).get('engine');
        if (fromUrl === 'univer' || fromUrl === 'legacy') return fromUrl;
    } catch {
        // ignore
    }
    return getConfigs()?.engine === 'univer' ? 'univer' : 'legacy';
}

const isZhLang = () => /^zh/i.test(getConfigs()?.language ?? '');

/** 引擎在 webview 生命周期内不变，模块级解析一次（render 中不读 ref） */
const ENGINE: SpreadsheetEngine = resolveEngine();

/** Univer 编辑会话的保存上下文（initUniver 组装，保存后就地更新） */
interface UniverSaveState {
    loadResult: import('./univer/loader').UniverLoadResult;
    baseline: import('@univerjs/core').IWorkbookData;
    session: import('./univer/adapter').UniverEditSession;
}

type ExcelViewState = { ri: number; ci: number; sheetIndex: number };

const EXCEL_VIEW_STATE_SUFFIX = '-excel-view';
const DEFAULT_ZOOM_PERCENT = 100;
const ZOOM_OPTIONS = [50, 75, 90, 100, 125, 150, 200];

function clampZoomPercent(value: number): number {
    const n = Number(value);
    if (!Number.isFinite(n)) return DEFAULT_ZOOM_PERCENT;
    return Math.min(Math.max(Math.round(n), ZOOM_OPTIONS[0]), ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1]);
}

function getViewStateKey(documentCacheId: string): string {
    return `${documentCacheId}${EXCEL_VIEW_STATE_SUFFIX}`;
}

function loadViewState(documentCacheId: string): ExcelViewState | null {
    if (!documentCacheId) return null;
    const key = getViewStateKey(documentCacheId);
    try {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        return JSON.parse(raw) as ExcelViewState;
    } catch {
        return null;
    }
}

function saveViewState(documentCacheId: string, view: ExcelViewState) {
    if (!documentCacheId) return;
    const key = getViewStateKey(documentCacheId);
    try {
        localStorage.setItem(key, JSON.stringify(view));
    } catch {
        // ignore quota / private mode errors
    }
}

function restoreViewState(spreadSheet: Spreadsheet, saved: ExcelViewState) {
    const sheets = spreadSheet.getData();
    if (!sheets.length) return;
    const sheetIndex = Math.min(Math.max(0, saved.sheetIndex), sheets.length - 1);
    const sheet = sheets[sheetIndex];
    const maxRi = Math.max(0, (sheet.rows?.len ?? 1) - 1);
    const maxCi = Math.max(0, (sheet.cols?.len ?? 1) - 1);
    const ri = Math.min(Math.max(0, saved.ri), maxRi);
    const ci = Math.min(Math.max(0, saved.ci), maxCi);
    spreadSheet.scrollToCell(ri, ci, sheetIndex);
}

function ExcelViewer() {
    const { message, modal } = App.useApp();
    const [loading, setLoading] = useState(true)
    const [dark, setDark] = useState(loadDarkMode)
    const [readOnly, setReadOnly] = useState(false)
    const [findPanel, setFindPanel] = useState<'find' | 'replace' | null>(null)
    const [loadError, setLoadError] = useState<string | null>(null)
    const [saveAsVisible, setSaveAsVisible] = useState(false)
    const [saveAsFormat, setSaveAsFormat] = useState('xlsx')
    const [activeSpreadsheet, setActiveSpreadsheet] = useState<Spreadsheet | null>(null)
    const [zoomPercent, setZoomPercent] = useState(DEFAULT_ZOOM_PERCENT)
    const extRef = useRef('')
    const documentCacheIdRef = useRef('')
    const readOnlyRef = useRef(false)
    const spreadSheetRef = useRef<Spreadsheet | null>(null)
    const zoomPercentRef = useRef(DEFAULT_ZOOM_PERCENT)
    const csvEncodingRef = useRef<'utf8' | 'gbk'>('utf8')
    const csvDelimiterRef = useRef(',')
    const initialFormattingRef = useRef('')
    const engineRef = useRef<SpreadsheetEngine>(ENGINE)
    const univerAdapterRef = useRef<UniverAdapter | null>(null)
    const univerViewUnbindRef = useRef<(() => void) | null>(null)
    const univerCtxRef = useRef<UniverSaveState | null>(null)
    const univerLossyWarnedRef = useRef(false)
    const darkRef = useRef(dark)

    useEffect(() => {
        document.body.classList.toggle('office-dark', dark)
        darkRef.current = dark
    }, [dark])

    const toggleDark = () => {
        setDark(prev => {
            const next = !prev
            applyDarkMode(next)
            return next
        })
    }

    useEffect(() => {
        spreadSheetRef.current?.reRender()
        univerAdapterRef.current?.setDarkMode(dark)
    }, [dark])

    const handleSaveAs = useCallback(() => {
        setSaveAsVisible(true);
    }, []);

    const applyZoomPercent = useCallback((value: number) => {
        const next = clampZoomPercent(value);
        zoomPercentRef.current = next;
        setZoomPercent(next);
        spreadSheetRef.current?.setZoom(next / 100);
        univerAdapterRef.current?.setZoom(next / 100);
    }, []);

    const stepZoom = useCallback((direction: -1 | 1) => {
        const current = zoomPercentRef.current;
        const next = direction < 0
            ? [...ZOOM_OPTIONS].reverse().find(value => value < current) ?? ZOOM_OPTIONS[0]
            : ZOOM_OPTIONS.find(value => value > current) ?? ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1];
        applyZoomPercent(next);
    }, [applyZoomPercent]);

    const handleSave = useCallback(async () => {
        const spreadSheet = spreadSheetRef.current;
        if (!spreadSheet) return;
        if (readOnlyRef.current) {
            await handleSaveAs();
            return;
        }

        const ext = extRef.current.replace(/^\./, '').toLowerCase();
        const sheets = spreadSheet.getData();
        const csvEncoding = csvEncodingRef.current;
        const csvDelimiter = csvDelimiterRef.current;

        if (ext !== 'xlsx' && ext !== 'xlsm' && hasFormattingChanged(initialFormattingRef.current, sheets)) {
            await new Promise<void>((resolve) => {
                const dialog = modal.confirm({
                    title: t('viewer.formatCannotPreserveTitle'),
                    content: t('viewer.formatCannotPreserveContent', ext.toUpperCase()),
                    okText: t('viewer.saveAsXlsx'),
                    cancelText: t('button.cancel'),
                    centered: true,
                    getContainer: () => document.body,
                    onOk: async () => {
                        try {
                            await export_xlsx(spreadSheet, 'xlsx', csvEncoding, { saveAs: true }, csvDelimiter);
                        } catch (error) {
                            console.error(`Failed to save Excel file: ${(error as Error).message}`);
                            throw error;
                        }
                    },
                    onCancel: () => { },
                    footer: () => (
                        <>
                            <Button
                                style={{ padding: '3px 12px', height: 'auto' }}
                                onClick={() => dialog.destroy()}
                            >
                                {t('button.cancel')}
                            </Button>
                            <Button
                                style={{ padding: '3px 12px', height: 'auto' }}
                                onClick={() => {
                                    void (async () => {
                                        dialog.destroy();
                                        try {
                                            await export_xlsx(spreadSheet, extRef.current, csvEncoding, undefined, csvDelimiter);
                                        } catch (error) {
                                            console.error(`Failed to save Excel file: ${(error as Error).message}`);
                                        }
                                    })();
                                }}
                            >
                                {t('viewer.saveAsOriginal')}
                            </Button>
                            <Button
                                type="primary"
                                style={{ padding: '3px 12px', height: 'auto' }}
                                onClick={() => {
                                    void (async () => {
                                        try {
                                            dialog.destroy();
                                            await export_xlsx(spreadSheet, 'xlsx', csvEncoding, { saveAs: true }, csvDelimiter);
                                        } catch (error) {
                                            console.error(`Failed to save Excel file: ${(error as Error).message}`);
                                        }
                                    })();
                                }}
                            >
                                {t('viewer.saveAsXlsx')}
                            </Button>
                        </>
                    ),
                    afterClose: () => resolve(),
                });
            });
            return;
        }

        try {
            await export_xlsx(spreadSheet, extRef.current, csvEncoding, undefined, csvDelimiter);
            spreadSheet.setSaveEnabled(false);
        } catch (error) {
            console.error(`Failed to save Excel file: ${(error as Error).message}`);
        }
    }, [modal, handleSaveAs]);

    const univerSave = useCallback(async (options?: { saveAs?: boolean; saveAsExt?: string }) => {
        const adapter = univerAdapterRef.current;
        const ctx = univerCtxRef.current;
        if (!adapter || !ctx) return;

        const { saveUniverWorkbook, hasFormattingChangedUniver } = await import('./univer/export');
        const current = adapter.getWorkbookDataCopy() as UniverSaveState['baseline'] | null;
        if (!current) return;

        const targetExt = (options?.saveAs ? options.saveAsExt ?? 'xlsx' : extRef.current)
            .replace(/^\./, '').toLowerCase();

        // 图表/透视表/宏安全网：现状是这些文件根本打不开；现在可打开可编辑，
        // 但 ExcelJS 无法承载这些部件，保存（含另存 xlsx）会丢失，明示一次
        const lossy = ctx.loadResult.lossy;
        const hasLossy = !!(lossy && (lossy.charts || lossy.pivotTables || lossy.vba));
        if (hasLossy && (targetExt === 'xlsx' || targetExt === 'xlsm') && !univerLossyWarnedRef.current) {
            const proceed = await new Promise<boolean>((resolve) => {
                modal.confirm({
                    title: isZhLang() ? '部分内容无法保留' : 'Some content cannot be preserved',
                    content: isZhLang()
                        ? '此文件包含图表、透视表或宏。当前引擎保存后这些内容将丢失（原文件在保存前不受影响）。'
                        : 'This file contains charts, pivot tables or macros. Saving with the current engine will drop them (the original file is untouched until you save).',
                    okText: isZhLang() ? '仍要保存' : 'Save anyway',
                    cancelText: t('button.cancel'),
                    centered: true,
                    getContainer: () => document.body,
                    onOk: () => resolve(true),
                    onCancel: () => resolve(false),
                });
            });
            if (!proceed) return;
            univerLossyWarnedRef.current = true;
        }

        // 非 xlsx 目标：格式变更过则确认（与 legacy 行为一致）
        if (targetExt !== 'xlsx' && targetExt !== 'xlsm'
            && hasFormattingChangedUniver(ctx.baseline, current)) {
            const choice = await new Promise<'xlsx' | 'original' | 'cancel'>((resolve) => {
                const dialog = modal.confirm({
                    title: t('viewer.formatCannotPreserveTitle'),
                    content: t('viewer.formatCannotPreserveContent', targetExt.toUpperCase()),
                    centered: true,
                    getContainer: () => document.body,
                    footer: () => (
                        <>
                            <Button style={{ padding: '3px 12px', height: 'auto' }} onClick={() => { dialog.destroy(); resolve('cancel'); }}>
                                {t('button.cancel')}
                            </Button>
                            <Button style={{ padding: '3px 12px', height: 'auto' }} onClick={() => { dialog.destroy(); resolve('original'); }}>
                                {t('viewer.saveAsOriginal')}
                            </Button>
                            <Button type="primary" style={{ padding: '3px 12px', height: 'auto' }} onClick={() => { dialog.destroy(); resolve('xlsx'); }}>
                                {t('viewer.saveAsXlsx')}
                            </Button>
                        </>
                    ),
                });
            });
            if (choice === 'cancel') return;
            if (choice === 'xlsx') {
                await univerSaveInner(current, ctx, { saveAs: true, saveAsExt: 'xlsx' });
                return;
            }
        }

        await univerSaveInner(current, ctx, options);

        async function univerSaveInner(
            cur: UniverSaveState['baseline'],
            saveCtx: UniverSaveState,
            saveOptions?: { saveAs?: boolean; saveAsExt?: string },
        ) {
            try {
                const { newSheetIdMap } = await saveUniverWorkbook(cur, {
                    ext: extRef.current.replace(/^\./, '').toLowerCase() || 'xlsx',
                    loadResult: saveCtx.loadResult,
                    baseline: saveCtx.baseline,
                    structuralSheetIds: saveCtx.session.structuralSheetIds,
                    csvEncoding: csvEncodingRef.current,
                    csvDelimiter: csvDelimiterRef.current,
                }, saveOptions);
                // 增量导出会把原始 workbook 推进到 current 状态：同步基线与会话
                if (newSheetIdMap) saveCtx.loadResult.sheetIdMap = newSheetIdMap;
                saveCtx.baseline = cur;
                saveCtx.session.reset();
            } catch (error) {
                console.error(`Failed to save Excel file: ${(error as Error).message}`, error);
                message.error({
                    duration: 3,
                    content: isZhLang() ? '保存失败，详见开发者工具控制台' : 'Save failed, see developer console',
                });
            }
        }
    }, [modal, message]);

    const confirmSaveAs = useCallback(async (fmt: string) => {
        setSaveAsVisible(false);
        if (engineRef.current === 'univer') {
            await univerSave({ saveAs: true, saveAsExt: fmt });
            return;
        }
        const spreadSheet = spreadSheetRef.current;
        if (!spreadSheet) return;
        try {
            await exportSaveAs(spreadSheet, fmt, csvEncodingRef.current, csvDelimiterRef.current);
            if (!readOnlyRef.current) {
                spreadSheet.setSaveEnabled(false);
            }
        } catch (error) {
            console.error(`Failed to save Excel file: ${(error as Error).message}`);
        }
    }, [univerSave]);

    useEffect(() => {
        const onKeyDown = (e: KeyboardEvent) => {
            if (engineRef.current === 'univer') {
                // Ctrl+F/H 交给 Univer 自带查找替换
                if ((e.ctrlKey || e.metaKey) && e.code === 'KeyS') {
                    e.preventDefault();
                    if (readOnlyRef.current) {
                        void handleSaveAs();
                    } else {
                        void univerSave();
                    }
                }
                return;
            }
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyS') {
                e.preventDefault();
                if (readOnlyRef.current) {
                    void handleSaveAs();
                } else {
                    void handleSave();
                }
                return;
            }
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyF') {
                e.preventDefault();
                setFindPanel('find');
                return;
            }
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyH') {
                e.preventDefault();
                setFindPanel(readOnlyRef.current ? 'find' : 'replace');
            }
        };
        window.addEventListener('keydown', onKeyDown);
        return () => window.removeEventListener('keydown', onKeyDown);
    }, [handleSave, handleSaveAs, message, univerSave]);

    useEffect(() => {
        const container = document.getElementById('container');

        const initUniver = async (buffer: ArrayBuffer, payload: any) => {
            const fileReadOnly = payload.readOnly === true;
            univerViewUnbindRef.current?.();
            univerViewUnbindRef.current = null;
            univerCtxRef.current?.session.stop();
            univerCtxRef.current = null;
            univerAdapterRef.current?.dispose();
            univerAdapterRef.current = null;
            univerLossyWarnedRef.current = false;
            container.innerHTML = '';
            container.style.height = '100vh';

            if (payload.ext?.match(/csv/i)) {
                csvEncodingRef.current = detectCsvEncoding(buffer);
            }

            const [{ UniverAdapter }, { loadForUniver }] = await Promise.all([
                import('./univer/adapter'),
                import('./univer/loader'),
            ]);
            const result = await loadForUniver(buffer, payload.ext ?? 'xlsx', payload.fileName ?? 'Workbook');
            if (result.csvDelimiter) {
                csvDelimiterRef.current = result.csvDelimiter;
            }
            const adapter = UniverAdapter.create(container, {
                darkMode: darkRef.current,
                language: getConfigs()?.language,
                readOnly: fileReadOnly,
            });
            univerAdapterRef.current = adapter;
            setLoading(false);
            await adapter.loadWorkbook(result, { readOnly: fileReadOnly });
            if (zoomPercentRef.current !== DEFAULT_ZOOM_PERCENT) {
                adapter.setZoom(zoomPercentRef.current / 100);
            }
            const savedView = loadViewState(documentCacheIdRef.current);
            if (savedView) {
                adapter.restoreViewState(savedView);
            }
            univerViewUnbindRef.current = adapter.onViewStateChange(() => {
                const vs = adapter.getViewState();
                if (vs) saveViewState(documentCacheIdRef.current, vs);
            });

            // 基线必须在超链接应用与视图恢复之后取（此后的 mutation 才算用户编辑）
            const baseline = adapter.getWorkbookDataCopy() as UniverSaveState['baseline'] | null;
            if (baseline) {
                const session = adapter.startEditSession(() => {
                    if (!fileReadOnly) handler.emit('change');
                });
                univerCtxRef.current = { loadResult: result, baseline, session };
            }
        };

        const initSpreadsheet = async (buffer: ArrayBuffer, payload: any) => {
            const fileReadOnly = payload.readOnly === true;
            if (payload.ext?.match(/csv/i)) {
                csvEncodingRef.current = detectCsvEncoding(buffer);
            }
            const { sheets, maxLength, maxCols, csvDelimiter } = await loadSheets(buffer, payload.ext);
            if (csvDelimiter) {
                csvDelimiterRef.current = csvDelimiter;
            }
            const viewRowLen = Math.max(maxLength ?? 0, MIN_VIEW_ROWS);
            const viewColLen = Math.max(maxCols ?? 0, MIN_VIEW_COLS);
            container.innerHTML = '';
            const spreadSheet = new Spreadsheet(container, {
                mode: fileReadOnly ? 'read' : 'edit',
                showToolbar: true,
                zoom: zoomPercentRef.current / 100,
                row: { len: viewRowLen, height: 30 },
                col: { len: viewColLen },
                view: { height: () => window.innerHeight - 2 },
            });
            spreadSheetRef.current = spreadSheet;
            setActiveSpreadsheet(spreadSheet);
            setLoading(false);
            spreadSheet.loadData(sheets);
            if (!fileReadOnly) {
                spreadSheet.on('save', () => void handleSave());
            }
            spreadSheet.on('save-as', () => { void handleSaveAs(); });
            spreadSheet.on('find', () => { setFindPanel('find'); });
            const persistView = () => {
                saveViewState(documentCacheIdRef.current, spreadSheet.getSelection());
            };
            spreadSheet.on('cell-selected', () => { persistView(); });
            spreadSheet.onSheetChange(() => { persistView(); });
            spreadSheet.onOpenLink((linkPayload) => {
                const parsed = parseSpreadsheetLink(linkPayload.link);
                if (parsed.type === 'internal') {
                    spreadSheet.followHyperlink(linkPayload);
                } else {
                    handler.emit('openExternal', parsed.url);
                }
            });
            spreadSheet.onProtectedCellDblClick(() => {
                message.info({ duration: 2, content: t('viewer.protectedCell'), className: 'excel-protected-cell-message' });
            });
            spreadSheet.onValidationError((errMessage) => {
                message.warning({ duration: 2, content: errMessage, className: 'excel-validation-error-message' });
            });
            spreadSheet.on('change', () => {
                if (!fileReadOnly) {
                    spreadSheet.setSaveEnabled(true);
                    handler.emit('change');
                }
            });
            const savedView = loadViewState(documentCacheIdRef.current);
            if (savedView) {
                requestAnimationFrame(() => {
                    requestAnimationFrame(() => { restoreViewState(spreadSheet, savedView); });
                });
            }
            initialFormattingRef.current = buildFormattingSnapshot(spreadSheet.getData());
        };

        handler.on("open", (payload) => {
            extRef.current = payload.ext ?? '';
            documentCacheIdRef.current = payload.documentCacheId ?? '';
            const fileReadOnly = payload.readOnly === true;
            readOnlyRef.current = fileReadOnly;
            setReadOnly(fileReadOnly);
            loadOfficeBuffer(payload).then(async (buffer) => {
                try {
                    if (engineRef.current === 'univer') {
                        await initUniver(buffer, payload);
                    } else {
                        await initSpreadsheet(buffer, payload);
                    }
                } catch (e) {
                    const msg = (e as Error).message || String(e);
                    console.error(`Failed to load Excel file: ${msg}`, e);
                    setLoadError(msg);
                    setLoading(false);
                }
            }).catch(error => {
                const msg = (error as Error).message || String(error);
                console.error(`Failed to load Excel file: ${msg}`, error);
                setLoadError(msg);
                setLoading(false);
            });
        }).on("saveDone", () => {
        }).emit("init")

        let themeTimer: ReturnType<typeof setTimeout>;
        const themeObserver = new MutationObserver(() => {
            clearTimeout(themeTimer);
            themeTimer = setTimeout(() => spreadSheetRef.current?.reRender(), 120);
        });
        themeObserver.observe(document.head, { childList: true, subtree: true });

        return () => {
            spreadSheetRef.current = null;
            setActiveSpreadsheet(null);
            univerViewUnbindRef.current?.();
            univerViewUnbindRef.current = null;
            univerCtxRef.current?.session.stop();
            univerCtxRef.current = null;
            univerAdapterRef.current?.dispose();
            univerAdapterRef.current = null;
            themeObserver.disconnect();
            clearTimeout(themeTimer);
        };
    }, [message, handleSave, handleSaveAs])

    return (
        <div className='excel-viewer'>
            <Spin spinning={loading} fullscreen={true} />
            {loadError && !loading && (
                <div className="excel-load-error">
                    <div className="excel-load-error-panel">
                        <svg className="excel-load-error-icon" width="44" height="44" viewBox="0 0 44 44" fill="none" aria-hidden>
                            <circle cx="22" cy="22" r="20" stroke="currentColor" strokeWidth="1.8" />
                            <path d="M22 13v12" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" />
                            <circle cx="22" cy="31" r="1.8" fill="currentColor" />
                        </svg>
                        <h2 className="excel-load-error-title">Failed to open file</h2>
                        <span className="excel-load-error-message">{loadError}</span>
                    </div>
                </div>
            )}
            {readOnly && !loading && !loadError && (
                <div className="excel-readonly-banner">
                    {t('viewer.readonlyBanner')}
                </div>
            )}
            {ENGINE !== 'univer' && findPanel && !loading && !loadError && (
                <FindReplacePanel
                    spreadSheet={activeSpreadsheet}
                    mode={findPanel}
                    onClose={() => setFindPanel(null)}
                    readOnly={readOnly}
                    onChanged={() => {
                        if (!readOnlyRef.current) {
                            spreadSheetRef.current?.setSaveEnabled(true);
                        }
                    }}
                />
            )}
            <Modal
                open={saveAsVisible}
                title={t('button.saveAs')}
                onCancel={() => setSaveAsVisible(false)}
                footer={[
                    <Button key="cancel" onClick={() => setSaveAsVisible(false)} style={{ padding: '3px 12px', height: 'auto' }}>
                        {t('button.cancel')}
                    </Button>,
                    <Button key="ok" type="primary" onClick={() => void confirmSaveAs(saveAsFormat)} style={{ padding: '3px 12px', height: 'auto' }}>
                        {t('button.save')}
                    </Button>,
                ]}
                getContainer={() => document.body}
                centered
                width={360}
            >
                <div style={{ padding: '8px 0 16px' }}>
                    <div style={{ marginBottom: 12, opacity: 0.65, fontSize: 12 }}>
                        {t('viewer.chooseExportFormat')}
                    </div>
                    <Radio.Group
                        value={saveAsFormat}
                        onChange={e => setSaveAsFormat(e.target.value as string)}
                        style={{ display: 'flex', flexDirection: 'column', gap: 10 }}
                    >
                        {[
                            { value: 'xlsx', label: t('viewer.exportXlsxLabel'), desc: t('viewer.exportXlsxDesc') },
                            { value: 'csv', label: t('viewer.exportCsvLabel'), desc: t('viewer.exportCsvDesc') },
                            { value: 'xls', label: t('viewer.exportXlsLabel'), desc: t('viewer.exportXlsDesc') },
                            { value: 'ods', label: t('viewer.exportOdsLabel'), desc: t('viewer.exportOdsDesc') },
                        ].map(f => (
                            <Radio key={f.value} value={f.value} style={{ alignItems: 'flex-start' }}>
                                <span style={{ fontWeight: 500 }}>{f.label}</span>
                                <span style={{ display: 'block', fontSize: 11, opacity: 0.55, marginTop: 1 }}>{f.desc}</span>
                            </Radio>
                        ))}
                    </Radio.Group>
                </div>
            </Modal>
            <div id='container'></div>
            <div className="excel-footer-actions">
                <div className="excel-zoom-control" aria-label={t('viewer.zoom')}>
                    <button
                        type="button"
                        className="excel-zoom-button"
                        title={t('viewer.zoomOut')}
                        onClick={() => stepZoom(-1)}
                        disabled={zoomPercent <= ZOOM_OPTIONS[0]}
                    >
                        <MinusOutlined />
                    </button>
                    <select
                        className="excel-zoom-select"
                        value={zoomPercent}
                        title={t('viewer.zoom')}
                        onChange={event => applyZoomPercent(Number(event.target.value))}
                    >
                        {ZOOM_OPTIONS.map(value => (
                            <option key={value} value={value}>{value}%</option>
                        ))}
                    </select>
                    <button
                        type="button"
                        className="excel-zoom-button"
                        title={t('viewer.zoomIn')}
                        onClick={() => stepZoom(1)}
                        disabled={zoomPercent >= ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1]}
                    >
                        <PlusOutlined />
                    </button>
                </div>
                <button
                    type="button"
                    className="dark-mode-toggle"
                    title={dark ? t('viewer.switchToLightMode') : t('viewer.switchToDarkMode')}
                    onClick={toggleDark}
                >
                    {dark ? <SunOutlined /> : <MoonOutlined />}
                </button>
            </div>
        </div>
    )
}

export default function Excel() {
    return (
        <App className="excel-app" message={{ top: 16 }}>
            <ExcelViewer />
        </App>
    );
}
