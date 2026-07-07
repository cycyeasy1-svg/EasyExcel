import { ExportOutlined, MinusOutlined, MoonOutlined, PlusOutlined, SaveOutlined, SunOutlined } from "@ant-design/icons";
import { App, Button, Modal, Radio, Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler, loadDarkMode, applyDarkMode } from "../../util/vscode.ts";
import { loadOfficeBuffer } from "../../util/loadOfficeContent.ts";
import './Excel.less';
import { detectCsvEncoding } from "./csvEncoding.ts";
import { initExcelLocale, t } from './excel_i18n';
import { getConfigs } from '../../util/vscodeConfig.ts';
import type { UniverAdapter, UniverEditSession } from './univer/adapter';
import type { UniverLoadResult } from './univer/loader';
import type { IWorkbookData } from '@univerjs/core';

initExcelLocale();

type ExcelViewState = { ri: number; ci: number; sheetIndex: number };

const EXCEL_VIEW_STATE_SUFFIX = '-excel-view';
const DEFAULT_ZOOM_PERCENT = 100;
const ZOOM_OPTIONS = [50, 75, 90, 100, 125, 150, 200];

/** Univer 编辑会话的保存上下文（initUniver 组装，保存后就地更新） */
interface UniverSaveState {
    loadResult: UniverLoadResult;
    baseline: IWorkbookData;
    session: UniverEditSession;
}

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
    try {
        const raw = localStorage.getItem(getViewStateKey(documentCacheId));
        if (!raw) return null;
        return JSON.parse(raw) as ExcelViewState;
    } catch {
        return null;
    }
}

function saveViewState(documentCacheId: string, view: ExcelViewState) {
    if (!documentCacheId) return;
    try {
        localStorage.setItem(getViewStateKey(documentCacheId), JSON.stringify(view));
    } catch {
        // ignore quota / private mode errors
    }
}

function ExcelViewer() {
    const { message, modal } = App.useApp();
    const [loading, setLoading] = useState(true)
    const [dark, setDark] = useState(loadDarkMode)
    const [readOnly, setReadOnly] = useState(false)
    const [loadError, setLoadError] = useState<string | null>(null)
    const [saveAsVisible, setSaveAsVisible] = useState(false)
    const [saveAsFormat, setSaveAsFormat] = useState('xlsx')
    const [zoomPercent, setZoomPercent] = useState(DEFAULT_ZOOM_PERCENT)
    const extRef = useRef('')
    const documentCacheIdRef = useRef('')
    const readOnlyRef = useRef(false)
    const zoomPercentRef = useRef(DEFAULT_ZOOM_PERCENT)
    const csvEncodingRef = useRef<'utf8' | 'gbk'>('utf8')
    const csvDelimiterRef = useRef(',')
    const univerAdapterRef = useRef<UniverAdapter | null>(null)
    const univerViewUnbindRef = useRef<(() => void) | null>(null)
    const univerRestrictUnbindRef = useRef<(() => void) | null>(null)
    const univerCtxRef = useRef<UniverSaveState | null>(null)
    const univerLossyWarnedRef = useRef(false)
    const darkRef = useRef(dark)

    useEffect(() => {
        document.body.classList.toggle('office-dark', dark)
        darkRef.current = dark
        univerAdapterRef.current?.setDarkMode(dark)
    }, [dark])

    const toggleDark = () => {
        setDark(prev => {
            const next = !prev
            applyDarkMode(next)
            return next
        })
    }

    const handleSaveAs = useCallback(() => {
        setSaveAsVisible(true);
    }, []);

    const applyZoomPercent = useCallback((value: number) => {
        const next = clampZoomPercent(value);
        zoomPercentRef.current = next;
        setZoomPercent(next);
        univerAdapterRef.current?.setZoom(next / 100);
    }, []);

    const stepZoom = useCallback((direction: -1 | 1) => {
        const current = zoomPercentRef.current;
        const next = direction < 0
            ? [...ZOOM_OPTIONS].reverse().find(value => value < current) ?? ZOOM_OPTIONS[0]
            : ZOOM_OPTIONS.find(value => value > current) ?? ZOOM_OPTIONS[ZOOM_OPTIONS.length - 1];
        applyZoomPercent(next);
    }, [applyZoomPercent]);

    const univerSave = useCallback(async (options?: { saveAs?: boolean; saveAsExt?: string }) => {
        const adapter = univerAdapterRef.current;
        const ctx = univerCtxRef.current;
        if (!adapter || !ctx) return;

        const { saveUniverWorkbook, hasFormattingChangedUniver, XmlPatchBlockedError } = await import('./univer/export');
        const current = adapter.getWorkbookDataCopy() as IWorkbookData | null;
        if (!current) return;

        const targetExt = (options?.saveAs ? options.saveAsExt ?? 'xlsx' : extRef.current)
            .replace(/^\./, '').toLowerCase();

        // 图表/透视表/宏安全网：ExcelJS 重建路径（另存为副本）会丢失这些部件。
        // M5 起保存回原文件走 XML 补丁，部件物理透传，无需警告。
        const usesXmlPatch = !options?.saveAs && !!ctx.loadResult.originalBuffer
            && (targetExt === 'xlsx' || targetExt === 'xlsm');
        const lossy = ctx.loadResult.lossy;
        const hasLossy = !!(lossy && (lossy.charts || lossy.pivotTables || lossy.vba));
        if (hasLossy && !usesXmlPatch && (targetExt === 'xlsx' || targetExt === 'xlsm') && !univerLossyWarnedRef.current) {
            const proceed = await new Promise<boolean>((resolve) => {
                modal.confirm({
                    title: t('viewer.lossyTitle'),
                    content: t('viewer.lossyContent'),
                    okText: t('viewer.lossySaveAnyway'),
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

        // 非 xlsx 目标：格式变更过则确认
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
            cur: IWorkbookData,
            saveCtx: UniverSaveState,
            saveOptions?: { saveAs?: boolean; saveAsExt?: string },
        ) {
            try {
                const { newSheetIdMap, richTextDowngraded } = await saveUniverWorkbook(cur, {
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
                if (richTextDowngraded) {
                    message.info({ duration: 4, content: t('viewer.richTextDowngraded') });
                }
            } catch (error) {
                if (error instanceof XmlPatchBlockedError) {
                    // 命令拦截失守（如粘贴带入合并）：拒绝保存并告知出路
                    message.warning({ duration: 5, content: t('viewer.patchBlockedSave') });
                    return;
                }
                console.error(`Failed to save Excel file: ${(error as Error).message}`, error);
                message.error({ duration: 3, content: t('viewer.saveFailed') });
            }
        }
    }, [modal, message]);

    const confirmSaveAs = useCallback(async (fmt: string) => {
        setSaveAsVisible(false);
        await univerSave({ saveAs: true, saveAsExt: fmt });
    }, [univerSave]);

    useEffect(() => {
        const onKeyDown = (e: KeyboardEvent) => {
            // Ctrl+F/H 交给 Univer 自带查找替换
            if ((e.ctrlKey || e.metaKey) && e.code === 'KeyS') {
                e.preventDefault();
                if (readOnlyRef.current) {
                    void handleSaveAs();
                } else {
                    void univerSave();
                }
            }
        };
        window.addEventListener('keydown', onKeyDown);
        return () => window.removeEventListener('keydown', onKeyDown);
    }, [handleSaveAs, univerSave]);

    useEffect(() => {
        const container = document.getElementById('container');

        const initUniver = async (buffer: ArrayBuffer, payload: any) => {
            const fileReadOnly = payload.readOnly === true;
            univerViewUnbindRef.current?.();
            univerViewUnbindRef.current = null;
            univerRestrictUnbindRef.current?.();
            univerRestrictUnbindRef.current = null;
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
            await adapter.loadWorkbook(result, {
                readOnly: fileReadOnly,
                onOpenExternal: (url) => handler.emit('openExternal', url),
            });
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

            // 基线必须在超链接/特性应用与视图恢复之后取（此后的 mutation 才算用户编辑）
            const baseline = adapter.getWorkbookDataCopy() as IWorkbookData | null;
            if (baseline) {
                const session = adapter.startEditSession(() => {
                    if (!fileReadOnly) handler.emit('change');
                });
                univerCtxRef.current = { loadResult: result, baseline, session };
            }

            // M5：xlsx/xlsm 原文件的保存走 XML 补丁 → 无法无损写回的编辑
            // （结构/合并/行高列宽/冻结/DV/CF）在命令层禁用，toast 引导去 Excel
            const patchExt = /^\.?(xlsx|xlsm)$/i.test(payload.ext ?? '');
            if (!fileReadOnly && patchExt && result.originalBuffer) {
                let lastToastAt = 0;
                univerRestrictUnbindRef.current = adapter.restrictLossyEdits(() => {
                    const now = Date.now();
                    if (now - lastToastAt < 2000) return;
                    lastToastAt = now;
                    message.warning({ duration: 4, content: t('viewer.structuralBlocked') });
                });
            }
        };

        handler.on("open", (payload) => {
            extRef.current = payload.ext ?? '';
            documentCacheIdRef.current = payload.documentCacheId ?? '';
            const fileReadOnly = payload.readOnly === true;
            readOnlyRef.current = fileReadOnly;
            setReadOnly(fileReadOnly);
            loadOfficeBuffer(payload).then(async (buffer) => {
                try {
                    await initUniver(buffer, payload);
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
            message.success({ duration: 2, content: t('viewer.saved') });
        }).emit("init")

        return () => {
            univerViewUnbindRef.current?.();
            univerViewUnbindRef.current = null;
            univerRestrictUnbindRef.current?.();
            univerRestrictUnbindRef.current = null;
            univerCtxRef.current?.session.stop();
            univerCtxRef.current = null;
            univerAdapterRef.current?.dispose();
            univerAdapterRef.current = null;
        };
    }, [message])

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
                <div className="excel-zoom-control" aria-label={t('button.save')}>
                    {!readOnly && (
                        <button
                            type="button"
                            className="excel-zoom-button"
                            title={`${t('button.save')} (Ctrl+S)`}
                            onClick={() => void univerSave()}
                        >
                            <SaveOutlined />
                        </button>
                    )}
                    <button
                        type="button"
                        className="excel-zoom-button"
                        title={t('button.saveAs')}
                        onClick={handleSaveAs}
                    >
                        <ExportOutlined />
                    </button>
                </div>
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
