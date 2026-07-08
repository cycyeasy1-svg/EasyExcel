import { ExportOutlined, MoonOutlined, SaveOutlined, SunOutlined, TranslationOutlined } from "@ant-design/icons";
import { App, Button, Dropdown, Modal, Radio, Spin } from "antd";
import { useCallback, useEffect, useRef, useState } from "react";
import { handler, loadDarkMode, applyDarkMode, loadLanguage, applyLanguage } from "../../util/vscode.ts";
import { loadWorkbookBuffer } from "../../util/loadWorkbookContent.ts";
import './Excel.less';
import { detectCsvEncoding } from "./csvEncoding.ts";
import { initExcelLocale, setExcelLocale, t } from './excel_i18n';
import type { UniverAdapter, UniverEditSession } from './univer/adapter';
import type { UniverLoadResult } from './univer/loader';
import type { IWorkbookData } from '@univerjs/core';

initExcelLocale();

type ExcelViewState = { ri: number; ci: number; sheetIndex: number };

const EXCEL_VIEW_STATE_SUFFIX = '-excel-view';

/** 界面语言选项（打包了 zh-CN / en-US / ja-JP 三套 Univer locale） */
const LANGUAGE_OPTIONS = [
    { key: 'zh-cn', label: '简体中文' },
    { key: 'en', label: 'English' },
    { key: 'ja', label: '日本語' },
];

/** Univer 编辑会话的保存上下文（initUniver 组装，保存后就地更新） */
interface UniverSaveState {
    loadResult: UniverLoadResult;
    baseline: IWorkbookData;
    session: UniverEditSession;
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
    const [language, setLanguage] = useState(loadLanguage)
    const extRef = useRef('')
    const documentCacheIdRef = useRef('')
    const readOnlyRef = useRef(false)
    const dirtyRef = useRef(false)
    const csvEncodingRef = useRef<'utf8' | 'gbk'>('utf8')
    const csvDelimiterRef = useRef(',')
    const univerAdapterRef = useRef<UniverAdapter | null>(null)
    const univerViewUnbindRef = useRef<(() => void) | null>(null)
    const univerRestrictUnbindRef = useRef<(() => void) | null>(null)
    const univerCtxRef = useRef<UniverSaveState | null>(null)
    const univerLossyWarnedRef = useRef(false)
    const darkRef = useRef(dark)

    useEffect(() => {
        document.body.classList.toggle('easyexcel-dark', dark)
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

    // 切换语言：外壳文案即时生效；Univer locale 只能在创建时指定，
    // 让宿主重发文件内容（init → open）以新语言重建实例
    const switchLanguage = useCallback(async (next: string) => {
        if (next === loadLanguage()) return;
        if (dirtyRef.current) {
            const proceed = await new Promise<boolean>((resolve) => {
                modal.confirm({
                    title: t('viewer.langSwitchTitle'),
                    content: t('viewer.langSwitchContent'),
                    okText: t('viewer.langSwitchAnyway'),
                    cancelText: t('button.cancel'),
                    centered: true,
                    getContainer: () => document.body,
                    onOk: () => resolve(true),
                    onCancel: () => resolve(false),
                });
            });
            if (!proceed) return;
        }
        applyLanguage(next);
        setExcelLocale(next);
        setLanguage(next);
        setLoading(true);
        handler.emit('init');
    }, [modal]);

    // 返回是否已把字节派发给宿主写盘（true=保存流程进入写盘阶段；
    // false=用户取消/无法保存）。宿主 requestSave 桥据此决定原生 dirty 去留。
    const univerSave = useCallback(async (options?: { saveAs?: boolean; saveAsExt?: string }): Promise<boolean> => {
        const adapter = univerAdapterRef.current;
        const ctx = univerCtxRef.current;
        if (!adapter || !ctx) return false;

        const { saveUniverWorkbook, hasFormattingChangedUniver, XmlPatchBlockedError } = await import('./univer/export');
        const current = adapter.getWorkbookDataCopy() as IWorkbookData | null;
        if (!current) return false;

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
            if (!proceed) return false;
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
            if (choice === 'cancel') return false;
            if (choice === 'xlsx') {
                return await univerSaveInner(current, ctx, { saveAs: true, saveAsExt: 'xlsx' });
            }
        }

        return await univerSaveInner(current, ctx, options);

        async function univerSaveInner(
            cur: IWorkbookData,
            saveCtx: UniverSaveState,
            saveOptions?: { saveAs?: boolean; saveAsExt?: string },
        ): Promise<boolean> {
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
                dirtyRef.current = false;
                if (richTextDowngraded) {
                    message.info({ duration: 4, content: t('viewer.richTextDowngraded') });
                }
                return true;
            } catch (error) {
                if (error instanceof XmlPatchBlockedError) {
                    // 命令拦截失守（如粘贴带入合并）：拒绝保存并告知出路
                    message.warning({ duration: 5, content: t('viewer.patchBlockedSave') });
                    return false;
                }
                console.error(`Failed to save Excel file: ${(error as Error).message}`, error);
                message.error({ duration: 3, content: t('viewer.saveFailed') });
                return false;
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
                    // 交由 VSCode 发起保存 → saveCustomDocument → requestSave，
                    // 写盘后原生 dirty 状态才会正确清除
                    handler.emit('hostSave');
                }
            }
        };
        window.addEventListener('keydown', onKeyDown);
        return () => window.removeEventListener('keydown', onKeyDown);
    }, [handleSaveAs]);

    // 宿主（saveCustomDocument）请求保存：跑既有保存流程，未真正写盘则回报取消
    useEffect(() => {
        handler.on('requestSave', () => {
            void (async () => {
                const dispatched = readOnlyRef.current ? false : await univerSave();
                if (!dispatched) handler.emit('saveSettled', { cancelled: true });
            })();
        });
    }, [univerSave]);

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
            dirtyRef.current = false;
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
                language: loadLanguage(),
                readOnly: fileReadOnly,
            });
            univerAdapterRef.current = adapter;
            setLoading(false);
            await adapter.loadWorkbook(result, {
                readOnly: fileReadOnly,
                onOpenExternal: (url) => handler.emit('openExternal', url),
            });
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
                // 仅当相对基线真正产生内容/格式差异时才标记「已变更」。
                // 仅点选、进入/退出编辑但未改动内容不会触发星号。已 dirty 后
                // 不再重复比对（省开销；保存/重载后基线与 dirtyRef 会一并复位）。
                const session = adapter.startEditSession(() => {
                    if (dirtyRef.current || fileReadOnly) return;
                    void (async () => {
                        // 对比「最近一次保存后的基线」（保存会把 ctx.baseline 推进到
                        // 已存状态并 reset 结构日志），而非初始加载基线
                        const ctx = univerCtxRef.current;
                        if (!ctx || dirtyRef.current) return;
                        const cur = adapter.getWorkbookDataCopy() as IWorkbookData | null;
                        if (!cur || dirtyRef.current) return;
                        const { diffWorkbook } = await import('./univer/diff');
                        if (dirtyRef.current) return;
                        if (diffWorkbook(ctx.baseline, cur, ctx.session.structuralSheetIds).isEmpty) return;
                        dirtyRef.current = true;
                        handler.emit('change');
                    })();
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
            loadWorkbookBuffer(payload).then(async (buffer) => {
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
            <div className="excel-topbar-actions">
                <div className="excel-action-group">
                    {!readOnly && (
                        <button
                            type="button"
                            className="excel-action-button"
                            title={`${t('button.save')} (Ctrl+S)`}
                            onClick={() => handler.emit('hostSave')}
                        >
                            <SaveOutlined />
                        </button>
                    )}
                    <button
                        type="button"
                        className="excel-action-button"
                        title={t('button.saveAs')}
                        onClick={handleSaveAs}
                    >
                        <ExportOutlined />
                    </button>
                </div>
                <Dropdown
                    trigger={['click']}
                    menu={{
                        items: LANGUAGE_OPTIONS,
                        selectable: true,
                        selectedKeys: [language],
                        onClick: ({ key }) => void switchLanguage(key),
                    }}
                    getPopupContainer={() => document.body}
                >
                    <button
                        type="button"
                        className="dark-mode-toggle"
                        title={t('viewer.language')}
                    >
                        <TranslationOutlined />
                    </button>
                </Dropdown>
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
