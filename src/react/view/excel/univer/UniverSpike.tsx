// 开发/测试入口（?univer-spike）：不经 VSCode 消息桥直接驱动 Univer 路径。
// ?file=/test/fixtures/x.xlsx 时走完整 loadForUniver 导入管线（e2e 验收用），
// 否则加载内置 demo 数据。窗口暴露 __univerSpike 供 Playwright 探针。
import { useEffect, useRef, useState } from 'react';
import { UniverAdapter } from './adapter';
import { loadForUniver, type UniverLoadResult } from './loader';

const demoLoadResult = (): UniverLoadResult => ({
    workbookData: {
        id: 'spike-workbook',
        name: 'Spike',
        sheetOrder: ['sheet1'],
        styles: {
            s1: { bl: 1, fs: 14, cl: { rgb: '#c0392b' } },
            s2: { bg: { rgb: '#fff3cd' }, n: { pattern: '0.00"m";[Red]-0.00"m"' } },
        },
        sheets: {
            sheet1: {
                id: 'sheet1',
                name: 'Data',
                rowCount: 100,
                columnCount: 26,
                cellData: {
                    0: { 0: { v: 'Month', s: 's1' }, 1: { v: 'Sales', s: 's1' } },
                    1: { 0: { v: 'Jan' }, 1: { v: 10, s: 's2' } },
                    2: { 0: { v: 'Feb' }, 1: { v: 25, s: 's2' } },
                    3: { 0: { v: 'Mar' }, 1: { v: 18, s: 's2' } },
                    4: { 0: { v: 'Apr' }, 1: { v: 32, s: 's2' } },
                    5: { 0: { v: 'Total' }, 1: { f: '=SUM(B2:B5)' } },
                },
            },
        },
        locale: 'zhCN',
    } as never,
    hyperlinks: [],
});

export default function UniverSpike() {
    const containerRef = useRef<HTMLDivElement>(null);
    const adapterRef = useRef<UniverAdapter | null>(null);
    const [status, setStatus] = useState('booting');
    const [dark, setDark] = useState(false);

    const boot = async () => {
        const params = new URLSearchParams(window.location.search);
        const file = params.get('file');
        const t0 = performance.now();

        let loadResult: UniverLoadResult;
        if (file) {
            const resp = await fetch(file);
            if (!resp.ok) throw new Error(`fetch ${file}: ${resp.status}`);
            const buffer = await resp.arrayBuffer();
            const ext = file.split('.').pop() ?? 'xlsx';
            loadResult = await loadForUniver(buffer, ext, file.split('/').pop() ?? 'Workbook');
        } else {
            loadResult = demoLoadResult();
        }

        const adapter = UniverAdapter.create(containerRef.current!, { language: 'zh-cn' });
        await adapter.loadWorkbook(loadResult, { readOnly: params.has('readonly') });
        const bootMs = Math.round(performance.now() - t0);
        adapterRef.current = adapter;

        // 编辑会话（与 Excel.tsx 相同流程）：基线 + dirty/结构日志 + 保存探针
        const baseline = adapter.getWorkbookDataCopy();
        let dirtyCount = 0;
        const session = adapter.startEditSession(() => { dirtyCount += 1; });
        const spike = {
            adapter,
            univerAPI: adapter.univerAPI,
            univer: adapter.univer,
            loadResult,
            bootMs,
            get dirtyCount() { return dirtyCount; },
            get structuralSheetIds() { return [...session.structuralSheetIds]; },
            /** 完整保存链路（diff→apply→writeBuffer），返回 base64（e2e 校验用） */
            async saveProbe(): Promise<string> {
                const [{ diffWorkbook }, { applyDiffToWorkbook }] = await Promise.all([
                    import('./diff'),
                    import('./apply'),
                ]);
                const current = adapter.getWorkbookDataCopy() as never;
                const diff = diffWorkbook(baseline as never, current, session.structuralSheetIds);
                applyDiffToWorkbook(loadResult.originalWorkbook!, current, diff, { sheetIdMap: loadResult.sheetIdMap! });
                const buffer = await loadResult.originalWorkbook!.xlsx.writeBuffer();
                const bytes = new Uint8Array(buffer);
                let bin = '';
                for (let i = 0; i < bytes.length; i += 1) bin += String.fromCharCode(bytes[i]);
                return btoa(bin);
            },
        };
        (window as never as Record<string, unknown>).__univerSpike = spike;
        setStatus(`ready in ${bootMs}ms`);
        return adapter;
    };

    useEffect(() => {
        let disposed = false;
        const timer = setTimeout(() => {
            void boot().catch((e) => {
                console.error(e);
                if (!disposed) setStatus(`ERROR: ${(e as Error).message}`);
            });
        }, 0);
        return () => {
            disposed = true;
            clearTimeout(timer);
            adapterRef.current?.dispose();
            adapterRef.current = null;
        };
    }, []);

    return (
        <div style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
            <div style={{ padding: 4, display: 'flex', gap: 8, alignItems: 'center' }}>
                <span data-testid='spike-status'>{status}</span>
                <button
                    data-testid='spike-dark'
                    onClick={() => {
                        const next = !dark;
                        setDark(next);
                        adapterRef.current?.setDarkMode(next);
                    }}
                >
                    dark
                </button>
                <button
                    data-testid='spike-recreate'
                    onClick={() => {
                        adapterRef.current?.dispose();
                        adapterRef.current = null;
                        setStatus('recreating');
                        setTimeout(() => {
                            void boot().catch((e) => setStatus(`ERROR: ${(e as Error).message}`));
                        }, 0);
                    }}
                >
                    recreate
                </button>
            </div>
            <div ref={containerRef} style={{ flex: 1, overflow: 'hidden' }} />
        </div>
    );
}
