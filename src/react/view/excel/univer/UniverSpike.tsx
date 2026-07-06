// Phase 0 spike: minimal Univer sheet mounted with the full preset set the
// migration plan targets. Reached via ?univer-spike in dev / spike builds.
// Exposes window.__univerSpike for automated probing (Playwright).
import { useEffect, useRef, useState } from 'react';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/presets/preset-sheets-core';
import { UniverSheetsFilterPreset } from '@univerjs/presets/preset-sheets-filter';
import { UniverSheetsSortPreset } from '@univerjs/presets/preset-sheets-sort';
import { UniverSheetsFindReplacePreset } from '@univerjs/presets/preset-sheets-find-replace';
import { UniverSheetsDataValidationPreset } from '@univerjs/presets/preset-sheets-data-validation';
import { UniverSheetsHyperLinkPreset } from '@univerjs/presets/preset-sheets-hyper-link';
import { UniverSheetsDrawingPreset } from '@univerjs/presets/preset-sheets-drawing';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/presets/preset-sheets-conditional-formatting';
import CoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN';
import CoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import FilterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN';
import FilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import SortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN';
import SortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import FindReplaceZhCN from '@univerjs/preset-sheets-find-replace/locales/zh-CN';
import FindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US';
import DataValidationZhCN from '@univerjs/preset-sheets-data-validation/locales/zh-CN';
import DataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import HyperLinkZhCN from '@univerjs/preset-sheets-hyper-link/locales/zh-CN';
import HyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';
import DrawingZhCN from '@univerjs/preset-sheets-drawing/locales/zh-CN';
import DrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import CfZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN';
import CfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';

import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/presets/lib/styles/preset-sheets-filter.css';
import '@univerjs/presets/lib/styles/preset-sheets-sort.css';
import '@univerjs/presets/lib/styles/preset-sheets-find-replace.css';
import '@univerjs/presets/lib/styles/preset-sheets-data-validation.css';
import '@univerjs/presets/lib/styles/preset-sheets-hyper-link.css';
import '@univerjs/presets/lib/styles/preset-sheets-drawing.css';
import '@univerjs/presets/lib/styles/preset-sheets-conditional-formatting.css';

const demoWorkbookData = () => ({
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
                0: {
                    0: { v: 'Month', s: 's1' },
                    1: { v: 'Sales', s: 's1' },
                },
                1: { 0: { v: 'Jan' }, 1: { v: 10, s: 's2' } },
                2: { 0: { v: 'Feb' }, 1: { v: 25, s: 's2' } },
                3: { 0: { v: 'Mar' }, 1: { v: 18, s: 's2' } },
                4: { 0: { v: 'Apr' }, 1: { v: 32, s: 's2' } },
                5: { 0: { v: 'Total' }, 1: { f: '=SUM(B2:B5)' } },
            },
        },
    },
    locale: 'zhCN',
});

export default function UniverSpike() {
    const containerRef = useRef<HTMLDivElement>(null);
    const apiRef = useRef<ReturnType<typeof createUniver> | null>(null);
    const [status, setStatus] = useState('booting');

    const boot = () => {
        const t0 = performance.now();
        const inst = createUniver({
            locale: LocaleType.ZH_CN,
            locales: {
                [LocaleType.ZH_CN]: mergeLocales(
                    CoreZhCN, FilterZhCN, SortZhCN, FindReplaceZhCN,
                    DataValidationZhCN, HyperLinkZhCN, DrawingZhCN, CfZhCN,
                ),
                [LocaleType.EN_US]: mergeLocales(
                    CoreEnUS, FilterEnUS, SortEnUS, FindReplaceEnUS,
                    DataValidationEnUS, HyperLinkEnUS, DrawingEnUS, CfEnUS,
                ),
            },
            presets: [
                UniverSheetsCorePreset({ container: containerRef.current! }),
                UniverSheetsFilterPreset(),
                UniverSheetsSortPreset(),
                UniverSheetsFindReplacePreset(),
                UniverSheetsDataValidationPreset(),
                UniverSheetsHyperLinkPreset(),
                UniverSheetsDrawingPreset(),
                UniverSheetsConditionalFormattingPreset(),
            ],
        });
        inst.univerAPI.createWorkbook(demoWorkbookData() as never);
        const bootMs = Math.round(performance.now() - t0);
        apiRef.current = inst;
        (window as never as Record<string, unknown>).__univerSpike = {
            univerAPI: inst.univerAPI,
            univer: inst.univer,
            bootMs,
        };
        setStatus(`ready in ${bootMs}ms`);
        return inst;
    };

    useEffect(() => {
        const inst = boot();
        return () => {
            inst.univer.dispose();
            apiRef.current = null;
        };
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    return (
        <div style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
            <div style={{ padding: 4, display: 'flex', gap: 8, alignItems: 'center' }}>
                <span data-testid='spike-status'>{status}</span>
                <button
                    data-testid='spike-dark'
                    onClick={() => {
                        const api = apiRef.current?.univerAPI as never as { toggleDarkMode(v: boolean): void };
                        (window as never as Record<string, boolean>).__dark = !(window as never as Record<string, boolean>).__dark;
                        api?.toggleDarkMode((window as never as Record<string, boolean>).__dark);
                    }}
                >
                    dark
                </button>
                <button
                    data-testid='spike-recreate'
                    onClick={() => {
                        // leak probe: dispose + recreate, used by fileChange reload path
                        apiRef.current?.univer.dispose();
                        apiRef.current = null;
                        setStatus('recreating');
                        setTimeout(() => boot(), 0);
                    }}
                >
                    recreate
                </button>
            </div>
            <div ref={containerRef} style={{ flex: 1, overflow: 'hidden' }} />
        </div>
    );
}
