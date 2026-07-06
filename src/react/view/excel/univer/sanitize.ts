/**
 * xlsx 预清洗：ExcelJS（fork 与上游）在 load 含图表 drawing 的文件时会在
 * reconcile 阶段崩溃（drawing.anchors undefined）。打开前用 jszip 探测，
 * 把引用了 graphicFrame drawing 的 <drawing r:id/> 从 sheet XML 中剥离，
 * 使文件从「打不开」变为「可查看」。同时探测透视表 / VBA，供保存前警告用。
 *
 * 注意：清洗只影响喂给 ExcelJS 的副本；原始字节由调用方保留。
 */
import JSZip from 'jszip';

export interface SanitizeResult {
    /** 可安全交给 ExcelJS.load 的 buffer（未修改时即原 buffer） */
    buffer: ArrayBuffer;
    /** 探测到的会被 ExcelJS 丢弃/无法处理的特性 */
    lossy: {
        charts: boolean;
        pivotTables: boolean;
        vba: boolean;
    };
    /** 是否实际做过剥离（true 时喂入的已不是原文件） */
    modified: boolean;
}

const DRAWING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';

export async function sanitizeXlsxForExcelJs(buffer: ArrayBuffer): Promise<SanitizeResult> {
    const lossy = { charts: false, pivotTables: false, vba: false };
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer);
    } catch {
        // 非 zip / 损坏文件：交给 ExcelJS 自行报错
        return { buffer, lossy, modified: false };
    }

    const names = Object.keys(zip.files);
    lossy.pivotTables = names.some(n => n.startsWith('xl/pivotTables/') || n.startsWith('xl/pivotCache/'));
    lossy.vba = names.some(n => /vbaProject\.bin$/i.test(n));

    // 找出含 graphicFrame（图表等非图片锚点）的 drawing part
    const chartDrawings = new Set<string>();
    for (const name of names) {
        if (!/^xl\/drawings\/drawing\d+\.xml$/.test(name)) continue;
        const xml = await zip.file(name)!.async('string');
        if (xml.includes('graphicFrame')) {
            chartDrawings.add(name);
        }
    }
    lossy.charts = chartDrawings.size > 0 || names.some(n => n.startsWith('xl/charts/'));

    if (!chartDrawings.size) {
        return { buffer, lossy, modified: false };
    }

    // 逐 worksheet：若其 rels 指向 chart drawing，剥离 sheet XML 里的
    // <drawing/> 引用与对应 Relationship 条目
    let modified = false;
    for (const relName of names) {
        const m = relName.match(/^xl\/worksheets\/_rels\/(sheet\d+)\.xml\.rels$/);
        if (!m) continue;
        let relXml = await zip.file(relName)!.async('string');
        const relIds: string[] = [];
        const relRegex = /<Relationship\b[^>]*\/?>(?:<\/Relationship>)?/g;
        for (const relTag of relXml.match(relRegex) ?? []) {
            if (!relTag.includes(DRAWING_REL_TYPE)) continue;
            const target = relTag.match(/Target="([^"]+)"/)?.[1] ?? '';
            const resolved = target.replace(/^\.\.\//, 'xl/').replace(/^\//, '').replace(/^xl\/xl\//, 'xl/');
            const normalized = resolved.startsWith('xl/') ? resolved : `xl/drawings/${resolved.split('/').pop()}`;
            if (chartDrawings.has(normalized)) {
                const id = relTag.match(/Id="([^"]+)"/)?.[1];
                if (id) {
                    relIds.push(id);
                    relXml = relXml.replace(relTag, '');
                }
            }
        }
        if (!relIds.length) continue;
        zip.file(relName, relXml);

        const sheetName = `xl/worksheets/${m[1]}.xml`;
        const sheetFile = zip.file(sheetName);
        if (!sheetFile) continue;
        let sheetXml = await sheetFile.async('string');
        for (const id of relIds) {
            sheetXml = sheetXml.replace(new RegExp(`<drawing[^>]*r:id="${id}"[^>]*/>`, 'g'), '');
        }
        zip.file(sheetName, sheetXml);
        modified = true;
    }

    // 关键：drawing 部件本身也必须移除 —— ExcelJS 按文件名模式扫描并解析
    // drawing part，graphicFrame-only 的 drawing 会解析出 undefined 模型，
    // reconcile 阶段仍会崩溃（即使 sheet 已不引用它）
    for (const name of chartDrawings) {
        zip.remove(name);
        const relPath = name.replace(/^xl\/drawings\//, 'xl/drawings/_rels/') + '.rels';
        if (zip.file(relPath)) zip.remove(relPath);
        modified = true;
    }

    if (!modified) {
        return { buffer, lossy, modified: false };
    }
    const out = await zip.generateAsync({ type: 'arraybuffer' });
    return { buffer: out, lossy, modified: true };
}
