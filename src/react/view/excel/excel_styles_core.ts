/** 引擎无关的颜色工具（从 excel_styles 抽出，供 univer/* 使用） */

export function colorToHexCore(argbColor?: { argb?: string }): string | undefined {
    if (!argbColor?.argb) return undefined;
    const argb = argbColor.argb.replace(/^#/, '');
    if (argb.length === 8) return `#${argb.slice(2).toLowerCase()}`;
    if (argb.length === 6) return `#${argb.toLowerCase()}`;
    return undefined;
}

export function hexToArgb(hex?: string): string | undefined {
    if (!hex) return undefined;
    const normalized = hex.replace(/^#/, '');
    if (normalized.length === 6) return `FF${normalized.toUpperCase()}`;
    if (normalized.length === 8) return normalized.toUpperCase();
    return undefined;
}
