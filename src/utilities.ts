export function convertARGBToRGBA(argb: string): string {
    if (!argb || argb.length !== 8) {
        return 'rgba(0, 0, 0, 1)';
    }
    
    const a = parseInt(argb.substring(0, 2), 16) / 255;
    const r = parseInt(argb.substring(2, 4), 16);
    const g = parseInt(argb.substring(4, 6), 16);
    const b = parseInt(argb.substring(6, 8), 16);
    
    return `rgba(${r}, ${g}, ${b}, ${a})`;
}

export function isShadeOfBlack(color: string): boolean {
    // Parse RGB/RGBA values
    const match = color.match(/rgba?KATEX_INLINE_OPEN(\d+),\s*(\d+),\s*(\d+)(?:,\s*[\d.]+)?KATEX_INLINE_CLOSE/);
    if (!match) return false;
    
    const r = parseInt(match[1]);
    const g = parseInt(match[2]);
    const b = parseInt(match[3]);
    
    // Consider a color as "shade of black" only if ALL RGB values are very low
    // The color #c9daf8 = rgb(201, 218, 248) should NOT be considered black
    const threshold = 50; // Very dark colors only
    
    // ALL components must be below threshold to be considered "black"
    return r <= threshold && g <= threshold && b <= threshold;
}