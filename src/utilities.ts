/**
 * Converts an Excel ARGB color string ("AARRGGBB") to a CSS rgba() string.
 */
export function convertARGBToRGBA(argb: string): string {
    if (argb.length !== 8) {
        return `#${argb}`;
    }
    const alpha = parseInt(argb.substring(0, 2), 16) / 255;
    const red = parseInt(argb.substring(2, 4), 16);
    const green = parseInt(argb.substring(4, 6), 16);
    const blue = parseInt(argb.substring(6, 8), 16);
    
    return `rgba(${red}, ${green}, ${blue}, ${alpha.toFixed(2)})`;
} 