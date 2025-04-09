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

// Helper function to check if an RGBA color is black with opacity
const isBlackWithOpacity = (rgbaColor: string): boolean => {
    const match = rgbaColor.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
    if (!match) return false; // Not in rgba format with opacity
    const r = parseInt(match[1], 10);
    const g = parseInt(match[2], 10);
    const b = parseInt(match[3], 10);
    const a = parseFloat(match[4]);
    return r === 0 && g === 0 && b === 0 && a >= 0 && a <= 1;
};

// Helper function to check if an RGB color is black without opacity
const isBlackWithoutOpacity = (rgbColor: string): boolean => {
    const match = rgbColor.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (!match) return false; // Not in rgb format
    const r = parseInt(match[1], 10);
    const g = parseInt(match[2], 10);
    const b = parseInt(match[3], 10);
    return r === 0 && g === 0 && b === 0;
};

// Helper function to check if a color is exactly black (with or without opacity)
export const isShadeOfBlack = (color: string): boolean => {
    return isBlackWithOpacity(color) || isBlackWithoutOpacity(color);
};