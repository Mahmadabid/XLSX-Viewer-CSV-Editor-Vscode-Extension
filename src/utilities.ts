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

// Helper function to check if an RGBA color is black with opacity.
const isBlackWithOpacity = (rgbaColor: string): boolean => {
    const match = rgbaColor.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
    if (!match) return false; // Not in rgba format with opacity.
    const r = parseInt(match[1], 10);
    const g = parseInt(match[2], 10);
    const b = parseInt(match[3], 10);
    const a = parseFloat(match[4]);
    return r === 0 && g === 0 && b === 0 && a >= 0 && a <= 1;
};

// Helper function to check if an RGB color is black without opacity.
const isBlackWithoutOpacity = (rgbColor: string): boolean => {
    const match = rgbColor.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
    if (!match) return false; // Not in rgb format.
    const r = parseInt(match[1], 10);
    const g = parseInt(match[2], 10);
    const b = parseInt(match[3], 10);
    return r === 0 && g === 0 && b === 0;
};

/**
 * Helper function to check if a hex color represents a shade of black.
 * This function accepts a hex string with or without a leading '#' and
 * supports both 6-digit (e.g. "000000") and 8-digit (e.g. "FF000000") formats.
 * It determines if the red, green, and blue channels are below a specified threshold.
 */
const isHexBlack = (hexColor: string): boolean => {
    if (!hexColor) return false;
    
    // Remove a leading '#' if present.
    let cleaned = hexColor.startsWith('#') ? hexColor.substring(1) : hexColor;
    
    // If the color includes alpha (8 digits), ignore the alpha channel.
    if (cleaned.length === 8) {
        cleaned = cleaned.substring(2);
    }
    
    if (cleaned.length !== 6) return false;
    
    // Parse the RGB components.
    const r = parseInt(cleaned.substring(0, 2), 16);
    const g = parseInt(cleaned.substring(2, 4), 16);
    const b = parseInt(cleaned.substring(4, 6), 16);
    
    // Define an acceptable threshold for near-black.
    const threshold = 20; // Adjust this value as needed.
    
    return r <= threshold && g <= threshold && b <= threshold;
};

/**
 * Checks if a color (in rgba, rgb, or hex format) represents a shade of black.
 */
export const isShadeOfBlack = (color: string): boolean => {
    return isBlackWithOpacity(color) || isBlackWithoutOpacity(color) || isHexBlack(color);
};