// Empty stand-in for @napi-rs/canvas, a server-only library.
// The PowerPoint viewer only uses it when running in Node.js, never in Chrome.
const unavailable = () => { throw new Error('@napi-rs/canvas is not available in the browser'); };
export const createCanvas = unavailable;
export const loadImage = unavailable;
export const GlobalFonts = { registerFromPath: () => false, register: () => null, families: [] };
export class Canvas {}
export class Image {}
export class ImageData {}
export class Path2D {}
export class DOMMatrix {}
export class DOMPoint {}
export class DOMRect {}
export default { createCanvas, loadImage, GlobalFonts, Canvas, Image, ImageData, Path2D, DOMMatrix, DOMPoint, DOMRect };
