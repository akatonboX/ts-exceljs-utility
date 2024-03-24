import ExcelJS from "exceljs";
import { CellAddress } from "./core";
export declare function createAddImageParam(url: string): Promise<{
    extension: 'jpeg' | 'png' | 'gif';
    base64: string;
}>;
export declare function createAddImageParam(blob: Blob): Promise<{
    extension: 'jpeg' | 'png' | 'gif';
    base64: string;
}>;
export declare function addImageToCell(worksheet: ExcelJS.Worksheet, imageId: number, imageWidth: number, imageHeight: number, targetAddress: CellAddress, standardFontWidth?: number): void;
