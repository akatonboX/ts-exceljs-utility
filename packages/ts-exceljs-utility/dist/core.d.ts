export interface CellAddressRange {
    top: number;
    left: number;
    bottom: number;
    right: number;
}
/**
 * セルの範囲を表すインターフェース
 */
export type CellAddress = string | //文字列表現。("A1", "A1:B2")
{
    top: number;
    left: number;
} | //単一表現
CellAddressRange;
/**
 * CellAddressが文字列表現かどうかを調べる
 * @param address
 * @returns
 */
export declare function isCellAddressString(address: CellAddress): address is string;
/**
 * CellAddressが単一表現かどうかを調べる
 * @param address
 * @returns
 */
export declare function isCellAddressOne(address: CellAddress): address is {
    top: number;
    left: number;
};
/**
 * CellAddressが範囲表現かどうかを調べる
 * @param address
 * @returns
 */
export declare function isCellAddressRange(address: CellAddress): address is {
    top: number;
    left: number;
    bottom: number;
    right: number;
};
