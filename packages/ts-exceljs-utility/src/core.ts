
export interface CellAddressRange{
  top: number;
  left: number;
  bottom: number;
  right: number;
}

/**
 * セルの範囲を表すインターフェース
 */
export type CellAddress = 
  string | //文字列表現。("A1", "A1:B2")
  {
    top: number;
    left: number;
  } | //単一表現
  CellAddressRange; //範囲表現

/**
 * CellAddressが文字列表現かどうかを調べる
 * @param address 
 * @returns 
 */
export function isCellAddressString(address: CellAddress): address is string {
  return typeof address === 'string';
}

/**
 * CellAddressが単一表現かどうかを調べる
 * @param address 
 * @returns 
 */
export function isCellAddressOne(address: CellAddress): address is { top: number; left: number; } {
  return typeof address === 'object' && 'top' in address && 'left' in address;
}

/**
 * CellAddressが範囲表現かどうかを調べる
 * @param address 
 * @returns 
 */
export function isCellAddressRange(address: CellAddress): address is { top: number; left: number; bottom: number; right: number; } {
  return typeof address === 'object' && 'top' in address && 'left' in address && 'bottom' in address && 'right' in address;
}

