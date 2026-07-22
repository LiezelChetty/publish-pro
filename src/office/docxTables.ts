import { getChildrenByLocalName } from "./docxPackage";

export function getTableRows(table: Element) {
  return getChildrenByLocalName(table, "tr");
}

export function getTableCells(row: Element) {
  return getChildrenByLocalName(row, "tc");
}
