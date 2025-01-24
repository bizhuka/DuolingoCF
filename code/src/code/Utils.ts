/* global document, Excel, Blob, URL */

export const FRONT_COLUMN: number = 0;
export const BACK_COLUMN: number = 1;
export const IMAGE_COLUMN: number = 2;
export const SOUND_COLUMN: number = 5;

// interface TableRow {
//   [key: string]: any; // Dynamic keys, you can replace 'any' with specific types if known
// }

export interface IInfo{
  info_text: string;
}

export interface Card {
  Front: string;
  Back: string;
  Image: string;
  Hint: string;
  Context: string;
  Sound: string;
  Exported: boolean;
}

export function createCard(Front: string): Card {
  return {
    Front: Front,
    Back: "",
    Image: "",
    Hint: "",
    Context: "",
    Sound: "",
    Exported: false,
  };
}

interface PixabayImage {
  id: number;
  // webformatURL: string;
  previewURL: string;
}

export interface PixabayResponse {
  total: number;
  totalHits: number;
  hits: PixabayImage[];
}

interface GoogleImage {
  link: string;
  image: {
    thumbnailLink: string;
  };
}

export interface GoogleResponse {
  kind: string;
  items: GoogleImage[];
}

export function downloadBlob(csvRows: any[][], filename: string) {
  const as_line = csvRows.map((obj) =>
    Object.values(obj)
      .map((value) => `"${value.toString().replace(/"/g, '""')}"`)
      .join(";")
  );
  const blob = new Blob(["\uFEFF" + as_line.join("\n")], { type: "text/csv;charset=utf-8" });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(url);
  document.body.removeChild(a);
}

export async function get_table_data(context: Excel.RequestContext): Promise<Card[]> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemAt(0); // Assuming the table you want is the first table in the worksheet

  // Get the header row range and data body range of the table
  const headerRange = table.getHeaderRowRange();
  const dataBodyRange = table.getDataBodyRange();
  headerRange.load("values");
  dataBodyRange.load("values");

  // Sync the context
  await context.sync();

  // Extract header and data values
  const headers = headerRange.values[0] as (keyof Card)[];
  const data = dataBodyRange.values;

  // Convert to array of objects
  const dataObjects = data.map((row) => {
    let obj: Card = {} as Card;
    row.forEach((cell, index) => {
      (obj as any)[headers[index]] = cell;
    });
    return obj;
  });
  return dataObjects;
}

export async function set_table_data(context: Excel.RequestContext, all_data: any[]): Promise<void> {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemAt(0); // Assuming the table you want is the first table in the worksheet

  let bodyRange = table.getDataBodyRange();
  bodyRange.load(["rowCount", "address"]);
  await context.sync();

  // Add empty rows or with default values as needed
  if (bodyRange.rowCount !== all_data.length) {
    const headerRange = table.getHeaderRowRange();
    const newRange = headerRange.getResizedRange(all_data.length, 0);
    table.resize(newRange);
    newRange.load("address");
    await context.sync();

    const begCellAddress = bodyRange.address.split("!")[1].split(":")[0];
    const endCellAddress = newRange.address.split("!")[1].split(":")[1];

    bodyRange = sheet.getRange(`${begCellAddress}:${endCellAddress}`);
  }

  bodyRange.values = all_data.map((item) => Object.values(item));
  await context.sync();
}

export function removeHTMLTags(text: string) {
  return text.replace(/<[^>]*>/g, "");
}

export function containsImageExtension(url: string) {
  const imageExtensions = /\.(jpg|jpeg|png|gif|bmp|webp)/i;
  return imageExtensions.test(url);
}

export function getDomain(urlString: string): string {
  const url = new URL(urlString);
  return url.hostname;
}

export function isValidURL(url: string): boolean {
  try {
    new URL(url);
    return true;
  } catch (_) {
    return false;
  }
}

export function columnIndexToLetter(columnIndex: number) {
  let columnLetter = "";
  let tempIndex = columnIndex + 1; // Convert to 1-based index

  while (tempIndex > 0) {
    let remainder = (tempIndex - 1) % 26;
    columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
    tempIndex = Math.floor((tempIndex - 1) / 26);
  }

  return columnLetter;
}
