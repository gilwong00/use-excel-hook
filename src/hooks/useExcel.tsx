import { Workbook, Column, Cell, Style } from 'exceljs';
import { saveAs } from 'file-saver';

type Header = Pick<Column, 'header' | 'key'>;

interface Options {
  headers?: Array<Header>;
  data: any;
  fileName?: string;
  headerStyles?: Style;
  rowStyles?: Style;
}

export default function useExcel() {
  const defaults = {
    header: {
      font: {
        size: 12
      }
    },
    row: {
      font: {
        size: 12
      }
    }
  };

  const generateHeaders = (data: Array<any>): Array<Header> => {
    if (data && data[0]) {
      const value = data[0];
      const headers: Array<Header> = Object.keys(value).map((key) => ({
        header: key,
        key
      }));
      return headers;
    } else {
      return [];
    }
  };

  const exportToExcel = async ({
    headers = [],
    data = [],
    fileName = new Date().toISOString(),
    headerStyles,
    rowStyles
  }: Options): Promise<void> => {
    if (!data) {
      throw new Error('data is required');
    }

    if (!headers.length) {
      headers = generateHeaders(data);
    }

    const wb = new Workbook();
    const ws = wb.addWorksheet();
    ws.columns = headers as Array<Column>;
    ws.addRows(data);

    ws.columns.forEach((column) => (column.width = 25));
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
      // add style to headers only
      if (rowNum === 1) {
        row.eachCell((cell: Cell) => {
          cell.font = headerStyles?.font ?? defaults.header.font;

          if (headerStyles && headerStyles.fill) {
            cell.fill = headerStyles.fill;
          }
        });
      } else {
        row.eachCell((cell: Cell) => {
          cell.font = rowStyles?.font ?? defaults.row.font;

          if (rowStyles && rowStyles.fill) {
            cell.fill = rowStyles.fill;
          }
        });
      }
    });

    const buffer = await wb.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `${fileName}.xlsx`);
  };

  return {
    exportToExcel
  };
}
