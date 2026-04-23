import ExcelJS from 'exceljs';
import { StyleConfig } from '../types/WriterConfig';

export class Formatter {

  static formatHeader(
    row: ExcelJS.Row,
    styles?: StyleConfig
  ): void {
    row.eachCell(cell => {
      cell.font = {
        bold: styles?.headerBold ?? true,
        size: styles?.headerSize ?? 12
      };

      cell.alignment = {
        horizontal: 'left',
        vertical: 'middle',
        wrapText: true
      };
    });
  }

  static formatRow(
    row: ExcelJS.Row,
    isPass: boolean,
    styles?: StyleConfig,
    resultColIndex?: number
  ): void {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.alignment = {
        vertical: 'top',
        horizontal: 'left',
        wrapText: true
      };

      cell.border = {
        top: { style: styles?.borderStyle ?? 'thin' },
        left: { style: styles?.borderStyle ?? 'thin' },
        bottom: { style: styles?.borderStyle ?? 'thin' },
        right: { style: styles?.borderStyle ?? 'thin' }
      };
    });

    if (resultColIndex) {
      const resultCell = row.getCell(resultColIndex);
      resultCell.font = {
        bold: true,
        color: { argb: 'FFFFFFFF' }
      };

      resultCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: isPass
            ? styles?.passFill ?? 'FF00C853'
            : styles?.failFill ?? 'FFD50000'
        }
      };
    }
  }
}
