import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

import type { ExcelWriterConfig } from '../types/WriterConfig.js';
import type { TestResultRow } from '../types/TestResultRow.js';
import { LockManager } from './LockManager.js';
import { Formatter } from './Formatter.js';

/**
 * ExcelJS image extension type (official)
 */
type ExcelImageExtension = 'png' | 'jpeg' | 'gif';

/**
 * ExcelJS image positioning helper
 * This avoids Anchor typing issues while preserving runtime behavior.
 */
function imagePosition(
  col: number,
  row: number,
  width?: number,
  height?: number
): ExcelJS.ImagePosition | ExcelJS.ImageRange {
  if (typeof width === 'number' && typeof height === 'number') {
    return {
      tl: { col, row },
      ext: { width, height }
    } as ExcelJS.ImagePosition;
  }

  return {
    tl: { col, row },
    br: { col: col + 1, row: row + 1 }
  } as ExcelJS.ImageRange;
}

export class ExcelEngine {
  static async update(
    config: ExcelWriterConfig,
    data: TestResultRow
  ): Promise<void> {
    const dir = config.file.directory;
    const filePath = path.join(dir, config.file.fileName);
    const lockPath = path.join(dir, '.excel-lock');

    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    const lockFd = await LockManager.acquire(
      lockPath,
      config.lockTimeoutMs
    );

    try {
      const workbook = new ExcelJS.Workbook();

      /* ---------------- Sheet name ---------------- */
      let sheetName: string;
      if (config.sheet.strategy === 'by-browser') {
        if (!config.sheet.namePattern) {
          throw new Error(
            'sheet.namePattern must be provided for by-browser strategy'
          );
        }
        sheetName = config.sheet.namePattern.replace(
          '{browser}',
          String(data.browser).trim()
        );
      } else {
        if (!config.sheet.name) {
          throw new Error(
            'sheet.name must be provided for single-sheet strategy'
          );
        }
        sheetName = config.sheet.name;
      }

      /* ---------------- Column indexes ---------------- */
      const rowKeyColIndex =
        config.columns.findIndex(c => c.key === config.rowKey) + 1;

      if (rowKeyColIndex === 0) {
        throw new Error(
          `Row key '${config.rowKey}' not found in columns`
        );
      }

      const resultColIndex =
        config.columns.findIndex(c => c.key === 'result') + 1;

      const imageConfig = config.image;
      const imageColIndex =
        imageConfig
          ? config.columns.findIndex(c => c.key === imageConfig.columnKey)
          : -1;

      /* ---------------- Preserve rows & images ---------------- */
      let rowsToPreserve: {
        values: ExcelJS.CellValue[];
        height?: number;
        resultText: string;
        isNew?: boolean;
      }[] = [];

      let imagesToPreserve: {
        buffer: ExcelJS.Buffer;
        extension: ExcelImageExtension;
        col: number;
        row: number;
        width?: number;
        height?: number;
      }[] = [];

      let targetIndex: number | null = null;

      /* ---------------- Read existing worksheet ---------------- */
      if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
        const oldSheet = workbook.getWorksheet(sheetName);

        if (oldSheet) {
          oldSheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            const values = Array.isArray(row.values)
              ? row.values.slice(1)
              : [];

            const keyValue = String(
              row.getCell(rowKeyColIndex).text || ''
            ).trim();

            if (
              keyValue ===
              String(data[config.rowKey] ?? '').trim()
            ) {
              targetIndex = rowsToPreserve.length;
            }

            rowsToPreserve.push({
              values,
              height: row.height,
              resultText: String(
                row.getCell(resultColIndex).text || ''
              ).toLowerCase()
            });
          });

          /* -------- Preserve images (ExcelJS workaround) -------- */
          for (const img of oldSheet.getImages()) {
            const range = img.range as any;
            const imgRow =
              typeof range?.tl?.row === 'number'
                ? range.tl.row + 1
                : null;

            const isTargetRow =
              targetIndex !== null && imgRow === targetIndex + 2;

            if (!isTargetRow && imgRow !== null) {
              const image = workbook.getImage(
                Number(img.imageId)
              );

              imagesToPreserve.push({
                buffer: image.buffer as ExcelJS.Buffer,
                extension: image.extension as ExcelImageExtension,
                col: range.tl.col,
                row: imgRow,
                width: range.ext?.width,
                height: range.ext?.height
              });
            }
          }

          workbook.removeWorksheet(oldSheet.id);
        }
      }

      /* ---------------- Updated row ---------------- */
      const isPass = String(data.result)
        .toLowerCase()
        .includes('pass');

      const updatedRow = {
        values: config.columns.map(col =>
          col.type === 'image' ? '' : data[col.key] ?? ''
        ),
        height: isPass ? undefined : 200,
        resultText: String(data.result).toLowerCase(),
        isNew: true
      };

      if (targetIndex !== null) {
        rowsToPreserve[targetIndex] = updatedRow;
      } else {
        rowsToPreserve.push(updatedRow);
      }

      /* ---------------- Rebuild worksheet ---------------- */
      const sheet = workbook.addWorksheet(sheetName);

      sheet.addRow(config.columns.map(c => c.header));
      sheet.views = [{ state: 'frozen', ySplit: 1 }];

      Formatter.formatHeader(sheet.getRow(1), config.styles);

      config.columns.forEach((col, idx) => {
        if (col.width) sheet.getColumn(idx + 1).width = col.width;
      });

      rowsToPreserve.forEach((rowData, index) => {
        const row = sheet.addRow(rowData.values);

        if (rowData.height) {
          row.height = rowData.height;
        }

        Formatter.formatRow(
          row,
          rowData.resultText.includes('pass'),
          config.styles,
          resultColIndex
        );

        /* -------- Restore preserved images -------- */
        imagesToPreserve
          .filter(img => img.row === index + 2)
          .forEach(img => {
            const imageId = workbook.addImage({
              buffer: img.buffer,
              extension: img.extension
            });

            sheet.addImage(
              imageId,
              imagePosition(
                img.col,
                row.number - 1,
                img.width,
                img.height
              )
            );
          });

        /* -------- Add screenshot for FAIL -------- */
        if (
          rowData.isNew &&
          !isPass &&
          imageConfig &&
          data.screenshot &&
          fs.existsSync(data.screenshot)
        ) {
          const screenshotId = workbook.addImage({
            filename: data.screenshot,
            extension: 'png'
          });

          sheet.addImage(
            screenshotId,
            imagePosition(
              imageColIndex,
              row.number - 1,
              imageConfig.width,
              imageConfig.height
            )
          );
        }
      });

      await workbook.xlsx.writeFile(filePath);
    } finally {
      LockManager.release(lockPath, lockFd);
    }
  }
}
