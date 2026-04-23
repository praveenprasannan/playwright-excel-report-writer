import type { ExcelWriterConfig } from './types/WriterConfig.js';
import type { TestResultRow } from './types/TestResultRow.js';
import { ExcelEngine } from './core/ExcelEngine.js';

export class ExcelReportWriter {
  static async update(
    config: ExcelWriterConfig,
    data: TestResultRow
  ): Promise<void> {
    await ExcelEngine.update(config, data);
  }
}
