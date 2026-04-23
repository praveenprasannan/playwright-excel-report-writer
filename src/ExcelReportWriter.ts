import { ExcelWriterConfig } from './types/WriterConfig';
import { TestResultRow } from './types/TestResultRow';
import { ExcelEngine } from './core/ExcelEngine';

export class ExcelReportWriter {
  static async update(
    config: ExcelWriterConfig,
    data: TestResultRow
  ): Promise<void> {
    await ExcelEngine.update(config, data);
  }
}
