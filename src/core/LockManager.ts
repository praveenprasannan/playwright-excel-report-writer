import fs from 'fs';

export class LockManager {
  static async acquire(
    lockPath: string,
    timeoutMs = 30000
  ): Promise<number> {
    const start = Date.now();

    while (Date.now() - start < timeoutMs) {
      try {
        return fs.openSync(lockPath, 'wx');
      } catch {
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }

    throw new Error(
      `[playwright-excel-report-writer] Failed to acquire lock within ${timeoutMs}ms`
    );
  }

  static release(lockPath: string, fd?: number): void {
    try {
      if (fd) fs.closeSync(fd);
      if (fs.existsSync(lockPath)) fs.unlinkSync(lockPath);
    } catch {
      // swallow — lock cleanup should never fail tests
    }
  }
}
