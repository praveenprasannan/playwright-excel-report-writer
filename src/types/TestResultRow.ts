export type TestResultRow = {
  testCase: string;
  browser: string;
  result: string;
  screenshot?: string;

  // allow any additional user-defined columns
  [key: string]: any;
};
