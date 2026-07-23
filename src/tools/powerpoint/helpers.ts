/**
 * Office.js PowerPoint wrappers — thin abstraction over PowerPoint.run().
 */

export async function pptRun<T>(fn: (context: PowerPoint.RequestContext) => Promise<T>): Promise<T> {
  return PowerPoint.run(fn);
}
