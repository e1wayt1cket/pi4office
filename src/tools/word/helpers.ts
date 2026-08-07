/**
 * Office.js Word wrappers — thin abstraction over Word.run().
 */

export async function wordRun<T>(fn: (context: Word.RequestContext) => Promise<T>): Promise<T> {
  return Word.run(fn);
}
