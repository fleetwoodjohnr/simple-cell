import { executeSql, type SqlInput } from './sql-engine'
self.onmessage = async (e: MessageEvent<SqlInput>) => {
  try { postMessage({ result: await executeSql(e.data) }) }
  catch (error) { postMessage({ error: (error as Error).message }) }
}
