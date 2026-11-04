// sql.js's build with the WebAssembly inlined (no .wasm file to fetch) has the package's types.
declare module 'sql.js/dist/sql-wasm-inline.js' {
  import initSqlJs from 'sql.js'
  export default initSqlJs
}
