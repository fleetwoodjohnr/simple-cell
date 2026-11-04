# SQL over workbook data

Open **Automate → SQL**. All named tables in the workbook are available by name.
For ordinary cells, select a rectangle (or a cell in a connected dataset) first.
The dialog displays an editable source range and its SQL name, normally `Selection`.
A numbered suffix avoids collisions with existing table names.

Check **First row contains headers** to use that row as column names. Without headers,
columns are `Column1`, `Column2`, etc. Blank and duplicate headers receive unique names
shown in the source list. Click a column name to insert its quoted SQL identifier.

```sql
SELECT Region, SUM(Units) AS Total
FROM Selection
WHERE Units > 0
GROUP BY Region
ORDER BY Total DESC;
```

**Preview** (Ctrl+Enter or Cmd+Enter) shows the first 100 returned rows, the full row count,
and pending source changes. It never changes the workbook. **Run** applies the operation.
Returned results include column headers and go to ordinary cells on a **New sheet**
(the default), or to an **Output range** starting at a chosen cell on a chosen sheet.
SQL never silently overwrites occupied output cells or merged cells.

```sql
UPDATE Table1 SET Units = Units * 2 WHERE Region = 'East';
SELECT * FROM Table1 WHERE Region = 'East';
```

UPDATE, INSERT and DELETE edit source data. Each run, including source edits and any
returned output, is one undo/redo step. Deletes compact the dataset without deleting whole
worksheet rows. Inserts require empty space below the dataset. Unchanged cells retain
formulas and formatting; SQL-edited values become constants. Source tables with totals
or hidden headers require those options to be turned off before editing. To edit a range
that overlaps a named table, use the table instead. Merged sources can be queried but must
be unmerged before editing. Schema-changing statements such as ALTER and DROP are unsupported.

SQLite uses numbers for dates (spreadsheet serials) and booleans (1/0), NULL for blank
cells, and error-code text for spreadsheet errors. Returned data contains values, not formulas.
Source data and results are limited to one million cells each; narrow the query or use LIMIT for larger data.
Cancel closes the query and stops its worker. Queries stop after 30 seconds; cancellation,
SQL errors, collisions, and stale workbook snapshots never partially apply source changes.
Everything runs locally, offline, with no database credentials or external connections.
