# CurruptQueryAccessWorkaround
The latest workaround for the "Query is corrupt" error introduced with CVE-2019-1402 in MS Access.

### Instructions

Use the `basWorkaround.bas` module to automatically implement Microsofts suggested workaround (using a query instead of a table). As a precaution, backup your database first.

Call `AddWorkaroundForCorruptedQueryIssue()` to appöy the workaround and `RemoveWorkaroundForCorruptedQueryIssue()` to remove it at any time.

`AddWorkaroundForCorruptedQueryIssue()` will add the suffix `_Table` to all non-system tables, e.g. the table `IceCreams` would be renamed to `IceCreams_Table`.

It will also create a new query with the original table name, that will select all columns of the renamed table. In our example, the query would be named `IceCreams` and would execute the SQL `select * from [IceCreams_Table]`.

`RemoveWorkaroundForCorruptedQueryIssue()` does the reverse actions.

I tested this with all kinds of tables, including external non-MDB tables (like SQL Server). But be aware, that using a query instead of a table can lead to non-optimized queries being executed against a backend database in specific cases, especially if your original queries that used the tables are either of poor quality or very complex.

In my case I needed to manually rename `USysRibbons_Table` back to `USysRibbons`, as I hadn't marked it as as system table.

### Futher information

- [Born's Tech and Windows World](https://borncity.com/win/2019/11/13/office-november-2019-updates-are-causing-access-error-3340/?unapproved=6359&moderation-hash=597d97fc3d9abf61a4a8c4940f25bbc1#comment-6359)
- [administrator.de (German)](https://administrator.de/content/detail.php?id=514571&token=511#comment-1405456)
- [Microsoft Forum: The CVE-2019-1402 updates (KB4484119, etc.) break Access 2010/2013/2016/365: Query '' is corrupt](https://social.msdn.microsoft.com/Forums/office/en-US/7e7f24cc-f1f3-43f8-a9a2-45b77812b211/the-cve20191402-updates-kb4484119-etc-break-access-201020132016365-query-is-corrupt?forum=accessdev)
