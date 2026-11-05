# Tests

Scripts that need a database with `as_xlsx` installed and a directory object `OUT_DIR`.

| Script | What it checks |
|---|---|
| `demo_styles.sql` | Named styles, `use_template`, `define_style`, styled `query2sheet` with zebra rows, total row and formulas. Writes `demo_styles.xlsx`. |
| `usage_examples.sql` | All examples from `usage.md`. Run it with the old and the new package and compare the files to make sure existing calls are unchanged. |

## Local test database with Docker

```sh
docker run -d --name ora -e ORACLE_PASSWORD=secret -e APP_USER=xl -e APP_USER_PASSWORD=xl \
  -v "$PWD":/src:ro -v /tmp/xlsx_out:/out gvenzl/oracle-free:23-slim-faststart
# as sysdba in FREEPDB1:
#   create directory OUT_DIR as '/out';
#   grant read, write on directory OUT_DIR to xl;
#   grant execute on utl_file to xl;
docker exec ora sqlplus xl/xl@localhost/FREEPDB1 @/src/src/as_xlsx.pks
docker exec ora sqlplus xl/xl@localhost/FREEPDB1 @/src/src/as_xlsx.pkb
docker exec -w /src/tests ora sqlplus xl/xl@localhost/FREEPDB1 @demo_styles.sql
```
