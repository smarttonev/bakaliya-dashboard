import sys
import firebirdsql

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

cfg_mv = {
    'host': '192.168.5.254',
    'database': 'E:/db/MONTEVIDEO.FDB',
    'port': 3050,
    'user': 'SYSDBA',
    'password': 'masterkey',
    'charset': 'WIN1251'
}

conn = firebirdsql.connect(**cfg_mv)
cur = conn.cursor()

# Get current values
cur.execute("SELECT GEN_ID(G_OPERATIONS, 0), GEN_ID(G_OPERATIONSUNIQUENUM, 0) FROM RDB$DATABASE")
g_ops, g_uniq = cur.fetchone()
print(f"Before fix: G_OPERATIONS = {g_ops}, G_OPERATIONSUNIQUENUM = {g_uniq}")

# Target value: max(g_ops, g_uniq, max_id, max_uniq) + safety
cur.execute("SELECT MAX(ID), MAX(UNIQUENUM) FROM OPERATIONS WHERE LOCATIONID = 7")
max_id, max_uniq = cur.fetchone()
print(f"Max in OPERATIONS: MAX(ID) = {max_id}, MAX(UNIQUENUM) = {max_uniq}")

target = max(g_ops, g_uniq, max_id, max_uniq)
print(f"Setting both G_OPERATIONS and G_OPERATIONSUNIQUENUM to {target}...")

cur.execute(f"SET GENERATOR G_OPERATIONS TO {target}")
cur.execute(f"SET GENERATOR G_OPERATIONSUNIQUENUM TO {target}")
conn.commit()

# Verify
cur.execute("SELECT GEN_ID(G_OPERATIONS, 0), GEN_ID(G_OPERATIONSUNIQUENUM, 0) FROM RDB$DATABASE")
g_ops_new, g_uniq_new = cur.fetchone()
print(f"After fix: G_OPERATIONS = {g_ops_new}, G_OPERATIONSUNIQUENUM = {g_uniq_new}")

conn.close()
print("Success!")
