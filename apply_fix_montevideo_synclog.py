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

first_item_log_ids = [
    317409, 317412, 317415, 317420, 317425,
    317430, 317433, 317438, 317443, 317448,
    317453, 317455, 317457, 317459, 317461,
    317463, 317465, 317467, 317469, 317471
]

print(f"Connecting to MONTEVIDEO.FDB to update {len(first_item_log_ids)} records in SYNCHRONIZATIONLOG...")
conn = firebirdsql.connect(**cfg_mv)
cur = conn.cursor()

updated_count = 0
for lid in first_item_log_ids:
    cur.execute("SELECT SYNCHROLOG FROM SYNCHRONIZATIONLOG WHERE ID = ?", (lid,))
    row = cur.fetchone()
    if not row:
        print(f"Warning: LogId {lid} not found!")
        continue
    
    txt = row[0]
    if hasattr(txt, 'read'): txt = txt.read()
    if isinstance(txt, bytes): txt = txt.decode('windows-1251')
    
    if '\x1dF\x1d0\x1d' not in txt:
        if '\x1dF\x1d1\x1d' in txt:
            print(f"  LogId {lid} already has F=1")
            updated_count += 1
        else:
            print(f"Warning: LogId {lid} does not have F=0 or F=1!")
        continue
    
    new_txt = txt.replace('\x1dF\x1d0\x1d', '\x1dF\x1d1\x1d')
    cur.execute("UPDATE SYNCHRONIZATIONLOG SET SYNCHROLOG = ? WHERE ID = ?", (new_txt.encode('windows-1251'), lid))
    updated_count += 1
    print(f"  LogId {lid}: updated F from 0 to 1")

conn.commit()
print(f"\nSuccessfully committed updates for {updated_count} records in MONTEVIDEO.FDB!")

# Verification
print("\nVerifying updated records...")
all_ok = True
for lid in first_item_log_ids:
    cur.execute("SELECT SYNCHROLOG FROM SYNCHRONIZATIONLOG WHERE ID = ?", (lid,))
    txt = cur.fetchone()[0]
    if hasattr(txt, 'read'): txt = txt.read()
    if isinstance(txt, bytes): txt = txt.decode('windows-1251')
    if '\x1dF\x1d1\x1d' not in txt:
        print(f"  ERROR: LogId {lid} does not have F=1!")
        all_ok = False
    else:
        print(f"  LogId {lid} verified (F=1)")

conn.close()

if all_ok:
    print("\nALL 20 RECORDS VERIFIED WITH F=1!")
else:
    print("\nSOME RECORDS FAILED VERIFICATION!")
