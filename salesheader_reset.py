#! python3.11
# Reset SalesHeader/SalesDetail interface records for a given ExternalUID.
# Programmer: Rykan Collins
# Shows current errors first, then prompts for confirmation before resetting.

import configparser
import os
import pyodbc
import sys

# Reads connection details from db.ini (same directory as this script).
# db.ini format:
#   [sqlserver]
#   server   = mcs16-sql
#   database = dc00mes
#   uid      = DC00MES
#   pwd      = DC00MES
_cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "db.ini")
_cfg = configparser.ConfigParser()
_cfg.read(_cfg_path)

_sec = _cfg["sqlserver"]
CONN_STR = (
    "Driver=ODBC Driver 17 for SQL Server;"
    f"Server={_sec['server']};"
    f"Database={_sec['database']};"
    f"UID={_sec['uid']};"
    f"PWD={_sec['pwd']};"
)

QUERY_ERRORS = """
SELECT
    shi.ExternalUID,
    shi.IDSalesHeaderInterface,
    shi.InterfaceStatus        AS HeaderStatus,
    shi.InterfaceErrorText     AS HeaderErrorText,
    she.ErrorText              AS HeaderStatusErrorText,
    she.ErrorDate              AS HeaderErrorDate
FROM dbo.SalesHeader_Interface shi
LEFT JOIN dbo.SalesHeader_InterfaceError she
    ON she.IDSalesHeaderInterface = shi.IDSalesHeaderInterface
WHERE shi.ExternalUID = ?
ORDER BY she.ErrorDate DESC
"""

QUERY_DETAIL_ERRORS = """
SELECT
    sdi.IDSalesDetailInterface,
    sdi.InterfaceStatus        AS DetailStatus,
    sdi.InterfaceErrorText     AS DetailErrorText,
    sde.ErrorText              AS DetailStatusErrorText,
    sde.ErrorDate              AS DetailErrorDate
FROM dbo.SalesDetail_Interface sdi
LEFT JOIN dbo.SalesDetail_InterfaceError sde
    ON sde.IDSalesDetailInterface = sdi.IDSalesDetailInterface
WHERE sdi.IDSalesHeaderInterface IN (
    SELECT IDSalesHeaderInterface
    FROM dbo.SalesHeader_Interface
    WHERE ExternalUID = ?
)
ORDER BY sde.ErrorDate DESC
"""


def print_separator(label=""):
    width = 70
    if label:
        print(f"\n{'=' * 3} {label} {'=' * (width - len(label) - 5)}")
    else:
        print("=" * width)


def show_errors(cursor, external_uid):
    print_separator("HEADER ERRORS")
    cursor.execute(QUERY_ERRORS, external_uid)
    rows = cursor.fetchall()
    if not rows:
        print("  No header records found for ExternalUID:", external_uid)
        return False

    for row in rows:
        print(f"  ExternalUID          : {row.ExternalUID}")
        print(f"  IDSalesHeaderInterface: {row.IDSalesHeaderInterface}")
        print(f"  Header Status        : {row.HeaderStatus}")
        print(f"  Header Error Text    : {row.HeaderErrorText or '(none)'}")
        print(f"  Status Error Text    : {row.HeaderStatusErrorText or '(none)'}")
        print(f"  Error Date           : {row.HeaderErrorDate or '(none)'}")
        print()

    print_separator("DETAIL ERRORS")
    cursor.execute(QUERY_DETAIL_ERRORS, external_uid)
    detail_rows = cursor.fetchall()
    if not detail_rows:
        print("  No detail error records found.")
    else:
        for row in detail_rows:
            print(f"  IDSalesDetailInterface: {row.IDSalesDetailInterface}")
            print(f"  Detail Status         : {row.DetailStatus}")
            print(f"  Detail Error Text     : {row.DetailErrorText or '(none)'}")
            print(f"  Status Error Text     : {row.DetailStatusErrorText or '(none)'}")
            print(f"  Error Date            : {row.DetailErrorDate or '(none)'}")
            print()

    return True


def reset_interface(conn, cursor, external_uid):
    print_separator("RUNNING RESET")

    cursor.execute("""
        DELETE FROM DBO.SalesDetail_InterfaceStatus
        WHERE IDSalesDetailInterface IN (
            SELECT IDSalesDetailInterface
            FROM DBO.SalesDetail_Interface
            WHERE IDSalesHeaderInterface IN (
                SELECT IDSalesHeaderInterface
                FROM dbo.SalesHeader_Interface
                WHERE ExternalUID = ?
            )
        )
    """, external_uid)
    print(f"  Deleted {cursor.rowcount} SalesDetail_InterfaceStatus row(s).")

    cursor.execute("""
        DELETE FROM DBO.SalesDetail_InterfaceError
        WHERE IDSalesDetailInterface IN (
            SELECT IDSalesDetailInterface
            FROM DBO.SalesDetail_Interface
            WHERE IDSalesHeaderInterface IN (
                SELECT IDSalesHeaderInterface
                FROM dbo.SalesHeader_Interface
                WHERE ExternalUID = ?
            )
        )
    """, external_uid)
    print(f"  Deleted {cursor.rowcount} SalesDetail_InterfaceError row(s).")

    cursor.execute("""
        UPDATE DBO.SalesDetail_Interface
        SET InterfaceStatus  = 'NEW',
            InterfaceErrorText = '',
            orderid          = ''
        WHERE IDSalesHeaderInterface IN (
            SELECT IDSalesHeaderInterface
            FROM dbo.SalesHeader_Interface
            WHERE ExternalUID = ?
        )
    """, external_uid)
    print(f"  Updated {cursor.rowcount} SalesDetail_Interface row(s) to NEW.")

    cursor.execute("""
        DELETE FROM DBO.SalesHeader_InterfaceStatus
        WHERE IDSalesHeaderInterface IN (
            SELECT IDSalesHeaderInterface
            FROM dbo.SalesHeader_Interface
            WHERE ExternalUID = ?
        )
    """, external_uid)
    print(f"  Deleted {cursor.rowcount} SalesHeader_InterfaceStatus row(s).")

    cursor.execute("""
        DELETE FROM dbo.SalesHeader_InterfaceError
        WHERE IDSalesHeaderInterface IN (
            SELECT IDSalesHeaderInterface
            FROM dbo.SalesHeader_Interface
            WHERE ExternalUID = ?
        )
    """, external_uid)
    print(f"  Deleted {cursor.rowcount} SalesHeader_InterfaceError row(s).")

    cursor.execute("""
        UPDATE DBO.SalesHeader_Interface
        SET InterfaceStatus  = 'NEW',
            InterfaceErrorText = '',
            orderid          = ''
        WHERE ExternalUID = ?
    """, external_uid)
    print(f"  Updated {cursor.rowcount} SalesHeader_Interface row(s) to NEW.")

    conn.commit()
    print("\n  COMMIT successful — interface records reset to NEW.")


def main():
    external_uid = input("Enter ExternalUID to investigate: ").strip()
    if not external_uid:
        print("No ExternalUID entered. Exiting.")
        sys.exit(1)

    try:
        conn = pyodbc.connect(CONN_STR)
    except pyodbc.Error as e:
        print("Database connection failed:", e)
        sys.exit(1)

    cursor = conn.cursor()

    found = show_errors(cursor, external_uid)
    if not found:
        conn.close()
        sys.exit(0)

    print_separator()
    confirm = input(
        f"\nReset all interface records for ExternalUID '{external_uid}' to NEW? (yes/no): "
    ).strip().lower()

    if confirm == "yes":
        try:
            reset_interface(conn, cursor, external_uid)
        except pyodbc.Error as e:
            conn.rollback()
            print("\nError during reset — transaction rolled back:", e)
    else:
        print("Reset cancelled. No changes made.")

    conn.close()


if __name__ == "__main__":
    main()
