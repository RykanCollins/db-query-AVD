import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pyodbc


class PalletLabelApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Pallet Label Printer")
        self.root.geometry("450x220")

        # Warehouse Number
        tk.Label(root, text="Warehouse Number:").grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )

        self.warehouse_combo = ttk.Combobox(
            root,
            values=["01", "02"],
            width=17,
            state="readonly"
        )

        self.warehouse_combo.grid(row=0, column=1, padx=10)
        self.warehouse_combo.current(0)  # Default to 01

        # Order Number
        tk.Label(root, text="Order Number:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.order_entry = tk.Entry(root, width=20)
        self.order_entry.grid(row=1, column=1, padx=10)

        # Print Labels Button
        self.print_button = tk.Button(
            root,
            text="Print Pallet Labels",
            width=25,
            command=self.print_labels
        )
        self.print_button.grid(row=2, column=0, columnspan=2, pady=15)

        # Reset Labels Button
        self.reset_button = tk.Button(
            root,
            text="Reset Labels for Reprint",
            width=25,
            command=self.reset_labels
        )
        self.reset_button.grid(row=3, column=0, columnspan=2, pady=5)

    def get_connection(self):
        warehouse_number = self.warehouse_combo.get()

        if not warehouse_number:
            raise Exception("Warehouse Number is required.")

        conn_str = (
            "Driver={SQL Server};"
            "Server=mcs-sql01;"
            f"Database=DC{warehouse_number}MES;"
            "UID=DC00MES;"
            "PWD=DC00MES;"
        )

        return pyodbc.connect(conn_str)

    def print_labels(self):
        order_id = self.order_entry.get().strip()

        if not order_id:
            messagebox.showerror("Error", "Order Number is required.")
            return

        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()

                # Match the original SQL pattern: capture return value and
                # consume all result sets so the stored procedure runs fully.
                sql = """
                DECLARE @return_value INT;
                EXEC @return_value = dbo.BTXML_PalletLabelPrint @OrderID = ?;
                SELECT @return_value AS ReturnValue;
                """
                cursor.execute(sql, order_id)

                # Drain every result set the stored procedure emits so that
                # all label jobs are processed before we commit.
                return_value = None
                while True:
                    try:
                        rows = cursor.fetchall()
                        if rows and return_value is None:
                            # Last SELECT returns the scalar return value
                            for row in rows:
                                try:
                                    return_value = row[0]
                                except Exception:
                                    pass
                    except pyodbc.ProgrammingError:
                        pass
                    if not cursor.nextset():
                        break

                conn.commit()

                messagebox.showinfo(
                    "Success",
                    f"Labels generated successfully for Order {order_id}"
                )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def reset_labels(self):
        order_id = self.order_entry.get().strip()

        if not order_id:
            messagebox.showerror("Error", "Order Number is required.")
            return

        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()

                sql = """
                BEGIN TRAN

                UPDATE Report_Q
                SET Status = 'READY'
                WHERE MU_ID IN
                (
                    SELECT MU_ID
                    FROM Packlist_MU_SKU_Lot
                    WHERE Order_ID = ?
                )

                COMMIT
                """

                cursor.execute(sql, order_id)
                conn.commit()

                messagebox.showinfo(
                    "Success",
                    f"Labels reset to READY for Order {order_id}."
                )

        except Exception as e:
            messagebox.showerror("Database Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = PalletLabelApp(root)
    root.mainloop()
