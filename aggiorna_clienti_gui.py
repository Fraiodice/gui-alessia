"""
Archivos Actualizados Automaticamente
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os
import shutil



class AppAggiornaClienti:

    COLONNA_ID = "ID"

    def __init__(self, root):
        self.root = root
        self.root.title("Archivos Actualizados Automaticamente")
        self.root.geometry("600x400")
        self.root.resizable(False, False)

        self.file_vecchio = tk.StringVar(value="Ningun archivo seleccionado")
        self.file_nuevo = tk.StringVar(value="Ningun archivo seleccionado")
        self.file_output = None

        self._build_ui()

    def _build_ui(self):
        # Titulo
        tk.Label(
            self.root, text="Archivos Actualizados Automaticamente",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(16, 10))

        tk.Frame(self.root, height=1, bg="#c0c0c0").pack(fill="x", padx=14)

        # Archivo Viejo
        frame1 = tk.LabelFrame(self.root, text="Archivo Viejo", font=("Segoe UI", 10), padx=10, pady=10)
        frame1.pack(fill="x", padx=14, pady=6)

        row1 = tk.Frame(frame1)
        row1.pack(fill="x")
        tk.Label(row1, textvariable=self.file_vecchio, font=("Segoe UI", 9), fg="#444", anchor="w").pack(side="left", fill="x", expand=True)
        tk.Button(row1, text="Buscar...", width=12, font=("Segoe UI", 9), fg="#000000", command=lambda: self._elegir_archivo(self.file_vecchio)).pack(side="right", padx=(8, 0))

        # Archivo Nuevo
        frame2 = tk.LabelFrame(self.root, text="Archivo Nuevo", font=("Segoe UI", 10), padx=10, pady=10)
        frame2.pack(fill="x", padx=14, pady=6)

        row2 = tk.Frame(frame2)
        row2.pack(fill="x")
        tk.Label(row2, textvariable=self.file_nuevo, font=("Segoe UI", 9), fg="#444", anchor="w").pack(side="left", fill="x", expand=True)
        tk.Button(row2, text="Buscar...", width=12, font=("Segoe UI", 9), fg="#000000", command=lambda: self._elegir_archivo(self.file_nuevo)).pack(side="right", padx=(8, 0))

        tk.Frame(self.root, height=1, bg="#c0c0c0").pack(fill="x", padx=14, pady=(10, 0))

        # Botones
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=14)

        self.btn_actualizar = tk.Button(
            btn_frame, text="Actualizar Archivo", width=22, padx=10, pady=6,
            font=("Segoe UI", 11), fg="#000000", command=self._ejecutar_merge
        )
        self.btn_actualizar.pack(side="left", padx=8)

        self.btn_descargar = tk.Button(
            btn_frame, text="Descargar al Escritorio", width=22, padx=10, pady=6,
            font=("Segoe UI", 11), fg="#000000", state="disabled", command=self._guardar_escritorio
        )
        self.btn_descargar.pack(side="left", padx=8)


    def _elegir_archivo(self, var):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if path:
            var.set(path)

    def _ejecutar_merge(self):
        f_viejo = self.file_vecchio.get()
        f_nuevo = self.file_nuevo.get()

        if "Ningun archivo" in f_viejo or "Ningun archivo" in f_nuevo:
            messagebox.showwarning("Atencion", "Selecciona ambos archivos antes de continuar.")
            return

        try:
            df_viejo = pd.read_excel(f_viejo, dtype={self.COLONNA_ID: str})
            df_nuevo = pd.read_excel(f_nuevo, dtype={self.COLONNA_ID: str})

            for etiqueta, df in [("viejo", df_viejo), ("nuevo", df_nuevo)]:
                if self.COLONNA_ID not in df.columns:
                    messagebox.showerror(
                        "Error",
                        f"Columna '{self.COLONNA_ID}' no encontrada en el archivo {etiqueta}.\n\nColumnas disponibles: {', '.join(df.columns)}"
                    )
                    return

            ids_existentes = set(df_viejo[self.COLONNA_ID].astype(str))
            mask = ~df_nuevo[self.COLONNA_ID].astype(str).isin(ids_existentes)
            df_agregar = df_nuevo[mask]
            duplicados = mask.eq(False).sum()
            df_resultado = pd.concat([df_viejo, df_agregar], ignore_index=True)

            hoy = datetime.now().strftime("%Y-%m-%d")
            temp_path = os.path.join(os.environ.get("TEMP", "/tmp"), f"archivo_actualizado_{hoy}.xlsx")
            df_resultado.to_excel(temp_path, index=False)
            self.file_output = temp_path

            self.btn_descargar.config(state="normal")

            messagebox.showinfo(
                "Completado",
                f"Registros existentes: {len(df_viejo)}\n"
                f"Registros nuevos encontrados: {len(df_nuevo)}\n"
                f"Duplicados descartados: {duplicados}\n"
                f"Nuevos agregados: {len(df_agregar)}\n"
                f"Total final: {len(df_resultado)}\n\n"
                f"Pulsa 'Descargar al Escritorio' para guardar."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _guardar_escritorio(self):
        if not self.file_output or not os.path.exists(self.file_output):
            messagebox.showerror("Error", "No hay archivo para guardar. Ejecuta primero la actualizacion.")
            return

        escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        if not os.path.isdir(escritorio):
            escritorio = os.path.join(os.path.expanduser("~"), "Escritorio")
        if not os.path.isdir(escritorio):
            escritorio = os.path.expanduser("~")

        hoy = datetime.now().strftime("%Y-%m-%d")
        dest = os.path.join(escritorio, f"archivo_actualizado_{hoy}.xlsx")

        n = 1
        while os.path.exists(dest):
            dest = os.path.join(escritorio, f"archivo_actualizado_{hoy}_{n}.xlsx")
            n += 1

        try:
            shutil.copy2(self.file_output, dest)
            messagebox.showinfo("Guardado", f"Archivo guardado en el Escritorio:\n{os.path.basename(dest)}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    AppAggiornaClienti(root)
    root.mainloop()
