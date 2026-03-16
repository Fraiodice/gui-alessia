"""
Archivos Actualizados Automaticamente
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os
import shutil
import math


class PizzaAnimada(tk.Canvas):
    """Pizza margherita fumante animada."""

    def __init__(self, parent, size=80, **kwargs):
        super().__init__(parent, width=size, height=size + 20, highlightthickness=0, **kwargs)
        self.size = size
        self.fase = 0.0
        self.animando = False

    def iniciar(self):
        self.animando = True
        self._animar()

    def detener(self):
        self.animando = False

    def _animar(self):
        if not self.animando:
            return
        self.delete("all")
        cx = self.size // 2
        cy = self.size // 2 + 18
        r = self.size // 2 - 4

        # Plato
        self.create_oval(cx - r - 4, cy - r // 3 - 2, cx + r + 4, cy + r // 3 + 8, fill="#d4d4d4", outline="#bbb")

        # Base pizza (vista dall'alto, ovale per prospettiva)
        self.create_oval(cx - r, cy - r // 2, cx + r, cy + r // 2, fill="#e8a838", outline="#c4872a", width=2)

        # Salsa pomodoro
        self.create_oval(cx - r + 8, cy - r // 2 + 6, cx + r - 8, cy + r // 2 - 6, fill="#d63a2a", outline="")

        # Mozzarella (chiazze bianche)
        manchas = [(-14, -6), (10, -4), (-4, 4), (16, 2), (-18, 2), (6, -10), (-8, -12)]
        for mx, my in manchas:
            self.create_oval(cx + mx - 5, cy + my - 3, cx + mx + 5, cy + my + 3, fill="#fff8e1", outline="")

        # Basilico (foglioline verdi)
        hojas = [(-10, -2), (8, 4), (0, -8), (14, -4), (-6, 6)]
        for hx, hy in hojas:
            self.create_oval(cx + hx - 3, cy + hy - 2, cx + hx + 3, cy + hy + 2, fill="#2e7d32", outline="")

        # Vapore animado (3 fili di fumo)
        for i, offset_x in enumerate([-12, 0, 12]):
            fase_i = self.fase + i * 1.2
            for j in range(3):
                y_pos = cy - r // 2 - 8 - j * 8
                x_wave = math.sin(fase_i + j * 0.8) * 4
                alpha_sim = max(0.2, 1.0 - j * 0.35)
                gray = int(180 + (1 - alpha_sim) * 75)
                color = f"#{gray:02x}{gray:02x}{gray:02x}"
                self.create_line(
                    cx + offset_x + x_wave, y_pos + 4,
                    cx + offset_x - x_wave * 0.5, y_pos - 4,
                    fill=color, width=2, smooth=True
                )

        self.fase += 0.12
        self.after(60, self._animar)


class AppAggiornaClienti:

    COLONNA_ID = "Idusuario"

    def __init__(self, root):
        self.root = root
        self.root.title("Archivos Actualizados Automaticamente")
        self.root.geometry("600x420")
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
            btn_frame, text="Guardar Archivo", width=22, padx=10, pady=6,
            font=("Segoe UI", 11), fg="#000000", state="disabled", command=self._guardar_archivo
        )
        self.btn_descargar.pack(side="left", padx=8)

        # Pizza frame (oculto hasta despues del merge)
        self.pizza_frame = tk.Frame(self.root)
        self.pizza_frame.pack(fill="x", padx=20, pady=(4, 10), side="bottom")

        self.pizza = PizzaAnimada(
            self.pizza_frame, size=80,
            bg=self.root.cget("bg")
        )
        self.pizza.pack(side="left", padx=(10, 10))

        self.lbl_pizza = tk.Label(
            self.pizza_frame,
            text="Ahora puedes invitarme\na una pizza margherita!",
            font=("Segoe UI", 11, "italic"), fg="#c44a1a",
            justify="left"
        )
        self.lbl_pizza.pack(side="left", padx=(4, 0))

        self.pizza_frame.pack_forget()

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
                f"Pulsa 'Guardar Archivo' para guardar."
            )

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _guardar_archivo(self):
        if not self.file_output or not os.path.exists(self.file_output):
            messagebox.showerror("Error", "No hay archivo para guardar. Ejecuta primero la actualizacion.")
            return

        # Mostrar pizza animada
        self.pizza_frame.pack(fill="x", padx=20, pady=(4, 10), side="bottom")
        self.pizza.iniciar()
        self.root.update()

        # Esperar 2 segundos para que se vea la pizza, luego guardar
        self.root.after(2000, self._hacer_descarga)

    def _hacer_descarga(self):
        hoy = datetime.now().strftime("%Y-%m-%d")
        default_name = f"archivo_actualizado_{hoy}.xlsx"

        dest = filedialog.asksaveasfilename(
            title="Guardar archivo actualizado",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Archivos Excel", "*.xlsx")]
        )

        if not dest:
            return

        try:
            shutil.copy2(self.file_output, dest)
            messagebox.showinfo("Guardado", f"Archivo guardado correctamente:\n{os.path.basename(dest)}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    AppAggiornaClienti(root)
    root.mainloop()
