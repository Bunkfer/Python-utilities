import pandas as pd
import glob
import os


class ReporteConsolidado:
    def __init__(self, carpeta, archivo_base):
        self.carpeta = carpeta
        self.archivo_base = archivo_base
        self.archivos = self._obtener_archivos()
        self.df_final = None
        self.equivalencias = None
        self.acreedores = None

    def _obtener_archivos(self):
        archivos = glob.glob(os.path.join(self.carpeta, "*.csv"))
        return [f for f in archivos if not f.endswith("Consolidado.csv")]

    def _leer_csv(self, archivo):
        df_temp = pd.read_csv(archivo, header=None, encoding="latin1")
        numero = df_temp.iloc[0, 0]

        header_row = df_temp.index[
            df_temp.apply(
                lambda r: r.astype(str).str.contains("NumNomina").any(), axis=1
            )
        ][0]

        df = pd.read_csv(archivo, header=header_row, encoding="latin1")
        df["Archivo"] = os.path.basename(archivo)
        df["Fondo"] = numero
        return df

    def consolidar_csvs(self):
        dataframes = [self._leer_csv(archivo) for archivo in self.archivos]
        self.df_final = pd.concat(dataframes, ignore_index=True)
        self.df_final["Fondo"] = self.df_final["Fondo"].astype(int)

    def cargar_equivalencias_y_acreedores(self):
        self.equivalencias = pd.read_excel(
            self.archivo_base, sheet_name="Equivalencias"
        )
        self.acreedores = pd.read_excel(self.archivo_base, sheet_name="Acredores")

    def _escribir_formula(self, worksheet, col_idx, titulo, formula_fn, last_row):
        worksheet.write(0, col_idx, titulo)
        for row in range(1, last_row):
            worksheet.write_formula(row, col_idx, formula_fn(row))

    def generar_excel(self):
        salida = os.path.join(self.carpeta, "Consolidado.xlsx")
        with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
            # Guardar hoja principal y auxiliares
            self.df_final.to_excel(writer, sheet_name="Consolidado", index=False)
            self.equivalencias.to_excel(writer, sheet_name="Equivalencias", index=False)
            self.acreedores.to_excel(writer, sheet_name="Acredores", index=False)

            workbook = writer.book
            worksheet = writer.sheets["Consolidado"]

            fondo_col = self.df_final.columns.get_loc("Fondo")
            last_row = len(self.df_final) + 1
            col_formula = len(self.df_final.columns)

            # === 1. UEN SAP ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "UEN SAP",
                lambda row: f"=VLOOKUP({chr(65+fondo_col)}{row+1},Equivalencias!A:F,6,0)",
                last_row,
            )
            self.df_final = self.df_final.merge(
                self.equivalencias.iloc[:, [0, 5]],
                left_on="Fondo",
                right_on=self.equivalencias.columns[0],
                how="left",
            ).rename(columns={self.equivalencias.columns[5]: "UEN SAP"})
            self.df_final.drop(columns=[self.equivalencias.columns[0]], inplace=True)
            col_formula += 1

            # === 2. Nombre de la sociedad ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "Nombre de la sociedad",
                lambda row: f"=VLOOKUP({chr(65+fondo_col)}{row+1},Equivalencias!A:C,3,0)",
                last_row,
            )
            col_formula += 1

            # === 3. SOC + DECADA ===
            n_col = self.df_final.columns.get_loc("UEN SAP")
            l_col = self.df_final.columns.get_loc("Codigo Emi")
            self._escribir_formula(
                worksheet,
                col_formula,
                "SOC + DECADA",
                lambda row: f"=CONCATENATE({chr(65+n_col)}{row+1},{chr(65+l_col)}{row+1})",
                last_row,
            )
            self.df_final["UEN SAP"] = self.df_final["UEN SAP"].astype(str)
            self.df_final["Codigo Emi"] = self.df_final["Codigo Emi"].astype(str)
            self.df_final["SOC + DECADA"] = (
                self.df_final["UEN SAP"] + self.df_final["Codigo Emi"]
            )
            col_formula += 1

            # === 4. ACREEDOR ===
            SOC_DECADA_col = self.df_final.columns.get_loc("SOC + DECADA") + 1
            self._escribir_formula(
                worksheet,
                col_formula,
                "ACREEDOR",
                lambda row: f"=VLOOKUP(VALUE({chr(65+SOC_DECADA_col)}{row+1}),Acredores!D:G,4,0)",
                last_row,
            )
            col_formula += 1

            # === 5. Cta mayor ===
            Codigo_Cue_col = self.df_final.columns.get_loc("Codigo Cue")
            self._escribir_formula(
                worksheet,
                col_formula,
                "Cta mayor",
                lambda row: f"=VLOOKUP({chr(65+Codigo_Cue_col)}{row+1},Equivalencias!I:K,3,0)",
                last_row,
            )
            col_formula += 1

            # === 6. REF 1 ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "REF 1",
                lambda row: f'=CONCATENATE({chr(65+n_col)}{row+1},"Q"," ","PP142025")',
                last_row,
            )
            col_formula += 1

            # === 7. Ref Acreed ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "Ref Acreed",
                lambda row: f"=VLOOKUP({chr(65+Codigo_Cue_col)}{row+1},Equivalencias!I:J,2,0)",
                last_row,
            )
            ref_acreed = self.equivalencias.iloc[:, [8, 9]].copy()
            ref_acreed.columns = ["Codigo Cue", "Ref Acreed"]
            self.df_final["Codigo Cue"] = (
                self.df_final["Codigo Cue"].astype(str).str.strip()
            )
            ref_acreed["Codigo Cue"] = ref_acreed["Codigo Cue"].astype(str).str.strip()
            self.df_final = self.df_final.merge(ref_acreed, on="Codigo Cue", how="left")
            col_formula += 1

            # === 8. texto largo ===
            Ref_Acreed_col = self.df_final.columns.get_loc("Ref Acreed") + 4
            self._escribir_formula(
                worksheet,
                col_formula,
                "texto largo",
                lambda row: f"=VLOOKUP({chr(65+Ref_Acreed_col)}{row+1},Equivalencias!J:L,3,0)",
                last_row,
            )
            col_formula += 1

            # === 9. Soc ===
            worksheet.write(0, col_formula, "Soc")
            for row in range(1, last_row):
                worksheet.write_string(row, col_formula, "CLAV CONCEPT NOM")
            self.df_final["Soc"] = "CLAV CONCEPT NOM"
            col_formula += 1

            # === 10. Clav ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "Clav",
                lambda row: f"=VLOOKUP({chr(65+Ref_Acreed_col)}{row+1},Equivalencias!J:M,4,0)",
                last_row,
            )
            mapa_clav = self.equivalencias.iloc[:, [9, 12]]
            mapa_clav = mapa_clav.rename(
                columns={
                    mapa_clav.columns[0]: "Ref Acreed",
                    mapa_clav.columns[1]: "Clav",
                }
            )
            self.df_final = self.df_final.merge(mapa_clav, on="Ref Acreed", how="left")
            col_formula += 1

            # === 11. AP ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "AP",
                lambda row: f'=CONCATENATE({chr(65+n_col)}{row+1},"Q"," ","014"," ","2025PP")',
                last_row,
            )
            self.df_final["AP"] = (
                self.df_final.iloc[:, n_col].astype(str) + "Q 014 2025PP"
            )
            col_formula += 1

            # === 12. Q08 ===
            Soc_col = self.df_final.columns.get_loc("Soc") + 5
            Clav_col = self.df_final.columns.get_loc("Clav") + 5
            AP_col = self.df_final.columns.get_loc("AP") + 5
            self._escribir_formula(
                worksheet,
                col_formula,
                "Q08",
                lambda row: f'=CONCATENATE({chr(65+n_col)}{row+1}," ",{chr(65+Soc_col)}{row+1}," ",{chr(65+Clav_col)}{row+1}," ",{chr(65+AP_col)}{row+1})',
                last_row,
            )
            col_formula += 1

            # === 13. Record ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "Record",
                lambda row: f"=VLOOKUP(VALUE({chr(65+SOC_DECADA_col)}{row+1}),Acredores!D:H,5,0)",
                last_row,
            )
            col_formula += 1

            # === 14. contrato accival ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "contrato accival",
                lambda row: f"=VLOOKUP({chr(65+n_col)}{row+1},Acredores!A:J,9,0)",
                last_row,
            )
            col_formula += 1

            # === 15. Subcuneta Citi ===
            self._escribir_formula(
                worksheet,
                col_formula,
                "Subcuneta Citi",
                lambda row: f"=VLOOKUP({chr(65+n_col)}{row+1},Acredores!A:J,10,0)",
                last_row,
            )
