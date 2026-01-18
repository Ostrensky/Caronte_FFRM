# --- FILE: app/auto_text_dialog.py ---

import pandas as pd
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTextEdit,
                               QPushButton, QMessageBox, QDialogButtonBox)
from PySide6.QtCore import Qt
from pandas.api.types import is_datetime64_any_dtype 

# --- Local Application Imports ---
from .constants import Columns
from document_parts import format_invoice_numbers


class AutoTextDialog(QDialog):
    """
    A dialog for viewing, editing, and saving the detailed motive text
    for a specific Auto de Infração or IDD.
    """
    # ✅ FIX: Added 'idd_mode' to signature with default False
    def __init__(self, auto_data, df_invoices, pgdas_payments_map, calculated_auto_data, parent=None, idd_mode=False):
        super().__init__(parent)
        
        self.idd_mode = idd_mode
        self.term_label = "IDD" if self.idd_mode else "Auto de Infração"
        self.term_short = "IDD" if self.idd_mode else "Auto"

        self.setWindowTitle(f"Editar Texto para: {auto_data.get('numero', self.term_label)}")
        self.setMinimumSize(800, 600)
        self.setWindowModality(Qt.WindowModality.WindowModal)

        self.auto_data = auto_data
        self.df_invoices = df_invoices # This is the *original* unfiltered DF
        self.pgdas_payments_map = pgdas_payments_map
        self.calculated_auto_data = calculated_auto_data
        
        # Main text edit area
        self.text_edit = QTextEdit()
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        
        self.regenerate_btn = QPushButton("Regerar Texto (Perderá Edições)")
        self.regenerate_btn.clicked.connect(self.populate_text)
        
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.regenerate_btn)
        button_layout.addStretch()
        button_layout.addWidget(button_box)

        # Layout
        layout = QVBoxLayout(self)
        layout.addWidget(self.text_edit)
        layout.addLayout(button_layout)
        
        # Populate text
        existing_text = self.auto_data.get('auto_text')
        if existing_text:
            self.text_edit.setPlainText(existing_text)
        else:
            self.populate_text()

    def populate_text(self):
        """Generates and sets the text based on the auto data."""
        try:
            generated_text = self._generate_auto_text()
            self.text_edit.setPlainText(generated_text)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Gerar Texto",
                                 f"Ocorreu um erro ao gerar o texto: {e}")
            self.text_edit.setPlainText(f"ERRO AO GERAR TEXTO: {e}")

    def get_text(self):
        """Returns the final, potentially edited text."""
        return self.text_edit.toPlainText()

    def _generate_auto_text(self):
        """
        Generates the detailed motive text based on the user's template.
        """
        
        # --- PART I: Introduction & Invoice List ---
        part_I = ""
        
        # ✅ NEW FILTERING LOGIC
        df_invoices_para_listar = self.df_invoices 
        dados_anuais = self.calculated_auto_data.get('dados_anuais', [])
        meses_compensados_set = set()

        if (self.df_invoices is not None and not self.df_invoices.empty and 
            'DATA EMISSÃO' in self.df_invoices.columns):
            
            df_temp_for_filtering = self.df_invoices.copy()
            if not is_datetime64_any_dtype(df_temp_for_filtering['DATA EMISSÃO']):
                df_temp_for_filtering['DATA EMISSÃO'] = pd.to_datetime(df_temp_for_filtering['DATA EMISSÃO'], errors='coerce')

            for mes_data in dados_anuais:
                if mes_data.get('iss_apurado_op', 0.0) <= 0.01: 
                    meses_compensados_set.add(mes_data['mes_ano']) 

            if meses_compensados_set:
                df_temp_for_filtering['mes_ano_str'] = df_temp_for_filtering['DATA EMISSÃO'].dt.strftime('%m/%Y')
                df_temp_for_filtering['mes_ano_str'].fillna('INVALID_DATE', inplace=True)
                invoices_removidas_mask = df_temp_for_filtering['mes_ano_str'].isin(meses_compensados_set)
                df_invoices_para_listar = df_temp_for_filtering[~invoices_removidas_mask]
        
        # ✅ Use dynamic label
        label_entity = "contribuinte" if self.idd_mode else "sujeito passivo"
        
        if df_invoices_para_listar is None or df_invoices_para_listar.empty:
            part_I = f"O {label_entity} acima qualificado não possui notas fiscais associadas a este {self.term_short} (ou todas as notas foram compensadas)."
        else:
            activity_code_col = getattr(Columns, 'ACTIVITY_CODE', 'CÓDIGO SERVIÇO')
            activity_desc_col = Columns.ACTIVITY_DESC
            
            if activity_code_col not in df_invoices_para_listar.columns:
                part_I += f"[AVISO: Coluna '{activity_code_col}' não encontrada.]\n"
            if activity_desc_col not in df_invoices_para_listar.columns:
                 part_I += f"[AVISO: Coluna '{activity_desc_col}' não encontrada.]\n"

            service_groups = {}
            if activity_code_col in df_invoices_para_listar.columns and activity_desc_col in df_invoices_para_listar.columns:
                df_safe = df_invoices_para_listar.fillna({activity_code_col: 'N/A', activity_desc_col: 'N/A'})
                for (code, desc), group in df_safe.groupby([activity_code_col, activity_desc_col]):
                    key = (str(code).strip(), str(desc).strip())
                    if key not in service_groups:
                        service_groups[key] = []
                    service_groups[key].extend(group[Columns.INVOICE_NUMBER].astype(str).tolist())
            
            num_services = len(service_groups)
            
            intro_base = f"O {label_entity} acima qualificado deixou de recolher integralmente o Imposto Sobre Serviços (ISS) devido"
            
            if num_services == 0:
                all_invoice_nums = df_invoices_para_listar[Columns.INVOICE_NUMBER].astype(str).tolist()
                nfs_e_numeros = format_invoice_numbers(all_invoice_nums)
                part_I = (
                    f"{intro_base}, descritos nas Notas Fiscais de Serviços Eletrônica (NFS-e) de nº (s) {nfs_e_numeros}, "
                    "levantados em procedimento administrativo fiscal (monitoramento)."
                )
            elif num_services == 1:
                (code, desc) = list(service_groups.keys())[0]
                invoice_nums = service_groups[(code, desc)]
                nfs_e_numeros = format_invoice_numbers(invoice_nums)
                plural_s = "" if len(invoice_nums) == 1 else "(s)"
                nota_text = "Nota Fiscal de Serviço Eletrônica (NFS-e)" if len(invoice_nums) == 1 else "Notas Fiscais de Serviços Eletrônica (NFS-e)"

                part_I = (
                    f"{intro_base} relativo à prestação de serviços de {desc.lower()},\nenquadrados no código de atividade {code} da Lista Anexa à Lei Complementar nº 40/2001, "
                    f"descritos na {nota_text} de nº{plural_s} {nfs_e_numeros}, "
                    "levantados em procedimento administrativo fiscal (monitoramento)."
                )
            else:
                part_I = (
                    f"{intro_base}, levantados em procedimento administrativo fiscal (monitoramento), "
                    "relativo à prestação dos seguintes serviços:\n"
                )
                service_lines = []
                for (code, desc), invoice_nums in sorted(service_groups.items()):
                    nfs_e_numeros = format_invoice_numbers(invoice_nums)
                    service_lines.append(
                        f"- Serviços de {desc.capitalize()},\nenquadrados no código de atividade {code}, "
                        f"descritos nas Notas Fiscais de Serviços Eletrônica (NFS-e) de nº(s) {nfs_e_numeros};"
                    )
                part_I += "\n".join(service_lines)

            if self.pgdas_payments_map:
                if self.df_invoices is not None and not self.df_invoices.empty and 'DATA EMISSÃO' in self.df_invoices.columns:
                    if not is_datetime64_any_dtype(self.df_invoices['DATA EMISSÃO']):
                        self.df_invoices['DATA EMISSÃO'] = pd.to_datetime(self.df_invoices['DATA EMISSÃO'], errors='coerce')

                    invoice_periods = self.df_invoices['DATA EMISSÃO'].dropna().dt.to_period('M').astype(str).unique()
                    das_periods_found = []
                    for p_obj in invoice_periods:
                        try:
                            year, month = p_obj.split('-')
                            map_key = f"{int(month):02d}/{year}"
                            if map_key in self.pgdas_payments_map:
                                das_periods_found.append(map_key)
                        except Exception:
                            pass 
                    
                    if das_periods_found:
                        das_periods_str = ", ".join(sorted(list(set(das_periods_found))))
                        part_I += f"\n\nHouve pagamentos de DAS referente ao(s) período(s) de apuração {das_periods_str}. Os valores foram descontados do ISS devido."
        
        # --- PART II: Detailed Motive ---
        part_II = ""
        rule_name = self.auto_data.get('rule_name', 'desconhecido')
        
        correct_aliquota = self.auto_data.get('user_defined_aliquota')
        if correct_aliquota is None and self.df_invoices is not None and not self.df_invoices.empty:
            correct_aliquota = self.df_invoices.iloc[0].get(Columns.CORRECT_RATE)
        if correct_aliquota is None:
            correct_aliquota = 5.0 
        
        aliquota_str = f"{correct_aliquota:.2f}%".replace('.', ',')
        
        declared_rate_str = "[Alíquota Declarada N/A]"
        if self.df_invoices is not None and not self.df_invoices.empty:
            declared_rate = self.df_invoices.iloc[0].get(Columns.RATE)
            if declared_rate is not None:
                declared_rate_str = f"{declared_rate:.1f}%".replace('.', ',')

        codes_str = "[código N/A]"
        activity_code_col = getattr(Columns, 'ACTIVITY_CODE', 'CÓDIGO SERVIÇO')
        if (activity_code_col in df_invoices_para_listar.columns and 
            df_invoices_para_listar is not None and not df_invoices_para_listar.empty):
            unique_codes = df_invoices_para_listar[activity_code_col].astype(str).unique()
            codes_str = ", ".join(unique_codes)
        
        all_invoice_nums_str = "[N/A]"
        if df_invoices_para_listar is not None and not df_invoices_para_listar.empty:
            all_invoice_nums = df_invoices_para_listar[Columns.INVOICE_NUMBER].astype(str).tolist()
            all_invoice_nums_str = format_invoice_numbers(all_invoice_nums)
        
        auto_numero = self.auto_data.get('numero', f'[Nº {self.term_short} N/A]')
        
        # --- Text Templates (Adapted for IDD if needed) ---

        if rule_name == 'idd_nao_pago':
             part_II = (
                "A constituição do crédito tributário decorreu do não recolhimento integral do ISS devido, "
                "apurado com base na alíquota correta para os serviços prestados.\n\n"
                "De acordo com o art. 4º da Lei Complementar Municipal "
                f"40/2001 e alterações, a alíquota para os serviços prestados é de {aliquota_str}."
            )

        elif rule_name == 'regime_incorreto':
            part_II = (
                "A constituição do crédito tributário decorreu de declaração "
                "indevida do regime tributário informado nas NFS-e mencionadas como Simples\n"
                "Nacional, vez que o sujeito passivo não é optante ao regime diferenciado no "
                "período autuado.\n\n"
                "De acordo com o art. 4º da Lei Complementar Municipal "
                f"40/2001 e alterações, a alíquota para os serviços prestados é de {aliquota_str}."
            )
        
        elif rule_name in ['diferenca_aliquota', 'aliquota_incorreta'] or self.auto_data.get('motive').startswith('Diferença Alíquota') or self.auto_data.get('motive').startswith('Alíquota Incorreta'):
            part_II = (
                "A constituição do crédito tributário decorreu de diferença "
                "de alíquota, tendo em vista a declaração indevida nas NFS-e mencionadas no "
                f"percentual de {declared_rate_str}.\n\n"
                "De acordo com o art. 4º da Lei "
                f"Complementar Municipal 40/2001 e alterações, a alíquota para os serviços "
                f"prestados é de {aliquota_str}."
            )

        elif rule_name == 'local_incidencia_incorreto':
            part_II = (
                "A constituição do crédito tributário decorreu de declaração "
                "indevida do local da incidência do ISS nas NFS-e mencionadas, visto que os serviços nelas descritos\n"
                f"e enquadrados nos códigos de atividade {codes_str} possuem tributação no município do "
                "estabelecimento do prestador e não permitem o deslocamento do aspecto espacial\n"
                "da exigibilidade, de acordo com a regra estabelecida no art. 3º da Lei "
                "Complementar Federal 116/2003.\n\n"
                "De acordo com o art. 4º da Lei Complementar Municipal "
                f"40/2001 e alterações, a alíquota para os serviços prestados é de {aliquota_str}."
            )

        elif rule_name == 'isencao_imunidade_indevida':
            # ✅ FIX: Use dynamic auto_numero with correct label
            part_II = (
                "A constituição do crédito tributário decorreu de declaração "
                "indevida de benefício fiscal (isenção e/ou imunidade) quando da\n"
                "emissão das NFS-e mencionadas.\n\n"
                "De acordo com o art. 4º da Lei Complementar Municipal "
                f"40/2001 e alterações, a alíquota para os serviços prestados é de {aliquota_str}.\n\n"
                f"Os serviços prestados conforme NFS-e de nº {all_invoice_nums_str} não são isentos do recolhimento de ISS, uma vez\n"
                "que não satisfazem as condições previstas no art. 85 da Lei Complementar "
                f"40/2001. Dessa forma, o ISS foi constituído por meio do {self.term_label} nº "
                f"{auto_numero}.\n\n"
                "A constituição do crédito tributário decorreu de declaração "
                "indevida de benefício fiscal imunidade de ISS quando da emissão das NFS-e, visto\n"
                "que o sujeito passivo não se enquadra nas hipóteses de imunidade tributária "
                "instituídas pela Constituição Federal de 1988 em seu art. 150, inciso VI, e\n"
                "também não satisfaz as condições previstas no art. 85 da Lei Complementar nº "
                "40/2001 relativas à isenção do imposto.\n\n"
                "De acordo com o art. 4º da Lei Complementar Municipal "
                f"40/2001 e alterações, a alíquota para os serviços prestados é de {aliquota_str}."
            )

        elif rule_name == 'retencao_na_fonte_a_verificar':
            part_II = (
                "A constituição do crédito tributário decorreu de declaração "
                "indevida de retenção na fonte quando da emissão das NFS-e mencionadas, visto\n"
                "que não é hipótese legal de retenção ou substituição tributária, de acordo com "
                "a regra estabelecida nos arts. 8º e 8º-A da Lei Complementar Municipal 40/2001,consoante com a declaração\n"
                "indevida de benefício fiscal isenção, visto que também não são isentos do "
                "recolhimento de ISS, uma vez que não satisfazem as condições previstas no art.\n"
                "85 da Lei Complementar 40/2001.\n\n"
                "De acordo com o art. 4º da Lei Complementar Municipal "
                f"40/2001 e alterações, a alíquota para os serviços prestados é de {aliquota_str}."
            )
        else:
            part_II = (
                f"A constituição do crédito tributário decorreu da infração: {self.auto_data.get('motive', 'N/A')}.\n\n"
                "De acordo com o art. 4º da Lei Complementar Municipal "
                f"40/2001 e alterações, a alíquota para os serviços prestados é de {aliquota_str}."
            )
        
        return f"{part_I}\n\n{part_II}"