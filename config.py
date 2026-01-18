# --- FILE: app/config.py ---

import os
import json
from PySide6.QtCore import QSettings
from utils import resource_path

# --- Settings Keys ---
ALIQUOTAS_PATH = "paths/aliquotas"
TEMPLATE_INICIO_PATH = "paths/template_inicio"
TEMPLATE_RELATORIO_PATH = "paths/template_relatorio"
TEMPLATE_ENCERRAMENTO_DEC_PATH = "paths/template_encerramento_dec"
TEMPLATE_ENCERRAMENTO_AR_PATH = "paths/template_encerramento_ar"
OUTPUT_DIR = "paths/output_dir"
CUSTOM_GENERAL_TEXTS = "texts/general"
CUSTOM_AUTO_TEXTS = "texts/auto_specific"

# --- Default Fallback Values ---
DEFAULT_ALIQUOTAS = resource_path("atividades_aliquotas.xlsx")
DEFAULT_INICIO = resource_path("Anexo I_Modelo Termo de Início_v1.docx")
DEFAULT_RELATORIO = resource_path("template_receitas2.docx")
DEFAULT_ENCERRAMENTO_DEC = resource_path("Anexo IV_Modelo Termo de Encerramento_Receitas_DEC.docx")
DEFAULT_ENCERRAMENTO_AR = resource_path("Anexo III_Modelo Termo de Encerramento_Receitas_AR.docx") 
DEFAULT_OUTPUT = "output"

NEWS_SOURCE_URL = "https://raw.githubusercontent.com/Ostrensky/Caronte_FFRM/main/news.txt"

DEFAULT_GENERAL_TEXTS = {
    # --- ✅ NEW: IDD Specific Defaults ---
    "TITULO_DOCUMENTO_IDD": "Informação Fiscal - Operação IDD",
    "I_INTRO_IDD": "O presente procedimento teve como objeto a análise das informações constantes nas Notas Fiscais de Serviço Eletrônicas (NFS-e) emitidas pelo contribuinte no decorrer do ano-calendário de {{ ano }}, com vistas à formalização dos débitos de Imposto sobre Serviços de Qualquer Natureza (ISS) confessados e constituídos pelas NFS no sistema ISS-Curitiba, sem pagamento encontrado aos Cofres Municipais.",
    
    "I_INTRO": "O presente procedimento teve como objeto a análise das Notas Fiscais emitidas/declaradas, com respectivo cruzamento do Regime Tributário, Natureza da Operação, Deduções previstas na legislação, Alíquotas conforme os serviços descritos nos documentos, e outros indícios que possam prenunciar a emissão incorreta de notas fiscais que acarretem diferenças de Imposto sobre Serviços de Qualquer Natureza (ISS) a recolher para o Município de Curitiba.",
    "TITULO_DOCUMENTO_AUTO": "Relatório do Procedimento Fiscal da Operação Receitas (monitoramento)",
    
    "II_RECEITA_INTRO": "Da análise das receitas, foram verificadas as NFS-e emitidas no período fiscalizado e constataram-se divergências cujos créditos tributários de ISS foram lavrados mediante:",

    # --- Textos Condicionais ---
    "II_RECEITA_CONDITIONAL_IDDS": "Além disso, foram formalizados os créditos já confessados e constituídos, por meio de IDD {{ idds.numero }} = NFS-e de nº(s) {{ idds.nfs_e_numeros }} – período de {{ idds.periodo }}.",
    "II_RECEITA_CONDITIONAL_AVULSOS": "Ademais, identificou-se pagamentos avulsos que coincidem com as bases de cálculos das NFS-e de nº(s) {{ pagamentos_avulsos.nfs_e_numeros }} – período de {{ pagamentos_avulsos.periodo }}.",
    "II_RECEITA_CONDITIONAL_EXISTENTE": "Constatou-se Auto de Infração/ IDD nº {{ infracao_existente.numero }} já emitido para as NFS-e de nº(s) {{ infracao_existente.nfs_e_numeros }} – período de {{ infracao_existente.periodo }}.",
    "II_RECEITA_CONDITIONAL_DAS": "Verificou-se que houve pagamentos de DAS referente aos períodos de apuração {{ periodo_das }}. Os valores foram descontados do ISS devido.",
    "II_RECEITA_CONDITIONAL_AUTOS_COMPENSADOS": "Os autos de infração de nº(s) {{ achado_autos_compensados.lista_numeros }} foram gerados, porém seu valor foi totalmente compensado por pagamentos (DAM/DAS) efetuados, resultando em crédito tributário nulo.",
    "II_RECEITA_CONDITIONAL_INVOICES_COMPENSADAS": "As seguintes notas fiscais tiveram seu ISS apurado totalmente compensado por pagamentos efetuados e foram desconsideradas dos autos de infração correspondentes: {{ achado_invoices_compensadas.lista_formatada | join('; ') }}.",

    "III_CONDITIONAL_DECADENCIA": "Não foram constituídos os créditos tributários de ISS no período de {{ achado_decadencia_nao_autuado.periodo }}, referente às NFS-e de nº(s) {{ achado_decadencia_nao_autuado.nfs_numeros }}, por motivo de decadência.",
    "III_CONDITIONAL_PRESCRITO": "Não foram constituídos os créditos tributários de ISS no período de {{ achado_prescrito_nao_autuado.periodo }}, referente às NFS-e de nº(s) {{ achado_prescrito_nao_autuado.nfs_numeros }}, por motivo de prescrição.",
    "III_CONDITIONAL_SEM_PAGAMENTO": "Não foram identificados pagamentos de ISS no período de {{ achado_sem_pagamento.periodo }}. Não ocorreu decadência do ISS apurado nos períodos mencionados.",
    "III_CONDITIONAL_SEM_NOTAS": "Não foram identificadas notas emitidas nos meses {{ achado_sem_notas.periodo }}.",
    "III_CONDITIONAL_FORA_MUNICIPIO": "Não foram constituídos os créditos tributários de ISS no período de {{ achado_fora_municipio.periodo }}, referente às NFS-e de nº(s) {{ achado_fora_municipio.nfs_numeros }}, respectivamente, por terem sido declaradas com tributação para fora do município de Curitiba, observando-se que o serviço de construção civil, enquadrado no código de atividade 07.02, é devido no local de execução da obra.",

    "IV_CONDITIONAL_MULTA_DISPENSADA_IS": "Referente aos meses de janeiro e fevereiro, não foi aplicada a penalidade por descumprimento de dever instrumental, tendo em vista a Instrução de Serviço n.º 02/2022,determinando a dispensa da penalidade de acordo com o disposto na Instrução de Serviços.",
    "IV_CONDITIONAL_MULTA_DISPENSADA_SIMPLES": "Referente aos meses de janeiro e fevereiro de {{ multa_dispensada_simples.ano }}, conforme Instrução de Serviço nº 02/2022, não foi aplicada a penalidade por descumprimento de dever instrumental pela emissão incorreta dos das NFS-e referente ao regime tributário, pois a exclusão do regime Simples Nacional ocorreu em 01/2018.",
    "IV_CONDITIONAL_MULTA_SEM_INFRACAO": "Não foi constatada a emissão incorreta de notas no período em análise. Não há multa por descumprimento de dever instrumental.",
    "V_CONCLUSAO_INTRO_RECEITAS": "Em suma, os seguintes lançamentos foram realizados nesta Operação Receita:",
    "V_CONCLUSAO_FINAL": "Os lançamentos foram encaminhados para ciência do sujeito passivo, dando-se por encerrado o procedimento.",

    "AUDITOR_NOME": "Nome do Auditor Padrão",
    "AUDITOR_MATRICULA": "0000"
}

DEFAULT_AUTO_TEXTS = {
    "regime_incorreto": "declaração indevida do regime tributário informado nas NFS-e mencionadas, uma vez que o sujeito passivo não é optante pelo regime diferenciado no período autuado.",
    "aliquota_incorreta": (
        "A constituição do crédito tributário decorreu de diferença "
        "de alíquota, tendo em vista a declaração indevida nas NFS-e mencionadas no "
        "percentual de {{ declared_rate_str }}.\n\n"
        "De acordo com o art. 4º da Lei "
        "Complementar Municipal 40/2001 e alterações, a alíquota para os serviços "
        "prestados é de {{ correct_aliquota_str }} (art. 4º, inciso {{ inciso_roman }})."
    ),
    "idd_nao_pago": "constatação de falta de pagamento do ISS para as NFS-e mencionadas, que não foram pagas e não se enquadram em outras infrações.",
    "isencao_imunidade_indevida": "declaração indevida de Isenção/Imunidade nas NFS-e, não permitida pela legislação para o serviço prestado.",
    "deducao_indevida": "declaração indevida de dedução da base de cálculo nas NFS-e, não permitida pela legislação.",
    "natureza_operacao_incompativel": "declaração de Natureza da Operação incompatível com a tributação devida no município nas NFS-e.",
    "beneficio_fiscal_incorreto": "declaração indevida de benefício fiscal nas NFS-e.",
    "local_incidencia_incorreto": "declaração incorreta do local de incidência do ISS nas NFS-e.",
    "retencao_na_fonte_a_verificar": "verificar retenção na fonte declarada nas NFS-e (ISS Retido = Sim).",
    "IDD (Alíq. Declarada)": "ISS devido conforme alíquota declarada originalmente na NFS-e (parte de divisão de auto). Alíquota considerada: {{ correct_aliquota_str }}%.",
    "diferenca_aliquota": "Diferença entre a alíquota correta e a declarada na NFS-e (parte de divisão de auto). Alíquota correta considerada: {{ correct_aliquota_str }}%.",
    "DEFAULT_AUTO_FALLBACK": "{{ texto_simples }}"
}

# --- Functions --- 
def _get_setting(key, default_value):
    settings = QSettings("MyAuditApp", "AuditApp")
    value = settings.value(key)
    if isinstance(default_value, dict):
        if value:
            try:
                loaded_value = json.loads(value)
                if isinstance(loaded_value, dict):
                    default_copy = default_value.copy()
                    default_copy.update(loaded_value)
                    return default_copy
                else:
                    print(f"Warning: Stored value for '{key}' is not a dictionary. Using default.")
                    return default_value.copy()
            except json.JSONDecodeError:
                print(f"Warning: Could not decode JSON for setting '{key}'. Using default.")
                return default_value.copy()
        else:
            return default_value.copy()
    return value if value is not None else default_value

def _set_setting(key, value):
    settings = QSettings("MyAuditApp", "AuditApp")
    if isinstance(value, dict):
        try:
            settings.setValue(key, json.dumps(value, indent=4, ensure_ascii=False))
        except TypeError as e:
            print(f"Error saving setting '{key}' as JSON: {e}")
    else:
        settings.setValue(key, value)

def get_aliquotas_path(): return _get_setting(ALIQUOTAS_PATH, DEFAULT_ALIQUOTAS)
def get_template_inicio_path(): return _get_setting(TEMPLATE_INICIO_PATH, DEFAULT_INICIO)
def get_template_relatorio_path(): return _get_setting(TEMPLATE_RELATORIO_PATH, DEFAULT_RELATORIO)
def get_template_encerramento_dec_path(): return _get_setting(TEMPLATE_ENCERRAMENTO_DEC_PATH, DEFAULT_ENCERRAMENTO_DEC)
def get_template_encerramento_ar_path(): return _get_setting(TEMPLATE_ENCERRAMENTO_AR_PATH, DEFAULT_ENCERRAMENTO_AR)
def get_output_dir(): return _get_setting(OUTPUT_DIR, DEFAULT_OUTPUT)

def get_custom_general_texts():
    return _get_setting(CUSTOM_GENERAL_TEXTS, DEFAULT_GENERAL_TEXTS)
def set_custom_general_texts(texts_dict):
    _set_setting(CUSTOM_GENERAL_TEXTS, texts_dict)
def get_custom_auto_texts():
    return _get_setting(CUSTOM_AUTO_TEXTS, DEFAULT_AUTO_TEXTS)
def set_custom_auto_texts(texts_dict):
    _set_setting(CUSTOM_AUTO_TEXTS, texts_dict)

def set_aliquotas_path(path): _set_setting(ALIQUOTAS_PATH, path)
def set_template_inicio_path(path): _set_setting(TEMPLATE_INICIO_PATH, path)
def set_template_relatorio_path(path): _set_setting(TEMPLATE_RELATORIO_PATH, path)
def set_template_encerramento_dec_path(path): _set_setting(TEMPLATE_ENCERRAMENTO_DEC_PATH, path)
def set_template_encerramento_ar_path(path): _set_setting(TEMPLATE_ENCERRAMENTO_AR_PATH, path)
def set_output_dir(path): _set_setting(OUTPUT_DIR, path)