# --- app/ferramentas/decker.py ---

import time
import os
import glob
import re
import pyautogui
from DrissionPage import ChromiumPage, ChromiumOptions
from DrissionPage.common import Keys

def get_idd_and_files(imu, base_path):
    """
    Scans the base_path for a folder starting with {imu}_.
    Inside, looks for *_Comunicado.pdf and associated DAM.
    """
    search_pattern = os.path.join(base_path, f"{imu}_*")
    matching_dirs = glob.glob(search_pattern)
    
    if not matching_dirs: 
        return None, []
    
    company_dir = matching_dirs[0]
    
    # Find Comunicado to extract IDD Number
    pdf_pattern = os.path.join(company_dir, "*_Comunicado.pdf")
    comunicado_files = glob.glob(pdf_pattern)
    
    if not comunicado_files: 
        return None, []
    
    filename = os.path.basename(comunicado_files[0])
    idd_number = filename.split('_')[0]
    
    # Find DAM
    dam_pattern = os.path.join(company_dir, f"{idd_number}_DAM.pdf")
    dam_files = glob.glob(dam_pattern)
    
    found_files = [comunicado_files[0]]
    if dam_files: 
        found_files.append(dam_files[0])
        
    return idd_number, found_files

def process_company_email(page, cnpj, attachments, idd_number, manual_email=None):
    """
    Executes the browser actions to fill the form and send the email.
    """
    clean_cnpj = re.sub(r'\D', '', str(cnpj))
    
    title_text = f"Inscrição de Débitos Declarados Nº{idd_number}"
    
    # Body Text (Hardcoded as per original script)
    body_text = f"""Prezado(a),

Encaminha-se o COMUNICADO DO IDD Nº {idd_number} relacionado à sua situação fiscal. Detalhes específicos sobre o débito, incluindo data, descrição e valor, estão disponíveis nos documentos em anexo.

Em anexo, encontra-se o Documento de Arrecadação Municipal - DAM, para recolhimento do valor devido, com data de vencimento de 30 (trinta) dias após a emissão deste comunicado. Dentro do prazo de vencimento, o contribuinte poderá realizar o parcelamento amigável do débito, diretamente pelo Sistema Eletrônico de Gestão do Imposto sobre Serviços - ISSCuritiba (link:https://isscuritiba.curitiba.pr.gov.br/iss/default.aspx), conforme orientações disponíveis na página:https://www.curitiba.pr.gov.br/servicos/iss-parcelamento/311.

Ainda há tempo! Pague ou parcele o seu débito antes do vencimento. Evite que seu débito seja inscrito em dívida ativa, pois, além do incômodo de ver sua dívida sendo protestada ou cobrada judicialmente, terá de arcar, também, com honorários advocatícios e custas judiciais decorrentes."""

    try:
        page.get("https://adm-dec.curitiba.pr.gov.br/Mensagem/Cadastrar")
        
        if not page.ele("#Documento", timeout=5):
            return False, "Falha ao carregar formulário"

        # 1. Fill Title
        if page.ele("@name=Titulo"):
            page.ele("@name=Titulo").input(title_text)
        
        # 2. Fill CNPJ & Validate
        cnpj_input = page.ele("#Documento")
        if cnpj_input:
            cnpj_input.click()
            time.sleep(0.2)
            cnpj_input.clear()
            cnpj_input.input(clean_cnpj)
            time.sleep(0.5)
            cnpj_input.input(Keys.ENTER)
            
            # Validation Loop
            validation_success = False
            for _ in range(20): # 10s wait
                # Check for "Not Found" error
                error_span = page.ele("#divErroEmailAlerta")
                if error_span and "NÃO ENCONTRADO" in error_span.text:
                    if manual_email:
                         # Attempt to inject manual email
                         email_field = page.ele("#EmailAviso")
                         if email_field:
                             email_field.run_js(f"this.value = '{manual_email}';")
                             page.run_js("document.getElementById('divEmailAlerta').classList.remove('d-none');")
                             validation_success = True
                             break
                    else:
                        return False, "Email não encontrado no sistema"

                # Check for Success (Field Visible)
                parent_div = page.ele("#divEmailAlerta")
                if parent_div and "d-none" not in parent_div.attr("class"):
                    email_field = page.ele("#EmailAviso")
                    
                    if manual_email:
                        email_field.clear()
                        email_field.input(manual_email)
                    
                    validation_success = True
                    break
                
                time.sleep(0.5)

            if not validation_success:
                return False, "Timeout validando CNPJ/Email"

        # 3. Body Text
        editor_container = page.ele("#editorMensagem", timeout=3)
        if editor_container:
            # Try finding the editor content area
            text_area = editor_container.ele(".ql-editor", timeout=5) or editor_container.ele("css:div[contenteditable='true']")
            if text_area:
                text_area.click()
                js_text = body_text.replace('\n', '\\n') 
                text_area.run_js(f'this.innerText = `{js_text}`;') 

        # 4. Attachments
        for file_path in attachments:
            btn_add = page.ele("#btnAddAnexo_", timeout=3)
            if btn_add:
                btn_add.click()
                time.sleep(2.0)
                # Using pyautogui as DrissionPage file input handling can be tricky with OS dialogs sometimes, 
                # strictly following original script logic here.
                pyautogui.write(os.path.abspath(file_path))
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(2.0)

        # 5. Save & Send Sequence
        if page.ele("#btnSalvar"):
            page.ele("#btnSalvar").run_js("this.click()")
            time.sleep(4) 
        else:
            return False, "Botão Salvar não encontrado"

        # Popup 1
        if page.wait.ele_displayed("#btnEnviarMensagem", timeout=10):
            time.sleep(1) 
            page.ele("#btnEnviarMensagem").run_js("this.click()")
        else:
            return False, "Popup 'Enviar Mensagem' não apareceu"

        # Popup 2
        if page.wait.ele_displayed("#btnEnviar1", timeout=15):
            time.sleep(1)
            page.ele("#btnEnviar1").run_js("this.click()")
        else:
            return False, "Popup 'Enviar1' não apareceu"

        # Final Confirmation
        if page.wait.ele_displayed("#btnEnviar", timeout=10):
            time.sleep(1)
            page.ele("#btnEnviar").run_js("this.click()")
            return True, "Enviado com Sucesso"
        else:
            return False, "Popup Final não apareceu"

    except Exception as e:
        return False, f"Exceção: {str(e)}"

def run_decker_sender(tasks, base_directory, credentials, progress_callback):
    """
    Main worker function.
    tasks: list of dicts {'imu': ..., 'cnpj': ...}
    base_directory: folder where IDD folders are located
    credentials: dict {'user': ..., 'pass': ...}
    """
    progress_callback.emit("="*50)
    progress_callback.emit(" 📧 INICIANDO DECKER (Disparador de Emails)")
    progress_callback.emit(" ⚠️ Não mexa no mouse ou teclado durante o processo.")
    progress_callback.emit("="*50)

    try:
        co = ChromiumOptions()
        co.set_argument('--mute-audio')
        page = ChromiumPage(co)

        # LOGIN
        progress_callback.emit("🔑 Realizando Login...")
        page.get("https://adm-dec.curitiba.pr.gov.br/Login")
        
        if page.ele("#login_sup"):
            page.ele("#login_sup").input(credentials.get('user', ''))
            if page.ele("#senha_sup"):
                page.ele("#senha_sup").input(credentials.get('pass', ''))
            
            if page.ele("text:Entrar"):
                page.ele("text:Entrar").click()
            elif page.ele(".btn-primary"):
                page.ele(".btn-primary").click()
            
            time.sleep(3)
        
        # LOOP TASKS
        total = len(tasks)
        for i, task in enumerate(tasks):
            imu = task.get('imu')
            cnpj = task.get('cnpj')
            name = task.get('name', 'Empresa')
            
            progress_callback.emit(f"\n📨 [{i+1}/{total}] Processando: {name} (IMU: {imu})")
            
            # Locate Files
            idd_number, attachments = get_idd_and_files(imu, base_directory)
            
            if not idd_number:
                progress_callback.emit(f"   ❌ Pular: Arquivos PDF não encontrados na pasta.")
                continue

            progress_callback.emit(f"   📄 IDD: {idd_number} | Anexos: {len(attachments)}")
            
            # Send
            success, msg = process_company_email(page, cnpj, attachments, idd_number)
            
            if success:
                progress_callback.emit(f"   ✅ {msg}")
            else:
                progress_callback.emit(f"   ⛔ FALHA: {msg}")
            
            time.sleep(1)

        progress_callback.emit("\n🏁 Processo de envio finalizado.")

    except Exception as e:
        progress_callback.emit(f"❌ Erro Crítico no Navegador: {e}")