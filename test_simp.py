import os
import time
import random
import json
import shutil
from PySide6.QtCore import QThread
from DrissionPage import ChromiumPage, ChromiumOptions

# --- HUMAN SLEEP ---
def human_sleep(min_seconds=0.5, max_seconds=1.5):
    time.sleep(random.uniform(min_seconds, max_seconds))

def run_simples_downloader(tasks, progress_callback):
    """
    Função principal do worker usando DrissionPage.
    Fluxo: Preenche -> Captcha -> Mais Info -> Clica em Gerar PDF -> Monitora Temp -> Move para Final.
    """
    
    # --- HELPER: CHECK STOP ---
    def check_stop_flag():
        current_thread = QThread.currentThread()
        if hasattr(current_thread, 'check_stop') and current_thread.check_stop():
            return True
        return False

    # --- HELPER: LIMPAR PASTA TEMP ---
    def clear_temp_folder(folder):
        if os.path.exists(folder):
            for f in os.listdir(folder):
                fp = os.path.join(folder, f)
                try:
                    if os.path.isfile(fp): os.unlink(fp)
                except: pass

    progress_callback.emit("="*50)
    progress_callback.emit(" 🚀 INICIANDO DOWNLOADER (MODO ESTÁVEL)")
    progress_callback.emit(" 🔐 O navegador vai abrir. Resolva o Captcha quando solicitado.")
    progress_callback.emit("="*50)

    total_tasks = len(tasks)
    if total_tasks == 0:
        progress_callback.emit("⚠️ Lista de tarefas está vazia.")
        return

    # --- 1. SETUP - DEFINIR PASTA TEMPORÁRIA ÚNICA ---
    # Usamos uma pasta fixa para o Chrome salvar tudo lá.
    # Depois nós movemos para a pasta correta do cliente.
    temp_download_dir = os.path.join(os.getcwd(), "temp_pdf_cache")
    if not os.path.exists(temp_download_dir):
        os.makedirs(temp_download_dir)

    try:
        co = ChromiumOptions()
        co.set_argument('--mute-audio')
        
        # Configura a "Impressora" do Chrome para salvar PDF na pasta TEMP
        prefs = {
            'download.default_directory': temp_download_dir,
            'savefile.default_directory': temp_download_dir,
            'printing.print_preview_sticky_settings.appState': json.dumps({
                "recentDestinations": [{
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": ""
                }],
                "selectedDestinationId": "Save as PDF",
                "version": 2
            }),
            'plugins.always_open_pdf_externally': True 
        }
        
        for key, value in prefs.items():
            co.set_pref(key, value)
            
        co.set_argument('--kiosk-printing') # Imprime silenciosamente
        
        page = ChromiumPage(co)
        
    except Exception as e:
        progress_callback.emit(f"❌ Erro ao iniciar navegador: {e}")
        return

    # --- 2. LOOP TASKS ---
    for i, task in enumerate(tasks):
        
        if check_stop_flag():
            progress_callback.emit("🛑 Parada solicitada.")
            break 
        
        cnpj = task.get('cnpj')
        target_folder = task.get('folder')
        
        # Se não tiver pasta de destino, usa Downloads padrão
        if not target_folder:
            target_folder = os.path.join(os.getcwd(), "downloads")
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        # NÃO usamos page.set.download_path aqui para evitar o crash de Thread.
        # Deixamos o Chrome salvar no temp_download_dir configurado no início.

        progress_callback.emit(f"\n🔵 [{i+1}/{total_tasks}] Processando CNPJ: {cnpj}")

        try:
            # A. NAVEGAR
            if "simplesnacional" not in page.url:
                page.get("https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21")
            
            # B. INSERIR CNPJ
            if page.ele("#Cnpj"):
                page.ele("#Cnpj").clear()
                page.ele("#Cnpj").input(cnpj + '\n') 
                progress_callback.emit("   ✅ CNPJ inserido.")
            else:
                progress_callback.emit("   ❌ Campo CNPJ não encontrado. Recarregando...")
                page.get("https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21")
                time.sleep(2)
                if page.ele("#Cnpj"):
                    page.ele("#Cnpj").input(cnpj + '\n')
                else:
                    continue

            # C. ESPERAR CAPTCHA
            progress_callback.emit("   ⚠️  AGUARDANDO RESOLUÇÃO DO CAPTCHA...")
            
            found_success = False
            start_wait = time.time()
            while time.time() - start_wait < 120:
                if check_stop_flag(): break
                if page.ele("#btnMaisInfo", timeout=0.5):
                    found_success = True
                    break
                if page.ele("text:CNPJ inválido", timeout=0.1):
                    progress_callback.emit("   ❌ Site informou: CNPJ Inválido.")
                    break
                time.sleep(0.5)

            if not found_success:
                progress_callback.emit("   ❌ Tempo esgotado/Captcha falhou.")
                continue

            progress_callback.emit("   🔓 Acesso liberado!")
            time.sleep(0.5)

            # D. EXPANDIR MAIS INFO
            try:
                page.ele("#btnMaisInfo").click()
                time.sleep(1.0)
            except:
                pass

            # E. PREPARA PARA O DOWNLOAD
            # Limpa a pasta temporária para garantir que o arquivo novo seja único
            clear_temp_folder(temp_download_dir)
            
            # Define nomes finais
            expected_pdf_name = f"ConsultaOptantes_{cnpj}.pdf"
            folder_name = os.path.basename(target_folder)
            if '_' in folder_name:
                parts = folder_name.split('_', 1)
                if len(parts) > 1:
                    expected_pdf_name = f"ConsultaOptantes - {parts[1]}.pdf"
            
            final_full_path = os.path.join(target_folder, expected_pdf_name)

            # F. CLICAR EM 'GERAR PDF'
            progress_callback.emit("   🖱️ Clicando em 'Gerar PDF'...")
            
            pdf_saved = False
            
            if page.ele("#GerarPDF"):
                page.ele("#GerarPDF").click()
                
                # MONITORAMENTO NA PASTA TEMP
                progress_callback.emit("   ⏳ Aguardando geração do arquivo...")
                
                wait_download = 0
                found_temp_file = None
                
                while wait_download < 15:
                    time.sleep(1)
                    # Lista arquivos na pasta TEMP
                    temp_files = [f for f in os.listdir(temp_download_dir) if f.lower().endswith('.pdf')]
                    
                    if temp_files:
                        # Achou um PDF na pasta temp!
                        found_temp_file = os.path.join(temp_download_dir, temp_files[0])
                        break
                        
                    wait_download += 1
                
                if found_temp_file and os.path.exists(found_temp_file):
                    time.sleep(1) # Estabilizar escrita
                    
                    try:
                        # Se já existir na pasta final, remove
                        if os.path.exists(final_full_path):
                            os.remove(final_full_path)
                            
                        # Move da TEMP -> DESTINO FINAL
                        shutil.move(found_temp_file, final_full_path)
                        
                        progress_callback.emit(f"   ✅ PDF Salvo: {expected_pdf_name}")
                        pdf_saved = True
                    except Exception as move_err:
                        progress_callback.emit(f"   ⚠️ Erro ao mover arquivo: {move_err}")
                else:
                    progress_callback.emit("   ❌ O arquivo não apareceu na pasta temporária.")
            else:
                progress_callback.emit("   ⚠️ Botão 'Gerar PDF' não encontrado.")

            if not pdf_saved:
                progress_callback.emit("   ❌ FALHA: PDF não foi gerado.")

            # G. VOLTAR
            progress_callback.emit("   🔙 Voltando...")
            try:
                if page.ele("text:Voltar"):
                    page.ele("text:Voltar").click()
                elif page.ele('xpath://a[contains(@class, "btn-verde") and contains(text(), "Voltar")]'):
                     page.ele('xpath://a[contains(@class, "btn-verde") and contains(text(), "Voltar")]').click()
            except:
                pass

            time.sleep(1)

        except Exception as e:
            progress_callback.emit(f"   ❌ Erro Crítico: {e}")
            page.get("https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21")
            
    # --- END ---
    if not check_stop_flag():
        progress_callback.emit("\n🎉 TODOS OS PROCESSOS FORAM CONCLUÍDOS!")
    else:
        progress_callback.emit("\n🛑 Processamento interrompido.")

if __name__ == "__main__":
    class MockSignal:
        def emit(self, text):
            print(f"[GUI LOG]: {text}")

    test_cnpj = "00000000000191" 
    test_folder = os.path.join(os.getcwd(), "test_downloads")
    tasks_test = [{'cnpj': test_cnpj, 'folder': test_folder}]

    print(f"--- TESTE: {test_cnpj} ---")
    run_simples_downloader(tasks_test, MockSignal())