import requests
import os
import sys
import zipfile
import tempfile
import subprocess
import logging
from packaging import version
from app.constants import APP_VERSION, GITHUB_API_URL

class Updater:
    def check_for_updates(self):
        """
        Checks GitHub for a new version using a closed session to prevent QBasicTimer errors.
        """
        logger = logging.getLogger(__name__)
        # ✅ USE CONTEXT MANAGER (Fixes QBasicTimer::stop error)
        with requests.Session() as session:
            try:
                logger.info(f"Checking for updates... Current: {APP_VERSION}")
                
                # 1. Fetch Release Data
                response = session.get(GITHUB_API_URL, timeout=10) # Use 'session' here
                
                if response.status_code != 200:
                    logger.error(f"GitHub API returned status: {response.status_code}")
                    return False, "", ""
                    
                data = response.json()
                
                # 2. Extract Version Tag
                tag_name = data.get('tag_name')
                if not tag_name:
                    return False, "", ""

                latest_version_tag = tag_name.replace('v', '') 
                
                # 3. Find the .7z asset
                download_url = None
                assets = data.get('assets', [])
                for asset in assets:
                    if asset.get('name', '').endswith('.zip'):
                        download_url = asset.get('browser_download_url')
                        break
                
                if not download_url:
                    return False, "", ""

                # 4. Compare Versions
                if version.parse(latest_version_tag) > version.parse(APP_VERSION):
                    return True, download_url, latest_version_tag
                
                return False, "", ""
                
            except Exception as e:
                logger.exception("Update check failed.")
                return False, "", ""

    def download_and_install(self, download_url, status_callback=None):
        """
        Downloads the update, extracts it, creates a batch script to swap files,
        and restarts the application.
        """
        logger = logging.getLogger(__name__)
        try:
            msg = "⬇️ Baixando atualização..."
            logger.info(msg)
            if status_callback: status_callback(msg)
            
            # --- 1. Download Zip ---
            response = requests.get(download_url, stream=True)
            temp_dir = tempfile.mkdtemp()
            zip_path = os.path.join(temp_dir, "update.zip")
            
            with open(zip_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            msg = "📦 Extraindo arquivos..."
            logger.info(msg)
            if status_callback: status_callback(msg)

            # --- 2. Extract Zip ---
            extract_dir = os.path.join(temp_dir, "extracted")
            os.makedirs(extract_dir, exist_ok=True)
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)

            # --- 3. Prepare Paths ---
            current_exe = os.path.abspath(sys.argv[0]) 
            current_dir = os.path.dirname(current_exe)

            # Handle GitHub zip structure (sometimes it wraps files in a subfolder)
            source_dir = extract_dir
            items = os.listdir(extract_dir)
            
            # If the extracted folder contains ONLY one folder, assume the app is inside it
            if len(items) == 1 and os.path.isdir(os.path.join(extract_dir, items[0])):
                possible_subfolder = os.path.join(extract_dir, items[0])
                logger.info(f"Detected subfolder in zip: {possible_subfolder}")
                source_dir = possible_subfolder

            # --- 4. Create the Batch Script ---
            # This script waits, copies files (overwriting), cleans up, and restarts the app.
            bat_script = f"""
@echo off
echo Aguardando o fechamento da aplicacao...
timeout /t 3 /nobreak > NUL

echo Copiando novos arquivos...
robocopy "{source_dir}" "{current_dir}" /E /IS /MOVE

echo Limpando arquivos temporarios...
rmdir /s /q "{temp_dir}"

echo Reiniciando aplicacao...
start "" "{current_exe}"

echo Concluido.
del "%~f0"
            """
            
            bat_path = os.path.join(current_dir, "update_installer.bat")
            with open(bat_path, "w") as bat_file:
                bat_file.write(bat_script)

            msg = "🚀 Reiniciando para aplicar..."
            logger.info(msg)
            if status_callback: status_callback(msg)
            
            # --- 5. Execute Bat and Kill Self ---
            # Launch the batch file detached from this process
            subprocess.Popen([bat_path], shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE)
            
            # Exit immediately to release file locks
            return True
        
        except Exception as e:
            logger.exception("Critical error during update installation.")
            if status_callback: status_callback(f"❌ Erro: {e}")
            raise e